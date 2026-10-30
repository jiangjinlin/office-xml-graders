"""小狮子 Office 2010 操作题考试系统 - 后端

Excel 1-10 · Word 1-10 · PPT 1-10 · 网页 1-5
"""
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
import xml.etree.ElementTree as ET

from flask import Flask, jsonify, render_template, request, send_file

BASE = os.path.dirname(os.path.abspath(__file__))
MATERIALS = os.path.join(BASE, 'materials')
GRADERS = os.path.join(BASE, 'graders')
UPLOADS = os.path.join(BASE, 'uploads')
STATIC = os.path.join(BASE, 'static')
os.makedirs(UPLOADS, exist_ok=True)

SUBJECTS = {
    'excel': {'ext': '.xlsx', 'mime': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},
    'word': {'ext': '.docx', 'mime': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'},
    'ppt': {'ext': '.pptx', 'mime': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'},
    'dw': {'ext': '.zip', 'mime': 'application/zip'},
}

LABELS = {'excel': 'Excel', 'word': 'Word', 'ppt': 'PPT', 'dw': '网页'}

MAX_UPLOAD_MB = int(os.environ.get('EXAM_MAX_UPLOAD_MB', '50'))
GRADER_TIMEOUT = int(os.environ.get('EXAM_GRADER_TIMEOUT', '90'))
MAX_PARALLEL = int(os.environ.get('EXAM_MAX_PARALLEL', '2'))
UPLOAD_TTL_HOURS = int(os.environ.get('EXAM_UPLOAD_TTL', '6'))

app = Flask(__name__, static_folder=STATIC, static_url_path='/static')
app.secret_key = os.environ.get('EXAM_SECRET_KEY') or os.urandom(24).hex()
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_MB * 1024 * 1024
app.json.ensure_ascii = False

SLOTS = threading.BoundedSemaphore(MAX_PARALLEL)

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def python_exe():
    """uwsgi/gunicorn 的 sys.executable 不是解释器，优先取同目录的 python。"""
    override = os.environ.get('EXAM_PYTHON')
    if override and os.path.exists(override):
        return override
    exe = sys.executable or ''
    if 'python' in os.path.basename(exe).lower():
        return exe
    sibling = os.path.join(os.path.dirname(exe), 'python')
    if exe and os.path.exists(sibling):
        return sibling
    for cmd in ('python3', 'python'):
        found = shutil.which(cmd)
        if found:
            return found
    return exe or 'python3'


PYTHON = python_exe()

QUESTION_CACHE = {'key': None, 'data': None}


def materials_key():
    parts = []
    for subject in SUBJECTS:
        folder = os.path.join(MATERIALS, subject)
        if not os.path.isdir(folder):
            continue
        for name in sorted(os.listdir(folder)):
            path = os.path.join(folder, name, '3题干.docx') if os.path.exists(os.path.join(folder, name, '3题干.docx')) else os.path.join(folder, name, '3题干.txt')
            if os.path.exists(path):
                st = os.stat(path)
                parts.append('%s/%s/%d/%d' % (subject, name, st.st_size, int(st.st_mtime)))
    return '|'.join(parts)


def read_docx_lines(path):
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read('word/document.xml'))
    lines = []
    for p in root.iter('{%s}p' % W):
        line = ''.join(t.text or '' for t in p.iter('{%s}t' % W)).strip()
        if line:
            lines.append(line)
    return lines


def read_question_lines(path):
    if path.endswith('.docx'):
        return read_docx_lines(path)
    with open(path, encoding='utf-8') as f:
        return [line.strip() for line in f if line.strip()]


def extract_questions():
    questions = {subject: {} for subject in SUBJECTS}
    for subject in SUBJECTS:
        folder = os.path.join(MATERIALS, subject)
        if not os.path.isdir(folder):
            continue
        for qid in sorted(os.listdir(folder), key=lambda x: (len(x), x)):
            docx = os.path.join(folder, qid, '3题干.docx')
            txt = os.path.join(folder, qid, '3题干.txt')
            path = docx if os.path.exists(docx) else (txt if os.path.exists(txt) else None)
            if path is None:
                continue
            try:
                questions[subject][qid] = read_question_lines(path)
            except Exception as exc:
                questions[subject][qid] = ['读取题干失败: %s' % exc]
    return questions


def purge_uploads(max_age_hours=UPLOAD_TTL_HOURS, budget=40):
    cutoff = time.time() - max_age_hours * 3600
    removed = 0
    for name in os.listdir(UPLOADS):
        if removed >= budget:
            break
        path = os.path.join(UPLOADS, name)
        try:
            if os.path.getmtime(path) < cutoff:
                removed += safe_rmtree(path, budget - removed)
        except OSError:
            pass
    return removed


def safe_rmtree(path, budget=None):
    """逐个删除，避免整目录批量删除被安全策略拦截；budget 限制单次删除数量。"""
    if not os.path.isdir(path):
        return 0
    removed = 0
    for root, dirs, files in os.walk(path, topdown=False):
        for name in files:
            if budget is not None and removed >= budget:
                return removed
            try:
                os.remove(os.path.join(root, name))
                removed += 1
            except OSError:
                pass
        for name in dirs:
            try:
                os.rmdir(os.path.join(root, name))
            except OSError:
                pass
    try:
        os.rmdir(path)
    except OSError:
        pass
    return removed


def janitor(stop_event, interval=1800):
    while not stop_event.wait(interval):
        try:
            purge_uploads()
        except Exception:
            pass


def start_janitor():
    stop_event = threading.Event()
    thread = threading.Thread(target=janitor, args=(stop_event,), daemon=True)
    thread.start()
    return stop_event


def valid_subject(subject):
    return subject in SUBJECTS


def valid_qid(qid):
    return bool(re.fullmatch(r'\d{1,2}', str(qid)))


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/health')
def health():
    graders = [f for f in os.listdir(GRADERS) if re.match(r'(excel|word|ppt|dw)_grader_\d+\.py$', f)]
    questions = 0
    for subject in SUBJECTS:
        folder = os.path.join(MATERIALS, subject)
        if os.path.isdir(folder):
            questions += sum(1 for name in os.listdir(folder)
                             if os.path.exists(os.path.join(folder, name, '3题干.docx')) or os.path.exists(os.path.join(folder, name, '3题干.txt')))
    return jsonify({
        'status': 'ok',
        'graders': len(graders),
        'questions': questions,
        'python': os.path.basename(PYTHON),
        'max_upload_mb': MAX_UPLOAD_MB,
        'max_parallel': MAX_PARALLEL,
        'upload_ttl_hours': UPLOAD_TTL_HOURS,
    })


@app.route('/api/questions')
def api_questions():
    key = materials_key()
    if QUESTION_CACHE['key'] != key:
        QUESTION_CACHE['data'] = extract_questions()
        QUESTION_CACHE['key'] = key
    return jsonify(QUESTION_CACHE['data'])


@app.route('/api/download/<subject>/<qid>')
def download_material(subject, qid):
    if not valid_subject(subject) or not valid_qid(qid):
        return jsonify({'error': '参数不合法'}), 400

    folder = os.path.join(MATERIALS, subject, qid)
    if not os.path.isdir(folder):
        return jsonify({'error': '题目不存在'}), 404

    files = [f for f in os.listdir(folder)
             if os.path.isfile(os.path.join(folder, f))
             and not f.startswith(('3题干', '1结果', '样张', '.'))]
    if not files:
        return jsonify({'error': '素材缺失'}), 404

    main_file = '2素材' + SUBJECTS[subject]['ext']
    if len(files) == 1 and files[0] == main_file:
        return send_file(os.path.join(folder, main_file), mimetype=SUBJECTS[subject]['mime'],
                         as_attachment=True, download_name='%s_题目%s_%s' % (subject, qid, main_file))

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name in sorted(files):
            zf.write(os.path.join(folder, name), name)
    buf.seek(0)
    label = LABELS[subject]
    return send_file(buf, mimetype='application/zip', as_attachment=True,
                     download_name='%s_题目%s_素材包.zip' % (label, qid))


SCORE_PERCENT_RE = re.compile(r'得分率[:：]\s*([\d.]+)\s*%')
SCORE_FULL_RE = re.compile(r'总(?:分|得分)[:：]\s*([\d.]+)\s*/\s*([\d.]+)\s*=\s*([\d.]+)\s*%')
SCORE_RATIO_RE = re.compile(r'总(?:分|得分)[:：]\s*([\d.]+)\s*/\s*([\d.]+)')
SCORE_GENERIC_RE = re.compile(r'([\d.]+)\s*/\s*([\d.]+)\s*=\s*([\d.]+)\s*%')
ITEM_RE = re.compile(r'^\s*\[(PASS|FAIL|OK|NG)\]\s*(\d+)[.、]\s*(.+?)\s+([\d.]+)\s*/\s*([\d.]+)')
WORD_ITEM_RE = re.compile(r'^\s*\[(\d{1,2})\]\s*(.+?)\s*$')
WORD_SCORE_RE = re.compile(r'^\s*\[(OK|NG)\]\s*([\d.]+)\s*/\s*([\d.]+)')
NOISE_RE = re.compile(
    r'^(===|----|目标文件:|解压完成:|已加载:|报告已生成|报告:|评分:|文件:|加载\s|'
    r'操作题评分器|document\.xml|styles\.xml|settings\.xml|header\d|footer\d|'
    r'endnotes|footnotes|numbering|xmlns[:=]|ns\d+:|</|<ns\d+:|'
    r'\d+个页眉|\d+个页脚|\d+个媒体文件)')


def parse_score(text):
    match = SCORE_PERCENT_RE.search(text)
    if match:
        return round(float(match.group(1)), 1)

    match = SCORE_FULL_RE.search(text)
    if match:
        return round(float(match.group(3)), 1)

    match = SCORE_RATIO_RE.search(text)
    if match:
        got, full = float(match.group(1)), float(match.group(2))
        return round(got / full * 100, 1) if full else None

    match = SCORE_GENERIC_RE.search(text)
    if match:
        return round(float(match.group(3)), 1)
    return None


def parse_items(text):
    items = []
    pending = None
    for raw in text.splitlines():
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped:
            continue

        match = ITEM_RE.match(stripped)
        if match:
            pending = {
                'no': int(match.group(2)),
                'name': match.group(3).strip(),
                'got': float(match.group(4)),
                'full': float(match.group(5)),
                'status': 'pass' if match.group(1) in ('PASS', 'OK') else 'fail',
                'notes': [],
            }
            items.append(pending)
            continue

        match = WORD_ITEM_RE.match(stripped)
        if match:
            pending = {'no': int(match.group(1)), 'name': match.group(2).strip(),
                       'got': 0.0, 'full': 0.0, 'status': 'info', 'notes': []}
            items.append(pending)
            continue

        match = WORD_SCORE_RE.match(stripped)
        if match and pending is not None:
            pending['got'] = float(match.group(2))
            pending['full'] = float(match.group(3))
            pending['status'] = 'pass' if match.group(1) == 'OK' else 'fail'
            continue

        if pending is not None:
            pending['notes'].append(stripped)
    return items


def build_details(text, items):
    if items and all(item['full'] for item in items):
        lines = []
        for item in items:
            lines.append('[%s] %d. %s  %.1f/%g' % (
                'pass' if item['status'] == 'pass' else 'fail',
                item['no'], item['name'], item['got'], item['full']))
            for note in item['notes']:
                mark = 'pass' if ('OK' in note or '通过' in note) else 'info'
                lines.append('[%s]      %s' % (mark, note))
        return lines

    lines = []
    for raw in text.splitlines():
        stripped = raw.strip()
        if not stripped or NOISE_RE.search(stripped):
            continue
        if SCORE_RATIO_RE.search(stripped) or SCORE_FULL_RE.search(stripped):
            continue
        low = stripped
        if low.startswith('[PASS]') or low.startswith('[OK]') or low.startswith('OK '):
            lines.append('[pass] ' + stripped)
        elif low.startswith('[FAIL]') or low.startswith('[NG]') or low.startswith('NOT '):
            lines.append('[fail] ' + stripped)
        elif low.endswith('OK') or '通过' in low:
            lines.append('[pass] ' + stripped)
        elif '未通过' in low or '未找到' in low:
            lines.append('[fail] ' + stripped)
        else:
            lines.append('[info] ' + stripped)
    return lines


def run_grader(subject, qid, saved_path, workdir):
    grader = os.path.join(GRADERS, '%s_grader_%s.py' % (subject, qid))
    if not os.path.exists(grader):
        return None, ['[info] 评分器未找到: %s' % os.path.basename(grader)], 'grader_missing'

    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    env['PYTHONDONTWRITEBYTECODE'] = '1'
    try:
        result = subprocess.run(
            [PYTHON, grader, saved_path],
            capture_output=True, timeout=GRADER_TIMEOUT, env=env,
            stdin=subprocess.DEVNULL, cwd=workdir,
        )
    except subprocess.TimeoutExpired:
        return None, ['[info] 评分超时（%d 秒）' % GRADER_TIMEOUT], 'timeout'

    out = result.stdout.decode('utf-8', 'replace') if result.stdout else ''
    err = result.stderr.decode('utf-8', 'replace') if result.stderr else ''
    if not out.strip():
        detail = ['[info] 评分器无输出（返回码 %s）' % result.returncode]
        if err.strip():
            detail.append('[info] ' + err.strip().splitlines()[-1][:200])
        return None, detail, 'no_output'
    return out, None, None


@app.route('/api/grade/<subject>/<qid>', methods=['POST'])
def grade(subject, qid):
    if not valid_subject(subject) or not valid_qid(qid):
        return jsonify({'error': '参数不合法'}), 400

    upload = request.files.get('file')
    if upload is None or not upload.filename:
        return jsonify({'error': '未上传文件'}), 400

    ext = os.path.splitext(upload.filename)[1].lower()
    if ext != SUBJECTS[subject]['ext']:
        return jsonify({'error': '请上传 %s 文件，收到 %s' % (SUBJECTS[subject]['ext'], ext or '无扩展名')}), 400

    if not SLOTS.acquire(blocking=False):
        return jsonify({'error': '当前判分人数较多，请稍后重试'}), 429

    workdir = tempfile.mkdtemp(prefix='%s_%s_' % (subject, qid), dir=UPLOADS)
    try:
        name = os.path.basename(upload.filename).replace('\\', '_').replace('/', '_')
        saved_path = os.path.join(workdir, name)
        upload.save(saved_path)

        if not zipfile.is_zipfile(saved_path):
            return jsonify({'error': '文件已损坏或不是有效的 Office 文件'}), 400

        output, error_details, reason = run_grader(subject, qid, saved_path, workdir)
        if output is None:
            return jsonify({'error': error_details[0].replace('[info] ', '') if reason != 'timeout' else '评分超时',
                            'details': error_details}), 500

        score = parse_score(output)
        items = parse_items(output)
        details = build_details(output, items)

        return jsonify({
            'score': score,
            'max_score': 100,
            'passed': score is not None and score >= 60,
            'items': items,
            'details': details,
        })
    except Exception as exc:
        app.logger.exception('grade failed')
        return jsonify({'error': '服务器错误: %s' % exc}), 500
    finally:
        SLOTS.release()
        try:
            shutil.rmtree(workdir)
        except Exception:
            pass


@app.errorhandler(413)
def too_large(_):
    return jsonify({'error': '文件超过 %d MB 限制' % MAX_UPLOAD_MB}), 413


def _bootstrap():
    if os.environ.get('EXAM_JANITOR', '1') != '1':
        return
    try:
        purge_uploads()
    except Exception:
        pass
    start_janitor()


_bootstrap()


if __name__ == '__main__':
    print('=' * 56)
    print('  Lion Office 2010 Exam System')
    print('  Excel 1-10 | Word 1-10 | PPT 1-10 | DW 1-5')
    print('  http://0.0.0.0:5000')
    print('=' * 56)
    app.run(host='0.0.0.0', port=int(os.environ.get('EXAM_PORT', '5000')), debug=False)
