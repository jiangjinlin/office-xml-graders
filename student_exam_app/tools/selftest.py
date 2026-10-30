"""批量回归：把每个评分器跑一遍标准答案，输出得分表。

用法：
    python tools/selftest.py            # 跑 app 内全部评分器
    python tools/selftest.py --json out.json
"""
import argparse
import json
import os
import re
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ROOT = os.path.dirname(BASE)
GRADERS = os.path.join(BASE, 'graders')
MATERIALS = os.path.join(ROOT, '{} 2010 操作题题库')
SUBJECT_DIR = {
    'excel': 'Excel',
    'word': 'Word',
    'ppt': 'PPT',
}
ANSWER = {'excel': '1结果.xlsx', 'word': '1结果.docx', 'ppt': '1结果.pptx'}
SCORE_RE = re.compile(r'总(?:分|得分)[:：]\s*([\d.]+)\s*/\s*(\d+)')


def discover():
    items = []
    for fname in sorted(os.listdir(GRADERS)):
        m = re.match(r'(excel|word|ppt|dw)_grader_(\d+)\.py$', fname)
        if not m:
            continue
        subject, qid = m.group(1), m.group(2)
        if subject == 'dw':
            answer = os.path.join(BASE, 'materials', 'dw', qid, '1结果.zip')
        else:
            answer = os.path.join(MATERIALS.format(SUBJECT_DIR[subject]), qid, ANSWER[subject])
        items.append({
            'subject': subject, 'qid': int(qid), 'grader': fname,
            'grader_path': os.path.join(GRADERS, fname),
            'answer': answer, 'answer_exists': os.path.exists(answer),
        })
    return items


def run_one(item, python):
    if not item['answer_exists']:
        return {**item, 'ok': False, 'score': None, 'error': '答案文件不存在'}
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    try:
        r = subprocess.run(
            [python, item['grader_path'], item['answer']],
            capture_output=True, timeout=120, env=env,
            stdin=subprocess.DEVNULL, cwd=GRADERS,
        )
    except subprocess.TimeoutExpired:
        return {**item, 'ok': False, 'score': None, 'error': '超时'}

    out = r.stdout.decode('utf-8', errors='replace')
    err = r.stderr.decode('utf-8', errors='replace')
    m = SCORE_RE.search(out)
    score = None
    if m:
        got, full = float(m.group(1)), float(m.group(2))
        score = round(got / full * 100, 1) if full else None
    return {
        **item, 'ok': score is not None, 'score': score,
        'raw': '{} / {}'.format(m.group(1), m.group(2)) if m else '',
        'returncode': r.returncode,
        'error': '' if score is not None else (err.strip()[-400:] or '未解析到总分'),
        'pass_count': out.count('[PASS]'),
        'fail_count': out.count('[FAIL]'),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--json')
    ap.add_argument('--python', default=sys.executable)
    args = ap.parse_args()

    results = [run_one(it, args.python) for it in discover()]
    if not results:
        print('未找到任何评分器')
        return 1

    print('{:<8}{:<5}{:<24}{:>7}{:>7}{:>7}  {}'.format('科目', '题号', '评分器', '得分', 'PASS', 'FAIL', '状态'))
    print('-' * 78)
    bad = 0
    for r in results:
        score = '{:.1f}'.format(r['score']) if r['score'] is not None else '--'
        flag = 'OK' if r['score'] == 100 else 'CHECK'
        if r['score'] != 100:
            bad += 1
        print('{:<8}{:<5}{:<24}{:>7}{:>7}{:>7}  {}'.format(
            r['subject'], r['qid'], r['grader'], score,
            r.get('pass_count', '-'), r.get('fail_count', '-'), flag))
        if r['error']:
            print('        ! {}'.format(r['error'].replace('\n', ' ')[:160]))

    print('-' * 78)
    print('共 {} 个评分器，满分 {} 个，需检查 {} 个'.format(
        len(results), len(results) - bad, bad))

    if args.json:
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print('明细已写入 {}'.format(args.json))
    return 0 if bad == 0 else 2


if __name__ == '__main__':
    sys.exit(main())
