"""区分度检查：每个评分器跑 2素材（未作答），得分越低越好。

用法: python tools/discriminate.py [--max 60]
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
SUBJECT_DIR = {'excel': 'Excel', 'word': 'Word', 'ppt': 'PPT'}
MATERIAL = {'excel': '2素材.xlsx', 'word': '2素材.docx', 'ppt': '2素材.pptx'}
SCORE_RE = re.compile(r'得分率[:：]\s*([\d.]+)\s*%|总(?:分|得分)[:：]\s*([\d.]+)\s*/\s*(\d+)')


def discover():
    items = []
    for fname in sorted(os.listdir(GRADERS)):
        m = re.match(r'(excel|word|ppt|dw)_grader_(\d+)\.py$', fname)
        if not m:
            continue
        subject, qid = m.group(1), m.group(2)
        if subject == 'dw':
            material = os.path.join(BASE, 'materials', 'dw', qid, '2素材.zip')
        else:
            material = os.path.join(ROOT, '%s 2010 操作题题库' % SUBJECT_DIR[subject],
                                    qid, MATERIAL[subject])
        items.append({'subject': subject, 'qid': int(qid),
                      'grader': fname, 'material': material,
                      'exists': os.path.exists(material)})
    return items


def run_one(item, python):
    if not item['exists']:
        return {**item, 'score': None, 'error': '素材不存在'}
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    try:
        r = subprocess.run([python, os.path.join(GRADERS, item['grader']), item['material']],
                           capture_output=True, timeout=120, env=env,
                           stdin=subprocess.DEVNULL, cwd=GRADERS)
    except subprocess.TimeoutExpired:
        return {**item, 'score': None, 'error': '超时'}
    out = r.stdout.decode('utf-8', 'replace')
    score = None
    m = SCORE_RE.search(out)
    if m:
        if m.group(1):
            score = float(m.group(1))
        else:
            got, full = float(m.group(2)), float(m.group(3))
            score = round(got / full * 100, 1) if full else None
    return {**item, 'score': score, 'error': '' if score is not None else out[-200:]}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--max', type=float, default=60.0, help='超过该阈值判为过松')
    ap.add_argument('--json')
    args = ap.parse_args()

    python = sys.executable
    results = [run_one(it, python) for it in discover()]
    loose = 0
    print('{:<8}{:<5}{:<24}{:>8}  {}'.format('科目', '题号', '评分器', '素材得分', '状态'))
    print('-' * 62)
    for r in results:
        score = '{:.1f}'.format(r['score']) if r['score'] is not None else '--'
        flag = ''
        if r['score'] is None:
            flag = r['error'][:30]
        elif r['score'] > args.max:
            flag = '过松!'
            loose += 1
        print('{:<8}{:<5}{:<24}{:>8}  {}'.format(r['subject'], r['qid'], r['grader'], score, flag))
    print('-' * 62)
    print('共 %d 个，素材得分高于 %.0f 的 %d 个' % (len(results), args.max, loose))
    if args.json:
        with open(args.json, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
    return 0


if __name__ == '__main__':
    sys.exit(main())
