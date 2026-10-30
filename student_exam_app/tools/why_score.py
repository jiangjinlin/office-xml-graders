"""打印某评分器在素材上的得分点，便于逐个收紧。

用法: python tools/why_score.py excel 1
"""
import os
import re
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ROOT = os.path.dirname(BASE)
SUBJECT_DIR = {'excel': 'Excel', 'word': 'Word', 'ppt': 'Ppt'}
MATERIAL = {'excel': '2素材.xlsx', 'word': '2素材.docx', 'ppt': '2素材.pptx'}


def main():
    subject, qid = sys.argv[1], sys.argv[2]
    path = os.path.join(ROOT, '2022年准易：%s 2010 操作题--素材' % SUBJECT_DIR[subject],
                        qid, MATERIAL[subject])
    env = os.environ.copy()
    env['PYTHONIOENCODING'] = 'utf-8'
    r = subprocess.run([sys.executable, os.path.join(BASE, 'graders',
                        '%s_grader_%s.py' % (subject, qid)), path],
                       capture_output=True, timeout=120, env=env,
                       stdin=subprocess.DEVNULL, cwd=os.path.join(BASE, 'graders'))
    out = r.stdout.decode('utf-8', 'replace')
    for line in out.splitlines():
        s = line.strip()
        if not s:
            continue
        if ('OK' in s or '通过' in s or re.search(r'\(\+\d', s)
                or re.match(r'^\[(PASS|OK)\]', s) or re.match(r'^\[\d+\]', s)):
            print(line)
    m = re.search(r'得分率[:：]\s*([\d.]+)%|总(?:分|得分)[:：]\s*([\d.]+)\s*/\s*(\d+)', out)
    if m:
        if m.group(1):
            print('>>> 素材得分率 %s%%' % m.group(1))
        else:
            got, full = float(m.group(2)), float(m.group(3))
            print('>>> 素材得分 %s / %s' % (got, full))


if __name__ == '__main__':
    main()
