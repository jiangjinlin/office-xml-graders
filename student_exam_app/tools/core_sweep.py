"""核心库批量装载检查：所有标准答案文件都必须能无异常解析。"""
import os
import sys
import traceback

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, 'graders'))

from _core import ExcelBook, WordDoc, PptDeck  # noqa: E402

ROOT = os.path.dirname(BASE)
CASES = [
    ('Excel', ExcelBook, '1结果.xlsx', lambda d: '%d表/%d图/%d表对象' % (
        len(d.sheets), len(d.charts()), len(d.tables()))),
    ('Word', WordDoc, '1结果.docx', lambda d: '%d段/%d表/%d节/%d页眉' % (
        len(d.paragraphs), len(d.tables), len(d.sections), len(d.headers))),
    ('Ppt', PptDeck, '1结果.pptx', lambda d: '%d页/主题%s' % (len(d.slides), d.theme.get('name'))),
]


def main():
    bad = 0
    for label, cls, answer, summary in CASES:
        for qid in range(1, 11):
            path = os.path.join(ROOT, '2022年准易：%s 2010 操作题--素材' % label, str(qid), answer)
            if not os.path.exists(path):
                print('%-6s %-3d 缺少答案文件' % (label, qid))
                continue
            try:
                with cls(path) as doc:
                    print('%-6s %-3d %s' % (label, qid, summary(doc)))
            except Exception as exc:
                bad += 1
                print('%-6s %-3d 解析失败: %s' % (label, qid, exc))
                traceback.print_exc(limit=2)
    print('失败 %d 个' % bad)
    return 1 if bad else 0


if __name__ == '__main__':
    sys.exit(main())
