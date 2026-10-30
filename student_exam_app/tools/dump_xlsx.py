"""打印 xlsx 答案文件的关键结构，用于对照题干写评分器。

用法: python tools/dump_xlsx.py <文件> [工作表名]
"""
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, 'graders'))

from _core import ExcelBook  # noqa: E402
from _core.xlsx import S  # noqa: E402


def style_brief(fmt):
    font = fmt.get('font', {})
    align = fmt.get('alignment', {})
    fill = fmt.get('fill', {})
    border = fmt.get('border', {})
    sides = [k for k, v in border.items() if v.get('style')]
    parts = [
        'font=%s/%s%s' % (font.get('name'), font.get('size'), '/B' if font.get('bold') else ''),
        'color=%s' % (font.get('color', {}).get('rgb') or font.get('color', {}).get('theme') or ''),
        'fill=%s' % (fill.get('pattern') or ''),
        'fg=%s' % (fill.get('fg', {}).get('rgb') or fill.get('fg', {}).get('theme') or ''),
        'h=%s' % align.get('horizontal', ''),
        'v=%s' % align.get('vertical', ''),
        'fmt=%r' % fmt.get('numfmt', ''),
    ]
    if sides:
        parts.append('border=%s' % ','.join('%s:%s' % (k, border[k]['style']) for k in sides))
    return ' '.join(p for p in parts if not p.endswith('='))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    path = sys.argv[1]
    only = sys.argv[2] if len(sys.argv) > 2 else ''

    with ExcelBook(path) as book:
        print('sheets:', [(s.name, s.part) for s in book.sheets])
        print('defined names:', [n.get('name') + '=' + (n.text or '')
                                 for n in book.root(book.workbook_part).iter('{%s}definedName' % S)])
        for sheet in book.sheets:
            if only and only != sheet.name:
                continue
            cells = sheet.cells()
            rows = sheet.rows()
            cols = sorted({c for c in cells})
            print('\n== %s (%s) 单元格=%d 行=%d' % (sheet.name, sheet.part, len(cells), len(rows)))
            print('  merges:', sheet.merged()[:15])
            tab = sheet.root.find('{%s}sheetPr/{%s}tabColor' % (S, S))
            if tab is not None:
                print('  tabColor:', tab.attrib)
            print('  cols:', [(c.get('min'), c.get('max'), c.get('width'), c.get('bestFit'), c.get('customWidth'))
                              for c in sheet.root.iter('{%s}col' % S)][:8])

            sample = ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3', 'J3']
            for ref in sample:
                cell = book.cell(ref, sheet)
                if cell is None:
                    continue
                fmt = book.styles.get(cell.style_index)
                print('  %-3s = %-24r f=%-40r s=%d %s' % (
                    ref, str(cell.value)[:24], cell.formula[:40], cell.style_index, style_brief(fmt)))

            formulas = {}
            for ref, node in cells.items():
                f = node.find('{%s}f' % S)
                if f is not None and f.text:
                    key = f.text.split('(')[0] if '(' in f.text else f.text[:12]
                    formulas.setdefault(key, []).append((ref, f.text))
            for key, items in formulas.items():
                print('  公式 %-10s x%d  例: %s = %s' % (key, len(items), items[0][0], items[0][1]))

            sorts = list(sheet.root.iter('{%s}sortState' % S))
            for st in sorts:
                print('  sortState ref=%s' % st.get('ref'),
                      [(sc.get('ref'), sc.get('descending')) for sc in st.iter('{%s}sortCondition' % S)])
            outline = {}
            for row in rows:
                lvl = row.get('outlineLevel')
                if lvl:
                    outline[lvl] = outline.get(lvl, 0) + 1
            if outline:
                print('  outlineLevel:', outline)
            hidden = sum(1 for r in rows if r.get('hidden') == '1')
            if hidden:
                print('  hidden rows:', hidden)
            for cf in sheet.root.iter('{%s}conditionalFormatting' % S):
                for rule in cf.iter('{%s}cfRule' % S):
                    print('  CF sqref=%s type=%s op=%s formula=%s dxf=%s' % (
                        cf.get('sqref'), rule.get('type'), rule.get('operator'),
                        rule.findtext('{%s}formula' % S), rule.get('dxfId')))
            ps = sheet.root.find('{%s}pageSetup' % S)
            po = sheet.root.find('{%s}printOptions' % S)
            hf = sheet.root.find('{%s}headerFooter' % S)
            if ps is not None or po is not None or hf is not None:
                print('  pageSetup:', ps.attrib if ps is not None else {},
                      'printOptions:', po.attrib if po is not None else {})
                if hf is not None:
                    print('  headerFooter:', [(c.tag.split('}')[-1], c.text) for c in hf])

        charts = book.charts()
        for part in charts:
            xml = book.read(part)
            print('\n== chart %s' % part)
            for key in ('bar3DChart', 'barChart', 'lineChart', 'pieChart', 'scatterChart', 'areaChart'):
                if '<c:%s' % key in xml or ':%s>' % key in xml:
                    print('  type:', key)
            import re
            print('  titles:', re.findall(r'<a:t>([^<]{1,40})</a:t>', xml)[:6])
            m = re.search(r'<c:legendPos val="(\w+)"', xml)
            print('  legendPos:', m.group(1) if m else '')
            m = re.search(r'<c:title>.*?<a:defRPr[^>]*sz="(\d+)"', xml, re.S)
            if m:
                print('  title size:', int(m.group(1)) / 100)
            print('  font refs:', re.findall(r'typeface="([^"]+)"', xml)[:4])
            print('  series refs:', re.findall(r'<c:f>([^<]+)</c:f>', xml)[:6])

        for part in book.tables():
            xml = book.read(part)
            import re
            print('\n== table %s' % part, re.findall(r'<tableStyleInfo name="([^"]+)"', xml))

        pivots = book.pivot_tables()
        for part in pivots:
            import re
            print('\n== pivot %s' % part)
            fields = []
            for pc in book.parts(book._base + 'pivotCache/', exclude=('_rels/',)):
                fields = re.findall(r'<cacheField name="([^"]+)"', book.read(pc))
                break
            print('  cacheFields:', fields)
            xml = book.read(part)
            print('  rowFields:', re.findall(r'<pivotField axis="axisRow"[^>]*>.*?<field x="(\d+)"', xml)[:5],
                  'colFields:', re.findall(r'<field x="(\d+)"', xml)[:8])
            print('  dataFields:', re.findall(r'<dataField name="([^"]+)"[^>]*subtotal="(\w+)"', xml))
    return 0


if __name__ == '__main__':
    sys.exit(main())
