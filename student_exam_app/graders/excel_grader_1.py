"""
Excel 2010 评分器 - 文件夹1 (v3 精确XML检测)
"""
import zipfile, os, sys, io, re, xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import ExcelBook

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def child(e, t):
    for c in e:
        if c.tag == t: return c
    return None

class ExcelGrader1:
    def __init__(self, path):
        self.path = path; self.zf = None
        self.ss = []           # sharedStrings
        self.sheets = {}       # {name: root}
        self.sheet_names = []  # [(name, sheetId)]
        self.all_xml = ''
        self.styles_xml = ''
        self.wb_xml = ''
        self.namelist = []

    def load(self):
        self.book = ExcelBook(self.path)
        self.zf = self.book.zf
        self.namelist = self.book.names
        self.ss = self.book.strings
        self.styles_xml = self.book.read('xl/styles.xml') if self.book.has('xl/styles.xml') else ''
        self.all_xml = self.book.all_xml()
        self.wb_xml = self.book.read(self.book.workbook_part)
        self.sheet_names = []
        for sheet in self.book.sheets:
            self.sheet_names.append((sheet.name, ''))
            self.sheets[sheet.name] = {'root': sheet.root, 'xml': sheet.xml}

    def fmt(self, root, ref):
        return self.book.format_in(root, ref)

    def sv(self, sheet_root, ref):
        return self.book.value_in(sheet_root, ref)

    def get_cell(self, sheet_root, ref):
        return self.book.node_in(sheet_root, ref)

    def grade(self):
        self.load()
        rr = []
        rr.append(self.r1()); rr.append(self.r2()); rr.append(self.r3())
        rr.append(self.r4()); rr.append(self.r5()); rr.append(self.r6())
        rr.append(self.r7()); rr.append(self.r8()); rr.append(self.r9())
        rr.append(self.r10())
        t = round(sum(x['s'] for x in rr), 1)
        rr.append({'req': 0, 'name': '总分', 's': t, 'full': 100, 'd': [], 'st': 'PASS' if t >= 60 else 'FAIL'})
        return rr

    def R(self, req, name, score, full, detail):
        s = min(float(score), float(full))
        st = 'PASS' if s >= full * 0.6 else 'FAIL'
        return {'req': req, 'name': name, 's': round(s, 1), 'full': int(full), 'd': detail, 'st': st}

    # ===== R1: 改名+标签颜色+插入行+标题 (10) =====
    def r1(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('员工工资表')
        if info is None:
            d.append('未找到"员工工资表"工作表'); return self.R(1, '工作表+标题', s, full, d)
        s += 2; d.append('改名"员工工资表" OK')

        # 标签颜色 — 在sheet XML中，字符串搜索
        if 'tabColor' in info['xml']: s += 2; d.append('蓝色标签 OK')

        # A1 标题 — 答案实际值是"2021年11月份员工工资表"
        root = info['root']
        a1 = self.sv(root, 'A1')
        if '2021' in a1 and '员工工资' in a1:
            s += 3; d.append(f'A1标题正确')
        else:
            d.append(f'A1={a1!r}')

        # 插入空行 → A2 应该=原来的表头 "姓名"
        a2 = self.sv(root, 'A2')
        if a2 == '姓名':
            s += 3; d.append('插入空行 OK')
        else:
            d.append(f'A2={a2!r} (期望=姓名)')
        return self.R(1, '工作表+标题', s, full, d)

    # ===== R2: 合并+字体+行高+列宽 (11) =====
    def r2(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('员工工资表')
        if info is None: return self.R(2, '格式设置', s, full, d)
        s1_xml = info['xml']
        root = info['root']

        # A1:I1 合并
        if 'mergeCell' in s1_xml and 'A1:I1' in s1_xml.replace(' ', ''):
            s += 2; d.append('A1:I1合并 OK')

        # 标题 A1：楷体/16/加粗/紫色；A2:I16 华文楷体（都按单元格取，不全文搜索）
        title_fmt = self.fmt(root, 'A1')
        t_font = title_fmt.get('font', {})
        if t_font.get('name') in ('楷体', '楷体_GB2312'):
            s += 1.5; d.append('标题楷体 OK')
        else:
            d.append('标题字体=%s' % t_font.get('name'))
        if t_font.get('bold'):
            s += 1.5; d.append('加粗 OK')
        if str(t_font.get('size')) == '16':
            s += 1.5; d.append('字号16 OK')
        else:
            d.append('标题字号=%s' % t_font.get('size'))
        t_color = self.book.resolve_color(t_font.get('color')) if hasattr(self, 'book') else ''
        if t_color in ('7030A0', '800080'):
            s += 1.5; d.append('紫色(7030A0) OK')
        else:
            d.append('标题颜色=%s' % (t_color or '未设置'))
        body_font = self.fmt(root, 'A2').get('font', {})
        if body_font.get('name') == '华文楷体':
            s += 1.5; d.append('A2:I16 华文楷体 OK')
        else:
            d.append('数据区字体=%s' % body_font.get('name'))
        # 行高32
        for row in root.iter('{%s}row' % S):
            if row.get('r') == '1' and row.get('ht'):
                ht = float(row.get('ht', '0'))
                if abs(ht - 32) < 2: s += 1.5; d.append(f'行高{ht:.1f} OK')
                break
        # 垂直居中
        if title_fmt.get('alignment', {}).get('vertical') == 'center':
            s += 0.5; d.append('垂直居中 OK')
        # 自动列宽 — 答案用cols/col的width属性
        if 'bestFit' in s1_xml: s += 1; d.append('自动列宽 OK')
        elif '<cols>' in s1_xml and 'width' in s1_xml: s += 1; d.append('列宽已设置 OK')
        return self.R(2, '格式设置', s, full, d)

    # ===== R3: 对齐+货币 (10) =====
    def r3(self):
        d = []; s = 0; full = 10
        # 水平居中
        if 'horizontal="center"' in self.all_xml: s += 3; d.append('居中 OK')
        # 右对齐
        if 'horizontal="right"' in self.all_xml: s += 2; d.append('右对齐 OK')
        # ￥货币格式 1位小数 — answer uses "¥"#,##0.0
        if '¥' in self.styles_xml or '\\u00a5' in self.styles_xml or '\u00a5' in self.styles_xml:
            s += 3; d.append('￥符号 OK')
        if '0.0' in self.styles_xml:
            s += 2; d.append('1位小数 OK')
        return self.R(3, '对齐+货币', s, full, d)

    # ===== R4: 加班总费用函数 + 评价列名 (10) =====
    def r4(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('员工工资表')
        if info is None: return self.R(4, '加班费用+评价', s, full, d)
        root = info['root']
        # H2 = "实发工资" → confirmed in answer
        h2 = self.sv(root, 'H2')
        if h2 == '实发工资': s += 2; d.append('H2=实发工资 OK')

        # H3 formula = SUM(E3,G3) → 应发+加班
        c = self.get_cell(root, 'H3')
        if c is not None:
            f = c.find('{%s}f' % S)
            if f is not None and 'SUM' in f.text:
                s += 4; d.append(f'加班公式({f.text}) OK')

        # I2 = "评价"
        i2 = self.sv(root, 'I2')
        if i2 == '评价':
            s += 4; d.append('I2=评价 OK')
        return self.R(4, '加班费用+评价', s, full, d)

    # ===== R5: IF函数 (10) =====
    def r5(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('员工工资表')
        if info is None: return self.R(5, 'IF评价函数', s, full, d)
        c = self.get_cell(info['root'], 'I3')
        if c is not None:
            f = c.find('{%s}f' % S)
            if f is not None and f.text:
                ft = f.text
                if '2.5' in ft: s += 4; d.append('>=2.5条件 OK')
                if '优秀' in ft: s += 3; d.append('优秀 OK')
                if '加油干' in ft: s += 3; d.append('加油干 OK')
        return self.R(5, 'IF评价函数', s, full, d)

    # ===== R6: 实发工资 = 应发+加班 (10) =====
    def r6(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('员工工资表')
        if info is None: return self.R(6, '实发工资函数', s, full, d)
        # H3: SUM(E3,G3) 即 应发+加班
        c = self.get_cell(info['root'], 'H3')
        if c is not None:
            f = c.find('{%s}f' % S)
            if f is not None and f.text:
                if 'SUM' in f.text: s += 6; d.append(f'SUM函数({f.text}) OK')
                elif '+' in f.text: s += 6; d.append(f'加法公式({f.text}) OK')
        # G2 列标题
        g2_val = self.sv(info['root'], 'G2')
        if '加班' in str(g2_val): s += 4; d.append('加班总费用列 OK')
        return self.R(6, '实发工资函数', s, full, d)

    # ===== R7: 男女平均年龄 (10) =====
    def r7(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('员工工资表')
        if info is None: return self.R(7, '平均年龄', s, full, d)
        s1_xml = info['xml']

        # 答案用 SUMIF+COUNTIF 计算男女平均（列K），不用AVERAGE
        has_avg = ('SUMIF' in s1_xml.upper() and 'COUNTIF' in s1_xml.upper())
        if has_avg: s += 6; d.append('平均函数(SUMIF/COUNTIF) OK')
        elif 'AVERAGE' in s1_xml.upper(): s += 4; d.append('AVERAGE函数 OK')

        # 男平均 and 女平均
        if '男' in s1_xml: s += 2; d.append('男平均 OK')
        if '女' in s1_xml: s += 2; d.append('女平均 OK')
        return self.R(7, '平均年龄', s, full, d)

    # ===== R8: 表样式+数据条+条件格式 (10) =====
    def r8(self):
        d = []; s = 0; full = 10
        # 套用表样式
        if 'xl/tables/table1.xml' in self.namelist:
            tbl = self.zf.read('xl/tables/table1.xml').decode('utf-8')
            if 'TableStyle' in tbl or 'tableStyle' in tbl:
                s += 3; d.append('套用表样式 OK')

        # 数据条 (blue gradient)
        if 'dataBar' in self.all_xml:
            s += 2; d.append('数据条 OK')

        # 条件格式 >1250：素材数据里本就可能有 1250 字样，
        # 必须落在条件格式公式/规则内才算
        import re as _re
        cf_blocks = _re.findall(r'<conditionalFormatting.*?</conditionalFormatting>',
                                self.all_xml, _re.S)
        if cf_blocks:
            s += 2; d.append('条件格式 OK')
            hit = False
            for block in cf_blocks:
                if '1250' in block or '&gt;1250' in block:
                    hit = True; break
            if hit:
                s += 3; d.append('>1250条件+格式 OK')
            else:
                d.append('条件格式中未含 1250')
        else:
            d.append('无条件格式')
        return self.R(8, '样式+条件格式', s, full, d)

    # ===== R9: 图表 (10) =====
    def r9(self):
        d = []; s = 0; full = 10
        if 'xl/charts/chart1.xml' not in self.namelist:
            d.append('未找到图表'); return self.R(9, '图表', s, full, d)

        chart = self.zf.read('xl/charts/chart1.xml').decode('utf-8')
        if 'bar3DChart' in chart: s += 3; d.append('三维簇状柱形图 OK')
        if '工资对比图' in chart: s += 3; d.append('标题"工资对比图" OK')
        if 'legend' in chart.lower(): s += 2; d.append('图例 OK')

        # 位置 A18:G30 area
        s1_xml = self.sheets.get('员工工资表', {}).get('xml', '')
        if 'drawing' in s1_xml.lower(): s += 2; d.append('图表嵌入位置 OK')
        return self.R(9, '图表', s, full, d)

    # ===== R10: 筛选+边框 (9) =====
    def r10(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('筛选')
        if info is None:
            d.append('未找到"筛选"工作表'); return self.R(10, '筛选+边框', s, full, d)

        root = info['root']; s3_xml = info['xml']
        rows = list(root.iter('{%s}row' % S))
        if len(rows) >= 5: s += 2; d.append(f'筛选数据({len(rows)}行) OK')

        # ≥1200 condition in A1:E2 area
        if '1200' in s3_xml: s += 2; d.append('>=1200条件 OK')

        # Blue double border — 在styles.xml中定义
        border_found = 'double' in self.all_xml.lower() and 'border' in self.all_xml.lower()
        if border_found: s += 3; d.append('内外双实线边框 OK')

        # No formulas = paste values
        has_f = any(f is not None and f.text for f in root.iter('{%s}f' % S))
        if not has_f: s += 2; d.append('粘贴值(无公式) OK')

        # Inserted 2 rows above → first data row should be row 3 or later
        has_row3 = any(row.get('r') == '3' for row in root.iter('{%s}row' % S))
        if has_row3: s += 1; d.append('插入2行 OK')
        return self.R(10, '筛选+边框', s, full, d)


def main():
    if sys.platform == 'win32':
        try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except: pass

    base = 'C:/Users/23000/Desktop/Word 2010 操作题题库/Excel 2010 操作题题库/1'
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base, '1结果.xlsx')
    print(f'评分: {os.path.basename(path)}')
    print('=' * 50)
    g = ExcelGrader1(path)
    results = g.grade()
    for r in results[:-1]:
        icon = 'PASS' if r['st'] == 'PASS' else 'FAIL'
        print(f"[{icon}] {r['req']:2d}. {r['name']:<12s}  {r['s']:5.1f} / {r['full']}")
        for d in r['d']: print(f'       {d}')
    print(f'{"=" * 50}\n  总分: {results[-1]["s"]} / 100')
    if len(sys.argv) < 2: input('\n按回车退出...')

if __name__ == '__main__':
    main()
