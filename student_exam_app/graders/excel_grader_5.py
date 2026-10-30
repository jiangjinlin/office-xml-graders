"""
Excel 2010 评分器 - 文件夹5 (v1, rels映射修复)

题干(9项, 10评分点):
1. Sheet1→"销售订单"+插入行+A1"2021年图书销售订单记录"
2. A1:H1合并居中+微软雅黑/加粗/12号/红色+行高28/垂直居中+A2:H58华文楷体/12/居中
3. VLOOKUP填充图书名称(D列)和作者(F列) from 图书编目表
4. 订单编号(A列)升序排序+重复项紫色字体+排顶端
5. 销售额=VLOOKUP单价*销量(本)+￥2位小数+A-H自动列宽
6. J4=SUMIF新华书店销量之和+表样式浅色11
7. 2021年图书销售分析SUMIFS(6-8月)+迷你折线图(高低点标记)
8. 数据透视表(新工作表→"2021年书店销量", A1起点)
9. 透视表: 图书名称=列标签, 书店名称=行标签, 销量(本)=求和项
"""
import zipfile, os, sys, io, xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import ExcelBook

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

class ExcelGrader5:
    def __init__(self, path):
        self.path = path; self.zf = None
        self.ss = []; self.sheets = {}; self.sheet_names = []
        self.all_xml = ''; self.styles_xml = ''; self.namelist = []

    def load(self):
        self.book = ExcelBook(self.path)
        self.zf = self.book.zf
        self.namelist = self.book.names
        self.ss = self.book.strings
        self.styles_xml = self.book.read('xl/styles.xml') if self.book.has('xl/styles.xml') else ''
        self.all_xml = self.book.all_xml()
        self.sheet_names = []
        for sheet in self.book.sheets:
            self.sheet_names.append((sheet.name, '', ''))
            self.sheets[sheet.name] = {'root': sheet.root, 'xml': sheet.xml}

    def sv(self, root, ref):
        return self.book.value_in(root, ref)

    def get_formula(self, root, ref):
        return self.book.formula_in(root, ref)

    def grade(self):
        self.load()
        rr = [self.r1(), self.r2(), self.r3(), self.r4(), self.r5(),
              self.r6(), self.r7(), self.r8(), self.r10()]
        t = round(sum(x['s'] for x in rr), 1)
        rr.append({'req': 0, 'name': '总分', 's': t, 'full': 100,
                   'd': [], 'st': 'PASS' if t >= 60 else 'FAIL'})
        return rr

    def R(self, req, name, score, full, detail):
        s = min(float(score), float(full))
        st = 'PASS' if s >= full * 0.6 else 'FAIL'
        return {'req': req, 'name': name, 's': round(s, 1),
                'full': int(full), 'd': detail, 'st': st}

    # R1: 改名+插入行+标题 (10)
    def r1(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售订单')
        if info is None: d.append('未找到"销售订单"'); return self.R(1, '改名+插入行+标题', s, full, d)
        s += 3; d.append('改名"销售订单" OK')
        root = info['root']
        a1 = self.sv(root, 'A1')
        if '2021' in a1 and '图书销售订单' in a1:
            s += 4; d.append('A1标题正确')
        else: d.append(f'A1={a1!r}')
        # 验证插入空行: A2应有原表头
        a2 = self.sv(root, 'A2')
        if a2 in ('订单编号',): s += 3; d.append('插入空行 OK')
        else: d.append(f'A2={a2!r}')
        return self.R(1, '改名+插入行+标题', s, full, d)

    # R2: 合并+标题格式+数据格式 (10)
    def r2(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售订单')
        if info is None: return self.R(2, '标题+数据格式', s, full, d)
        s1_xml = info['xml']; root = info['root']

        if 'mergeCell' in s1_xml and 'A1:H1' in s1_xml.replace(' ', ''):
            s += 1.5; d.append('A1:H1合并居中 OK')
        if '微软雅黑' in self.all_xml: s += 1; d.append('微软雅黑 OK')
        if self.book.format_in(root, 'A1').get('font', {}).get('bold'):
            s += 0.5; d.append('加粗 OK')
        if '<sz val="12"' in self.all_xml: s += 1; d.append('字号12 OK')
        if 'FF0000' in self.all_xml.upper(): s += 1; d.append('红色字体 OK')
        for row in root.iter('{%s}row' % S):
            if row.get('r') == '1':
                ht = row.get('ht', '')
                if ht and abs(float(ht) - 28) < 3: s += 1; d.append('行高28 OK')
                break
        if 'vertical="center"' in self.all_xml: s += 0.5; d.append('垂直居中 OK')
        if '华文楷体' in self.all_xml: s += 1.5; d.append('华文楷体 OK')
        if 'horizontal="center"' in self.all_xml: s += 0.5; d.append('水平居中 OK')
        # 数据区(58行)
        data_rows = len(list(root.iter('{%s}row' % S)))
        if data_rows >= 58: s += 1.5; d.append(f'数据{data_rows}行 OK')
        return self.R(2, '标题+数据格式', s, full, d)

    # R3: VLOOKUP填充图书名称+作者 (10)
    def r3(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售订单')
        if info is None: return self.R(3, 'VLOOKUP函数', s, full, d)
        root = info['root']

        # D3应该用VLOOKUP填图书名称
        d3_f = self.get_formula(root, 'D3')
        if 'VLOOKUP' in d3_f.upper():
            s += 5; d.append(f'VLOOKUP图书名称({d3_f}) OK')
        else: d.append(f'D3={d3_f!r}')

        # F3用VLOOKUP填作者
        f3_f = self.get_formula(root, 'F3')
        if 'VLOOKUP' in f3_f.upper():
            s += 5; d.append(f'VLOOKUP作者({f3_f}) OK')
        else: d.append(f'F3={f3_f!r}')
        return self.R(3, 'VLOOKUP函数', s, full, d)

    # R4: 排序+重复项紫色字体 (10)
    def r4(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售订单')
        if info is None: return self.R(4, '排序+紫色标记', s, full, d)
        root = info['root']

        # 排序: A列升序
        for ss in root.iter('{%s}sortState' % S):
            cons = list(ss.iter('{%s}sortCondition' % S))
            if cons and 'A' in cons[0].get('ref', '') and cons[0].get('descending') != '1':
                s += 4; d.append('订单编号升序 OK')

        # 紫色字体(7030A0)
        if '7030A0' in self.all_xml.upper():
            s += 6; d.append('重复项紫色字体(7030A0) OK')
        return self.R(4, '排序+紫色标记', s, full, d)

    # R5: 销售额计算+列宽 (10)
    def r5(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售订单')
        if info is None: return self.R(5, '销售额+列宽', s, full, d)
        root = info['root']

        # H3=销售额=VLOOKUP*G
        h3_f = self.get_formula(root, 'H3')
        if h3_f:
            s += 3; d.append(f'销售额公式({h3_f}) OK')

        # ￥货币+2位小数
        if '¥' in self.styles_xml or '00' in self.all_xml and '¥' in self.all_xml:
            s += 3; d.append('￥2位小数 OK')

        # A-H自动列宽
        s1_xml = info['xml']
        if 'bestFit' in s1_xml and 'customWidth' in s1_xml:
            s += 4; d.append('A-H自动列宽 OK')
        return self.R(5, '销售额+列宽', s, full, d)

    # R6: SUMIF+表样式 (10)
    def r6(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售订单')
        if info is None: return self.R(6, 'SUMIF+表样式', s, full, d)
        root = info['root']

        # J4=SUMIF新华书店
        j4_f = self.get_formula(root, 'J4')
        if 'SUMIF' in j4_f.upper():
            s += 5; d.append(f'SUMIF({j4_f}) OK')
        else: d.append(f'J4={j4_f!r}')

        # 表样式浅色11
        if 'xl/tables/table1.xml' in self.namelist:
            tbl = self.zf.read('xl/tables/table1.xml').decode('utf-8')
            if 'TableStyleLight' in tbl:
                s += 5; d.append('表样式浅色11 OK')
            elif 'tableStyleInfo' in tbl:
                s += 3; d.append('表样式已设置')
        return self.R(6, 'SUMIF+表样式', s, full, d)

    # R7: SUMIFS+迷你折线图 (10)
    def r7(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('2021年图书销售分析')
        if info is None:
            d.append('未找到"2021年图书销售分析"'); return self.R(7, 'SUMIFS+迷你图', s, full, d)
        s7_xml = info['xml']

        # SUMIFS
        if 'SUMIFS' in s7_xml.upper():
            s += 5; d.append('SUMIFS函数 OK')

        # 迷你折线图: sparkline
        if 'sparkline' in s7_xml.lower() or 'sparklineGroup' in s7_xml:
            s += 3; d.append('迷你折线图 OK')

        # 高低点标记
        if 'high' in s7_xml.lower() or 'markers' in s7_xml.lower():
            s += 2; d.append('高低点标记 OK')
        return self.R(7, 'SUMIFS+迷你图', s, full, d)

    # R8+R9: 数据透视表 (20)
    def r8(self):
        d = []; s = 0; full = 20
        info = self.sheets.get('2021年书店销量')
        if info is None:
            d.append('未找到"2021年书店销量"')
            return self.R(8, '数据透视表', s, full, d)
        s += 4; d.append('改名"2021年书店销量" OK')

        # 透视表文件存在
        if 'xl/pivotTables/pivotTable1.xml' in self.namelist:
            pt = self.zf.read('xl/pivotTables/pivotTable1.xml').decode('utf-8')
            s += 4; d.append('数据透视表文件 OK')

            # 检查pivotCache中的字段名
            pc = ''
            if 'xl/pivotCache/pivotCacheDefinition1.xml' in self.namelist:
                pc = self.zf.read('xl/pivotCache/pivotCacheDefinition1.xml').decode('utf-8')
            full_search = pt + pc

            if '书店名称' in full_search: s += 4; d.append('行标签:书店名称 OK')
            if '图书名称' in full_search: s += 4; d.append('列标签:图书名称 OK')

        # A1为起点=透视表在A1
        a1 = self.sv(info['root'], 'A1')
        if a1 and '销量' in a1: s += 4; d.append(f'透视表A1起点({a1}) OK')
        return self.R(8, '数据透视表', s, full, d)

    # R10: 综合检查(数据完整性+VLOOKUP准确性) (10)
    def r10(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售订单')
        if info is None: return self.R(10, '综合检查', s, full, d)
        root = info['root']

        # D3填了图书名称(非空)
        d3 = self.sv(root, 'D3')
        if d3 and len(d3) > 0: s += 3; d.append(f'D3图书名称={d3} OK')

        # F3填了作者
        f3 = self.sv(root, 'F3')
        if f3 and len(f3) > 0: s += 3; d.append(f'F3作者={f3} OK')

        # H58有销售额
        h58 = self.sv(root, 'H58')
        if h58: s += 2; d.append('H58销售额填充 OK')

        # SUMIF结果J4有值
        j4 = self.sv(root, 'J4')
        if j4: s += 2; d.append(f'J4={j4} OK')
        return self.R(10, '综合检查', s, full, d)


def main():
    if sys.platform == 'win32':
        try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except: pass
    base = os.path.dirname(os.path.abspath(__file__))
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base, '1结果.xlsx')
    print(f'评分: {os.path.basename(path)}')
    print('=' * 50)
    g = ExcelGrader5(path)
    results = g.grade()
    for r in results[:-1]:
        icon = 'PASS' if r['st'] == 'PASS' else 'FAIL'
        print(f"[{icon}] {r['req']:2d}. {r['name']:<18s}  {r['s']:5.1f} / {r['full']}")
        for det in r['d']: print(f'       {det}')
    print(f'{"=" * 50}\n  总分: {results[-1]["s"]} / 100')

if __name__ == '__main__':
    main()
