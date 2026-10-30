"""
Excel 2010 评分器 - 文件夹4 (v1)

题干(9项, 10评分点):
1. Sheet1→"员工培训成绩表"+插入行+A1"2021年员工培训成绩表"
2. A1:M1合并居中+华文行楷/加粗/20号/红色+行高30/垂直居中+A2:M31华文宋体/12/居中
3. SUM总分+AVERAGE平均分+RANK名次+2位小数+页边距居中+A-M自动列宽
4. 页眉"员工培训成绩表"12号红色居中+页脚靠右页码+A2:L31红粗外/紫双内边框
5. A2:K31复制到分类汇总(Sheet2)保留源列宽
6. L1="备注"+IF(>=80→"优秀"/"努力")
7. Sheet3→"筛选"+复制A2:L31到A3(值+数字格式)+高级筛选(总分>=450/平均分>=80)
8. Sheet2→"分类汇总"+排序(部门降序/姓名降序)+分类汇总(平均值/顾客心理/市场开拓)+2位小数
9. 新建"图表"+带标记折线图+标题"员工平均分统计图"(楷体/加粗/12)+图例靠右
10. 综合: 边框/数据完整性
"""
import zipfile, os, sys, io, xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import ExcelBook

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

class ExcelGrader4:
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
        self.wb_xml = self.book.read(self.book.workbook_part)
        self.sheet_names = []
        for sheet in self.book.sheets:
            self.sheet_names.append((sheet.name, ''))
            self.sheets[sheet.name] = {'root': sheet.root, 'xml': sheet.xml}

    def sv(self, sheet_root, ref):
        return self.book.value_in(sheet_root, ref)

    def get_formula(self, sheet_root, ref):
        return self.book.formula_in(sheet_root, ref)

    def grade(self):
        self.load()
        rr = [self.r1(), self.r2(), self.r3(), self.r4(), self.r5(),
              self.r6(), self.r7(), self.r8(), self.r9(), self.r10()]
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
        info = self.sheets.get('员工培训成绩表')
        if info is None: d.append('未找到"员工培训成绩表"'); return self.R(1, '改名+插入行+标题', s, full, d)
        s += 3; d.append('改名"员工培训成绩表" OK')
        root = info['root']
        a1 = self.sv(root, 'A1')
        if '2021' in a1 and '员工培训成绩表' in a1:
            s += 4; d.append('A1标题正确')
        else: d.append(f'A1={a1!r}')
        a2 = self.sv(root, 'A2')
        if a2 == '员工编号': s += 3; d.append('插入空行 OK')
        else: d.append(f'A2={a2!r}')
        return self.R(1, '改名+插入行+标题', s, full, d)

    # R2: 标题+数据格式 (10)
    def r2(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('员工培训成绩表')
        if info is None: return self.R(2, '标题+数据格式', s, full, d)
        s1_xml = info['xml']; root = info['root']

        if 'mergeCell' in s1_xml and 'A1:M1' in s1_xml.replace(' ', ''):
            s += 1.5; d.append('A1:M1合并居中 OK')
        if '华文行楷' in self.all_xml: s += 1.5; d.append('华文行楷 OK')
        if self.book.format_in(root, 'A1').get('font', {}).get('bold'):
            s += 1; d.append('加粗 OK')
        if '<sz val="20"' in self.all_xml: s += 1; d.append('字号20 OK')
        if 'FF0000' in self.all_xml.upper(): s += 1; d.append('红色字体 OK')
        for row in root.iter('{%s}row' % S):
            if row.get('r') == '1':
                ht = row.get('ht', '')
                if ht and abs(float(ht) - 30) < 3: s += 1; d.append(f'行高30 OK')
                break
        if 'vertical="center"' in self.all_xml: s += 0.5; d.append('垂直居中 OK')
        if '华文宋体' in self.all_xml: s += 1.5; d.append('华文宋体 OK')
        if '<sz val="12"' in self.all_xml: s += 0.5; d.append('字号12 OK')
        if 'horizontal="center"' in self.all_xml: s += 0.5; d.append('水平居中 OK')
        return self.R(2, '标题+数据格式', s, full, d)

    # R3: 总分+平均分+名次+页面+列宽 (10)
    def r3(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('员工培训成绩表')
        if info is None: return self.R(3, '公式+页面+列宽', s, full, d)
        root = info['root']; s1_xml = info['xml']

        j3_f = self.get_formula(root, 'J3')
        if 'SUM' in j3_f.upper(): s += 1.5; d.append(f'SUM({j3_f}) OK')
        k3_f = self.get_formula(root, 'K3')
        if 'AVERAGE' in k3_f.upper(): s += 1.5; d.append(f'AVERAGE({k3_f}) OK')
        l3_f = self.get_formula(root, 'L3')
        if 'RANK' in l3_f.upper(): s += 1.5; d.append(f'RANK({l3_f}) OK')
        if '0.00' in self.styles_xml: s += 1; d.append('2位小数 OK')

        if 'horizontalCentered="1"' in s1_xml: s += 1; d.append('水平居中打印 OK')
        if 'verticalCentered="1"' in s1_xml: s += 1; d.append('垂直居中打印 OK')
        if 'bestFit' in s1_xml and 'customWidth' in s1_xml: s += 2.5; d.append('A-M自动列宽 OK')
        return self.R(3, '公式+页面+列宽', s, full, d)

    # R4: 页眉/页脚+边框 (10)
    def r4(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('员工培训成绩表')
        if info is None: return self.R(4, '页眉页脚+边框', s, full, d)
        s1_xml = info['xml']

        # 页眉: &C&12&"字体"员工培训成绩表 (红色)
        if '员工培训成绩表' in s1_xml: s += 2; d.append('页眉"员工培训成绩表" OK')
        # 页脚靠右页码
        if '&amp;R&amp;P' in s1_xml: s += 1.5; d.append('页脚靠右页码 OK')
        elif '&R&P' in s1_xml: s += 1.5; d.append('页脚靠右页码 OK')

        # 边框: 红色粗线外(thick/FF0000) + 紫色双线内(double/7030A0)
        if 'thick' in self.all_xml and 'FF0000' in self.all_xml.upper():
            s += 3.5; d.append('外边框红色粗线 OK')
        if 'double' in self.all_xml and '7030A0' in self.all_xml.upper():
            s += 3; d.append('内边框紫色双线(7030A0) OK')
        return self.R(4, '页眉页脚+边框', s, full, d)

    # R5: 复制数据到分类汇总/Sheet2 (10)
    def r5(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('分类汇总')
        if info is None: d.append('未找到"分类汇总"'); return self.R(5, '复制数据', s, full, d)
        root = info['root']; s2_xml = info['xml']

        a1 = self.sv(root, 'A1')
        if a1 == 'PX01' or a1 == '员工编号': s += 4; d.append(f'数据已复制(A1={a1}) OK')
        # 保留源列宽
        if 'customWidth' in s2_xml: s += 3; d.append('保留源列宽 OK')
        # 数据填充到K30
        k30 = self.sv(root, 'K30')
        if k30: s += 3; d.append('数据填充完整 OK')
        return self.R(5, '复制数据', s, full, d)

    # R6: L1备注+IF函数 (10)
    def r6(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('分类汇总')
        if info is None: return self.R(6, '备注IF函数', s, full, d)
        root = info['root']

        l1 = self.sv(root, 'L1')
        if l1 == '备注': s += 3; d.append('L1=备注 OK')
        else: d.append(f'L1={l1!r}')

        l2_f = self.get_formula(root, 'L2')
        if 'IF' in l2_f.upper():
            s += 3; d.append(f'IF函数({l2_f}) OK')
            if '80' in l2_f: s += 2; d.append('>=80条件 OK')
            if '优秀' in l2_f: s += 1; d.append('"优秀" OK')
            if '努力' in l2_f: s += 1; d.append('"努力" OK')
        else: d.append(f'L2公式={l2_f!r}')
        return self.R(6, '备注IF函数', s, full, d)

    # R7: Sheet3→筛选+高级筛选 (10)
    def r7(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('筛选')
        if info is None: d.append('未找到"筛选"'); return self.R(7, '筛选+高级筛选', s, full, d)
        s += 2; d.append('改名"筛选" OK')
        root = info['root']; s3_xml = info['xml']

        # 条件: A1:B2 (总分>=450, 平均分>=80)
        a1 = self.sv(root, 'A1')
        b1 = self.sv(root, 'B1')
        a2 = self.sv(root, 'A2')
        b2 = self.sv(root, 'B2')
        if '总分' in a1: s += 1; d.append(f'A1=总分 OK')
        if '平均分' in b1: s += 1; d.append(f'B1=平均分 OK')
        if '450' in a2: s += 1; d.append(f'A2>=450 OK')
        if '80' in b2: s += 1; d.append(f'B2>=80 OK')

        # 数据从A3开始
        a3 = self.sv(root, 'A3')
        if a3: s += 1; d.append('数据从A3开始 OK')

        # 粘贴值(无公式)
        has_f = any(f is not None and f.text for f in root.iter('{%s}f' % S))
        if not has_f: s += 1; d.append('粘贴值(无公式) OK')

        # 高级筛选(hidden行)
        hidden = 0
        for row in root.iter('{%s}row' % S):
            if row.get('hidden') == '1':
                hidden += 1
        if hidden > 0: s += 2; d.append(f'高级筛选({hidden}行隐藏) OK')
        else: d.append('未发现隐藏行(筛选可能未应用)')
        return self.R(7, '筛选+高级筛选', s, full, d)

    # R8: 分类汇总: 排序+SUBTOTAL (10)
    def r8(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('分类汇总')
        if info is None: return self.R(8, '分类汇总', s, full, d)
        root = info['root']; s2_xml = info['xml']

        # 排序
        for ss in root.iter('{%s}sortState' % S):
            cons = list(ss.iter('{%s}sortCondition' % S))
            if len(cons) >= 2:
                sc1_ref = cons[0].get('ref', '')
                sc1_desc = cons[0].get('descending', '')
                sc2_ref = cons[1].get('ref', '')
                sc2_desc = cons[1].get('descending', '')
                if 'C' in sc1_ref and sc1_desc == '1': s += 2; d.append('部门降序 OK')
                if 'B' in sc2_ref and sc2_desc == '1': s += 2; d.append('姓名降序 OK')

        # 分类汇总: SUBTOTAL
        if 'SUBTOTAL' in s2_xml:
            s += 3; d.append('SUBTOTAL分类汇总 OK')

        # 2位小数
        if '0.00' in self.styles_xml: s += 3; d.append('2位小数 OK')
        return self.R(8, '分类汇总', s, full, d)

    # R9: 图表 (10)
    def r9(self):
        d = []; s = 0; full = 10
        if '图表' not in self.sheets: d.append('未找到"图表"'); return self.R(9, '图表', s, full, d)
        s += 2; d.append('新建"图表" OK')

        if 'xl/charts/chart1.xml' not in self.namelist:
            d.append('未找到图表文件'); return self.R(9, '图表', s, full, d)

        s4_xml = self.sheets['图表']['xml']
        chart = self.zf.read('xl/charts/chart1.xml').decode('utf-8')

        # 带标记折线图
        if 'lineChart' in chart:
            s += 2; d.append('折线图(lineChart) OK')
        if 'marker' in chart.lower():
            s += 1; d.append('带数据标记 OK')

        # 标题+字体
        if '员工平均分统计图' in chart:
            s += 2; d.append('标题"员工平均分统计图" OK')
        if '楷体' in chart: s += 1; d.append('标题楷体 OK')
        if 'b=' in chart or '<a:b' in chart: s += 0.5; d.append('标题加粗 OK')
        if '1200' in chart: s += 0.5; d.append('标题12号 OK')

        # 图例靠右
        if 'legendPos' in chart and 'val="r"' in chart:
            s += 1; d.append('图例靠右 OK')
        return self.R(9, '图表', s, full, d)

    # R10: 综合检查(数据完整性+合并+行高) (10)
    def r10(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('员工培训成绩表')
        if info is None: return self.R(10, '综合检查', s, full, d)
        root = info['root']

        # J31/K31/L31 有值(公式填充完整)
        j31 = self.sv(root, 'J31')
        k31 = self.sv(root, 'K31')
        l31 = self.sv(root, 'L31')
        if j31 and k31 and l31: s += 4; d.append('J/K/L列填充完整 OK')

        # A1:M1合并(已在R2检查, 这里验证范围)
        s1_xml = info['xml']
        if 'A1:M1' in s1_xml.replace(' ', ''): s += 2; d.append('合并范围A1:M1 OK')

        # 自动列宽A-M (已在R3检查, 这里验证12列)
        cols = root.find('{%s}cols' % S)
        if cols is not None:
            col_list = list(cols.iter('{%s}col' % S))
            max_col = max(int(c.get('max','0')) for c in col_list) if col_list else 0
            if max_col >= 12: s += 2; d.append(f'A-M共{max_col}列 OK')

        # IF公式填充到L31(分类汇总)
        info2 = self.sheets.get('分类汇总')
        if info2 is not None:
            l30 = self.sv(info2['root'], 'L30')
            if l30: s += 2; d.append('备注填充完整 OK')
        return self.R(10, '综合检查', s, full, d)


def main():
    if sys.platform == 'win32':
        try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except: pass
    base = os.path.dirname(os.path.abspath(__file__))
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base, '1结果.xlsx')
    print(f'评分: {os.path.basename(path)}')
    print('=' * 50)
    g = ExcelGrader4(path)
    results = g.grade()
    for r in results[:-1]:
        icon = 'PASS' if r['st'] == 'PASS' else 'FAIL'
        print(f"[{icon}] {r['req']:2d}. {r['name']:<18s}  {r['s']:5.1f} / {r['full']}")
        for d in r['d']: print(f'       {d}')
    print(f'{"=" * 50}\n  总分: {results[-1]["s"]} / 100')

if __name__ == '__main__':
    main()
