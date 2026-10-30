"""
Excel 2010 评分器 - 文件夹3 (v1 精确XML检测)

题干要求(9大项，拆为10个评分点):
1. Sheet1改名"销售报表" + 插入行 + A1标题
2. A1:G1合并居中 + 微软雅黑/加粗/16号/红色 + 行高30/垂直居中 + A2:G44华文楷体/11/居中
3. Sheet2复制数据 + G1="销售额" + G列=销售量*单价(PRODUCT) + 0位小数
4. G2="评语" + IF函数 (>=1000→"硕果累累"/"劳苦功高")
5. 套用表样式中等深浅2 + 条件格式(>=1000红色底纹深蓝字体) + A-G自动列宽
6. Sheet2页面设置: 页边距水平垂直居中 + 页眉"销售情况表"10号 + 页脚右页码起始5 + 横向 + A-G列宽
7. Sheet3→"筛选" + 复制Sheet2数据(保留源列宽) + 高级筛选>=1500
8. 筛选表红色边框(外最粗单线thick + 内双实线double)
9. 新建"分类汇总": 复制Sheet2数据 + 排序(日期降序/品名升序) + 分类汇总(品名/最大值/销售额)
10. 新建"图表": 三维簇状柱形图 + 标题"销售量统计图" + 图例靠右
"""
import zipfile, os, sys, io, xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import ExcelBook

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


class ExcelGrader3:
    def __init__(self, path):
        self.path = path; self.zf = None
        self.ss = []
        self.sheets = {}
        self.sheet_names = []
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

    def sv(self, sheet_root, ref):
        return self.book.value_in(sheet_root, ref)

    def get_cell(self, sheet_root, ref):
        return self.book.node_in(sheet_root, ref)

    def get_formula(self, sheet_root, ref):
        return self.book.formula_in(sheet_root, ref)

    def grade(self):
        self.load()
        rr = []
        rr.append(self.r1()); rr.append(self.r2()); rr.append(self.r3())
        rr.append(self.r4()); rr.append(self.r5()); rr.append(self.r6())
        rr.append(self.r7()); rr.append(self.r8()); rr.append(self.r9())
        rr.append(self.r10())
        t = round(sum(x['s'] for x in rr), 1)
        rr.append({'req': 0, 'name': '总分', 's': t, 'full': 100,
                   'd': [], 'st': 'PASS' if t >= 60 else 'FAIL'})
        return rr

    def R(self, req, name, score, full, detail):
        s = min(float(score), float(full))
        st = 'PASS' if s >= full * 0.6 else 'FAIL'
        return {'req': req, 'name': name, 's': round(s, 1),
                'full': int(full), 'd': detail, 'st': st}

    # ===== R1: 改名+插入行+A1标题 (10) =====
    def r1(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售报表')
        if info is None:
            d.append('未找到"销售报表"工作表')
            return self.R(1, '改名+插入行+标题', s, full, d)
        s += 3; d.append('改名"销售报表" OK')

        root = info['root']
        a1 = self.sv(root, 'A1')
        if '广东长贸' in a1 and '销售报表' in a1:
            s += 4; d.append(f'A1标题正确')
        else:
            d.append(f'A1={a1!r}')

        # 验证插入了空行: A2应该是原表头'日期'
        a2 = self.sv(root, 'A2')
        if a2 == '日期':
            s += 3; d.append('插入空行 OK')
        else:
            d.append(f'A2={a2!r} (期望=日期)')

        return self.R(1, '改名+插入行+标题', s, full, d)

    # ===== R2: A1:G1合并居中+标题格式+A2:G44格式 (10) =====
    def r2(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售报表')
        if info is None: return self.R(2, '标题+数据格式', s, full, d)
        s1_xml = info['xml']; root = info['root']

        # A1:G1 合并居中
        if 'mergeCell' in s1_xml and 'A1:G1' in s1_xml.replace(' ', ''):
            s += 2; d.append('A1:G1合并居中 OK')
        else:
            d.append('A1:G1未合并')

        # 微软雅黑
        if '微软雅黑' in self.all_xml:
            s += 1.5; d.append('微软雅黑 OK')

        # 加粗
        if self.book.format_in(root, 'A1').get('font', {}).get('bold'):
            s += 1; d.append('加粗 OK')

        # 字号16 + 红色(FF0000)
        if '<sz val="16"' in self.all_xml:
            s += 1; d.append('字号16 OK')
        if 'FF0000' in self.all_xml.upper():
            s += 1.5; d.append('红色字体 OK')

        # 行高30
        for row in root.iter('{%s}row' % S):
            if row.get('r') == '1':
                ht = row.get('ht', '')
                if ht and abs(float(ht) - 30) < 3:
                    s += 1; d.append(f'行高{float(ht):.0f} OK')
                break

        # 垂直居中 + A2:G44 华文楷体/11 + 居中
        if 'vertical="center"' in self.all_xml:
            s += 0.5; d.append('垂直居中 OK')
        if '华文楷体' in self.all_xml:
            s += 1; d.append('华文楷体 OK')
        if 'val="11"' in self.all_xml:
            s += 0.5; d.append('字号11 OK')
        if 'horizontal="center"' in self.all_xml:
            s += 0.5; d.append('水平居中 OK')

        return self.R(2, '标题+数据格式', s, full, d)

    # ===== R3: Sheet2复制数据+销售额列+PRODUCT+0位小数 (10) =====
    def r3(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('Sheet2')
        if info is None:
            d.append('未找到Sheet2'); return self.R(3, '销售额计算', s, full, d)
        root = info['root']

        # G1="销售额"
        g1 = self.sv(root, 'G1')
        if g1 == '销售额':
            s += 2; d.append('G1=销售额 OK')
        else:
            d.append(f'G1={g1!r}')

        # G2: PRODUCT(E2:F2) 或 E2*F2
        g2_f = self.get_formula(root, 'G2')
        if 'PRODUCT' in g2_f.upper():
            s += 4; d.append(f'PRODUCT公式({g2_f}) OK')
        elif '*' in g2_f:
            s += 3; d.append(f'乘法公式({g2_f}) OK')
        else:
            d.append(f'G2公式={g2_f!r}')

        # 数据填充到G43
        g43 = self.sv(root, 'G43')
        if g43 and g43 != '销售额':
            s += 2; d.append(f'数据填充至G43 OK')

        # 0位小数：素材样式表里也有普通 0 格式，必须作用在销售额列上才算
        fmt_ok = False
        for ref in ('G2', 'G3', 'G43'):
            numfmt = str(self.book.format_in(root, ref).get('numfmt', ''))
            if numfmt not in ('', 'General') and ('0' in numfmt):
                fmt_ok = True; break
        if fmt_ok:
            s += 2; d.append('数值格式 OK')
        else:
            d.append('销售额列未设置 0 位小数格式')

        return self.R(3, '销售额计算', s, full, d)

    # ===== R4: G2评语+IF函数 (10) =====
    def r4(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售报表')
        if info is None: return self.R(4, '评语IF函数', s, full, d)
        root = info['root']

        # G2="评语"
        g2 = self.sv(root, 'G2')
        if g2 == '评语':
            s += 3; d.append('G2=评语 OK')
        else:
            d.append(f'G2={g2!r}')

        # G3: IF(E3>=1000,"硕果累累","劳苦功高")
        g3_f = self.get_formula(root, 'G3')
        if 'IF' in g3_f.upper():
            s += 3; d.append(f'IF函数({g3_f}) OK')
            if '1000' in g3_f:
                s += 2; d.append('>=1000条件 OK')
            if '硕果累累' in g3_f:
                s += 1; d.append('"硕果累累" OK')
            if '劳苦功高' in g3_f:
                s += 1; d.append('"劳苦功高" OK')
        else:
            d.append(f'G3公式={g3_f!r}')

        return self.R(4, '评语IF函数', s, full, d)

    # ===== R5: 表样式+条件格式+自动列宽 (10) =====
    def r5(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('销售报表')
        if info is None: return self.R(5, '表样式+条件格式', s, full, d)

        # 套用表样式
        if 'xl/tables/table1.xml' in self.namelist:
            tbl = self.zf.read('xl/tables/table1.xml').decode('utf-8')
            if 'TableStyleMedium' in tbl:
                s += 4; d.append('套用表样式(TableStyleMedium) OK')
            elif 'tableStyleInfo' in tbl:
                s += 3; d.append('表样式已设置')

        # 条件格式: >=1000 红色底纹 + 深蓝字体
        all_parts = [self.all_xml]
        if 'xl/tables/table1.xml' in self.namelist:
            all_parts.append(tbl)
        full_search = '\n'.join(all_parts)
        has_cf = ('conditionalFormatting' in full_search or 'cfRule' in full_search)
        if has_cf:
            s += 3; d.append('条件格式 OK')

        # 自动列宽
        s1_xml = info['xml']
        if 'bestFit' in s1_xml and 'customWidth' in s1_xml:
            s += 3; d.append('A-G自动列宽 OK')
        elif 'customWidth' in s1_xml:
            s += 2; d.append('列宽已设置')

        return self.R(5, '表样式+条件格式', s, full, d)

    # ===== R6: Sheet2页面设置 (10) =====
    def r6(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('Sheet2')
        if info is None: return self.R(6, '页面设置', s, full, d)
        root = info['root']; s2_xml = info['xml']

        # 页边距水平垂直居中
        if 'horizontalCentered="1"' in s2_xml:
            s += 1.5; d.append('水平居中打印 OK')
        if 'verticalCentered="1"' in s2_xml:
            s += 1.5; d.append('垂直居中打印 OK')

        # 页眉: &C&10销售情况表
        if '销售情况表' in s2_xml:
            s += 2; d.append('页眉"销售情况表" OK')

        # 页脚: &R&P 靠右页码
        if '&amp;R&amp;P' in s2_xml:
            s += 1.5; d.append('页脚靠右页码 OK')
        elif '&R&P' in s2_xml:
            s += 1.5; d.append('页脚靠右页码 OK')

        # 起始页码5
        if 'firstPageNumber="5"' in s2_xml:
            s += 1; d.append('起始页码5 OK')

        # 横向
        if 'orientation="landscape"' in s2_xml:
            s += 1.5; d.append('纸张横向 OK')

        # A-G自动列宽
        if 'bestFit' in s2_xml and 'customWidth' in s2_xml:
            s += 1; d.append('A-G自动列宽 OK')

        return self.R(6, '页面设置', s, full, d)

    # ===== R7: Sheet3→筛选+复制数据+高级筛选 (10) =====
    def r7(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('筛选')
        if info is None:
            d.append('未找到"筛选"工作表')
            return self.R(7, '筛选表+高级筛选', s, full, d)
        s += 2; d.append('改名"筛选" OK')

        root = info['root']; s3_xml = info['xml']

        # 条件区域: A1:A2 >=1500
        a1 = self.sv(root, 'A1')
        a2 = self.sv(root, 'A2')
        if '1500' in a1 or '1500' in a2:
            s += 2; d.append(f'筛选条件>=1500 OK')
        else:
            # Check shared strings directly
            if '>=1500' in self.ss or '&gt;=1500' in s3_xml:
                s += 2; d.append('筛选条件>=1500 OK')

        # 数据从A3开始
        a3 = self.sv(root, 'A3')
        if a3:
            s += 1; d.append(f'数据从A3开始 OK')

        # 保留源列宽
        if 'customWidth' in s3_xml:
            s += 2; d.append('保留源列宽 OK')

        # 高级筛选 = 非匹配行隐藏
        hidden_count = sum(1 for row in root.iter('{%s}row' % S) if row.get('hidden') == '1')
        if hidden_count > 0:
            s += 3; d.append(f'高级筛选({hidden_count}行隐藏) OK')
        else:
            d.append('未发现隐藏行')

        return self.R(7, '筛选表+高级筛选', s, full, d)

    # ===== R8: 筛选表红色边框 (10) =====
    def r8(self):
        d = []; s = 0; full = 10
        # 红色边框: thick外 + double内
        # 外边框: thick + FF0000
        if 'thick' in self.all_xml and ('FF0000' in self.all_xml.upper() or '红' in self.all_xml):
            s += 5; d.append('外边框红色粗线(thick/FF0000) OK')
        elif 'thick' in self.all_xml:
            s += 3; d.append('外边框粗线(非红色)')

        # 内边框: double + FF0000
        if 'double' in self.all_xml and ('FF0000' in self.all_xml.upper()):
            s += 5; d.append('内边框红色双线(double/FF0000) OK')
        elif 'double' in self.all_xml:
            s += 3; d.append('内边框双线(非红色)')

        return self.R(8, '筛选表边框', s, full, d)

    # ===== R9: 分类汇总 (10) =====
    def r9(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('分类汇总')
        if info is None:
            d.append('未找到"分类汇总"工作表')
            return self.R(9, '分类汇总', s, full, d)
        s += 1.5; d.append('新建"分类汇总" OK')

        root = info['root']; s4_xml = info['xml']

        # 有数据
        rows = list(root.iter('{%s}row' % S))
        if len(rows) >= 5:
            s += 1.5; d.append(f'数据已填充({len(rows)}行) OK')

        # 排序: 日期降序(A) + 品名升序(B)
        has_sort = False
        for ss in root.iter('{%s}sortState' % S):
            has_sort = True
            conditions = list(ss.iter('{%s}sortCondition' % S))
            if len(conditions) >= 1:
                sc1_ref = conditions[0].get('ref', '')
                sc1_desc = conditions[0].get('descending', '')
                if 'A' in sc1_ref and sc1_desc == '1':
                    s += 2; d.append('日期降序 OK')
                else:
                    d.append(f'日期排序: {sc1_ref} desc={sc1_desc}')
            if len(conditions) >= 2:
                sc2_ref = conditions[1].get('ref', '')
                sc2_desc = conditions[1].get('descending', '')
                if 'B' in sc2_ref and sc2_desc != '1':
                    s += 1.5; d.append('品名升序 OK')
                else:
                    d.append(f'品名排序: {sc2_ref} desc={sc2_desc}')
        if not has_sort:
            d.append('未设置排序')

        # 分类汇总: outlineLevel on rows + SUBTOTAL
        has_outline = any(row.get('outlineLevel', '') != ''
                        for row in root.iter('{%s}row' % S))
        if has_outline:
            s += 2; d.append('分类汇总(outlineLevel) OK')

        # SUBTOTAL函数
        if 'SUBTOTAL' in s4_xml:
            s += 1.5; d.append('SUBTOTAL函数 OK')

        return self.R(9, '分类汇总', s, full, d)

    # ===== R10: 图表 (10) =====
    def r10(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('图表')
        if info is None:
            d.append('未找到"图表"工作表')
            return self.R(10, '图表', s, full, d)
        s += 2; d.append('新建"图表" OK')

        # 图表文件
        if 'xl/charts/chart1.xml' not in self.namelist:
            d.append('未找到图表文件')
            return self.R(10, '图表', s, full, d)

        chart = self.zf.read('xl/charts/chart1.xml').decode('utf-8')

        # 三维簇状柱形图
        if 'bar3DChart' in chart:
            s += 3; d.append('三维簇状柱形图(bar3DChart) OK')
        elif 'barChart' in chart:
            s += 2; d.append('簇状柱形图(非三维)')

        # 标题
        if '销售量统计图' in chart:
            s += 3; d.append('标题"销售量统计图" OK')

        # 图例靠右
        if 'legendPos' in chart and 'val="r"' in chart:
            s += 2; d.append('图例靠右(legendPos=r) OK')
        elif 'legend' in chart.lower():
            s += 1; d.append('图例存在')

        return self.R(10, '图表', s, full, d)


def main():
    if sys.platform == 'win32':
        try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except: pass

    base = os.path.dirname(os.path.abspath(__file__))
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base, '1结果.xlsx')
    print(f'评分: {os.path.basename(path)}')
    print('=' * 50)
    g = ExcelGrader3(path)
    results = g.grade()
    for r in results[:-1]:
        icon = 'PASS' if r['st'] == 'PASS' else 'FAIL'
        print(f"[{icon}] {r['req']:2d}. {r['name']:<16s}  {r['s']:5.1f} / {r['full']}")
        for d in r['d']:
            print(f'       {d}')
    print(f'{"=" * 50}\n  总分: {results[-1]["s"]} / 100')

if __name__ == '__main__':
    main()
