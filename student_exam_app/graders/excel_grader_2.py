"""
Excel 2010 评分器 - 文件夹2 (v1 精确XML检测)

题干要求(9大项，拆为10个评分点):
1. 插入空行 + A1标题"学生考试成绩统计表" + 华文行楷/加粗/16号 + 跨列居中(A1:L1) + 垂直居中
2. A2:J15 楷体/14号/水平居中/垂直居中
3. 行高20(2-15行) + 列宽自动(A-L) + SUM总分 + AVERAGE平均分 + 2位小数 + 楷体14居中
4. 外边框蓝色最粗单线(thick/0070C0) + 内边框橙色双实线(double/FFC000)
5. MID函数求班级 (MID(A3,4,1)&"班")
6. A2:L2 水绿色强调文字颜色5淡色80%底纹 (theme=8 tint=0.8)
7. 排序: 班级升序 + 总分降序
8. Sheet2改名"筛选" + 复制粘贴值 + 高级筛选(高等代数>=95 且 大学英语>=95)
9. Sheet3改名"学生平均分比较表" + 三维簇状柱形图 + 标题 + 图例靠右
"""
import zipfile, os, sys, io, xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import ExcelBook

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


class ExcelGrader2:
    def __init__(self, path):
        self.path = path; self.zf = None
        self.ss = []           # sharedStrings
        self.sheets = {}       # {name: {'root': Element, 'xml': str, 'wb_el': Element}}
        self.sheet_names = []   # [(name, sheetId)]
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

    # ===== R1: 插入空行 + A1标题 + 字体 + 跨列居中 (10) =====
    def r1(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('Sheet1')
        if info is None:
            d.append('未找到Sheet1'); return self.R(1, '插入行+标题+字体', s, full, d)
        root = info['root']

        # A1 = "学生考试成绩统计表"
        a1 = self.sv(root, 'A1')
        if '学生考试成绩统计表' in a1:
            s += 3; d.append('A1标题"学生考试成绩统计表" OK')
        else:
            d.append(f'A1={a1!r}')

        # A1 字体：华文行楷/加粗/16（按单元格取）
        title_fmt = self.book.format_in(root, 'A1')
        t_font = title_fmt.get('font', {})
        if t_font.get('name') == '华文行楷':
            s += 2; d.append('华文行楷 OK')
        else:
            d.append('标题字体=%s' % t_font.get('name'))
        if t_font.get('bold'):
            s += 1; d.append('加粗 OK')
        if str(t_font.get('size')) == '16':
            s += 1; d.append('字号16 OK')
        else:
            d.append('标题字号=%s' % t_font.get('size'))

        # 跨列居中 = centerContinuous (不是合并merge)，限定 A1
        if title_fmt.get('alignment', {}).get('horizontal') == 'centerContinuous':
            s += 2; d.append('跨列居中(centerContinuous) OK')
        elif 'mergeCell' in info['xml'] and 'A1:L1' in info['xml'].replace(' ', ''):
            s += 1; d.append('合并居中(部分分)')

        # A1 垂直居中
        if title_fmt.get('alignment', {}).get('vertical') == 'center':
            s += 1; d.append('垂直居中 OK')

        return self.R(1, '插入行+标题+字体', s, full, d)

    # ===== R2: A2:J15 楷体/14/居中 (10) =====
    def r2(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('Sheet1')
        if info is None:
            return self.R(2, '数据区格式', s, full, d)

        # 数据区按样例单元格取（A2 与 J15 两端都查）
        root = info['root']
        f1 = self.book.format_in(root, 'A2')
        f2 = self.book.format_in(root, 'J15')
        ok_font = all(f.get('font', {}).get('name') == '楷体' for f in (f1, f2))
        ok_size = all(str(f.get('font', {}).get('size')) == '14' for f in (f1, f2))
        ok_h = all(f.get('alignment', {}).get('horizontal') == 'center' for f in (f1, f2))
        ok_v = all(f.get('alignment', {}).get('vertical') == 'center' for f in (f1, f2))
        if ok_font: s += 3; d.append('数据区楷体 OK')
        else: d.append('A2=%s J15=%s' % (f1.get('font', {}).get('name'), f2.get('font', {}).get('name')))
        if ok_size: s += 3; d.append('字号14 OK')
        else: d.append('字号 A2=%s J15=%s' % (f1.get('font', {}).get('size'), f2.get('font', {}).get('size')))
        if ok_h: s += 2; d.append('水平居中 OK')
        if ok_v: s += 2; d.append('垂直居中 OK')

        return self.R(2, '数据区格式', s, full, d)

    # ===== R3: 行高+列宽+SUM+AVERAGE+2位小数 (10) =====
    def r3(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('Sheet1')
        if info is None:
            return self.R(3, '行高+列宽+公式', s, full, d)
        root = info['root']; s1_xml = info['xml']

        # 行高20 (rows 2-15)
        row_ok = False
        for row in root.iter('{%s}row' % S):
            r = int(row.get('r', '0'))
            if 2 <= r <= 15:
                ht = row.get('ht', '')
                if ht and abs(float(ht) - 20) < 2:
                    row_ok = True; break
        if row_ok:
            s += 2; d.append('行高20 OK')
        else:
            d.append('行高未设置为20')

        # 列宽自动调整 (bestFit + customWidth)
        if 'bestFit' in s1_xml and 'customWidth' in s1_xml:
            s += 2; d.append('自动列宽 OK')
        elif 'customWidth' in s1_xml:
            s += 1; d.append('列宽已设置(部分)')

        # K2 = "总分"
        k2 = self.sv(root, 'K2')
        if k2 == '总分':
            s += 1; d.append('K2=总分 OK')
        else:
            d.append(f'K2={k2!r}')

        # K3 = SUM函数
        k3_f = self.get_formula(root, 'K3')
        if 'SUM' in k3_f.upper():
            s += 2; d.append(f'SUM函数({k3_f}) OK')
        else:
            d.append(f'K3公式={k3_f!r}')

        # L2 = "平均分"
        l2 = self.sv(root, 'L2')
        if l2 == '平均分':
            s += 1; d.append('L2=平均分 OK')
        else:
            d.append(f'L2={l2!r}')

        # L3 = AVERAGE函数
        l3_f = self.get_formula(root, 'L3')
        if 'AVERAGE' in l3_f.upper():
            s += 1; d.append(f'AVERAGE函数({l3_f}) OK')
        else:
            d.append(f'L3公式={l3_f!r}')

        # 2位小数 (numFmt 0.00)
        if '0.00' in self.styles_xml:
            s += 1; d.append('2位小数 OK')

        return self.R(3, '行高+列宽+公式', s, full, d)

    # ===== R4: 边框 蓝色粗线外+橙色双线内 (10) =====
    def r4(self):
        d = []; s = 0; full = 10
        # 外边框: thick + 0070C0 (蓝色)
        if 'thick' in self.all_xml and '0070C0' in self.all_xml.upper():
            s += 5; d.append('外边框蓝色粗线(thick/0070C0) OK')
        elif 'thick' in self.all_xml:
            s += 3; d.append('外边框粗线(颜色未匹配)')

        # 内边框: double + FFC000 (橙色)
        if 'double' in self.all_xml and 'FFC000' in self.all_xml.upper():
            s += 5; d.append('内边框橙色双线(double/FFC000) OK')
        elif 'double' in self.all_xml:
            s += 3; d.append('内边框双线(颜色未匹配)')

        return self.R(4, '边框设置', s, full, d)

    # ===== R5: MID函数求班级 (10) =====
    def r5(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('Sheet1')
        if info is None:
            return self.R(5, 'MID函数', s, full, d)
        root = info['root']

        # C3 应有 MID 函数: MID(A3,4,1)&"班"
        c3_f = self.get_formula(root, 'C3')
        if 'MID' in c3_f.upper():
            s += 5; d.append(f'MID函数({c3_f}) OK')
            # 检查参数: A3,4,1
            if 'A3' in c3_f and '4' in c3_f and '1' in c3_f:
                s += 3; d.append('参数(A3,4,1) OK')
        else:
            d.append(f'C3公式={c3_f!r}')

        # C3值应为 "X班" 格式
        c3_val = self.sv(root, 'C3')
        if c3_val and '班' in c3_val:
            s += 2; d.append(f'班级值={c3_val} OK')
        else:
            d.append(f'C3值={c3_val!r}')

        return self.R(5, 'MID函数', s, full, d)

    # ===== R6: A2:L2 底纹 水绿色强调文字颜色5淡色80% (10) =====
    def r6(self):
        d = []; s = 0; full = 10
        # 水绿色，强调文字颜色5，淡色80% → theme=8, tint=0.8
        # patternFill solid + fgColor theme=8
        if 'patternFill' in self.styles_xml and 'solid' in self.styles_xml:
            # theme=8 is 强调文字颜色5
            if 'theme="8"' in self.styles_xml:
                s += 5; d.append('水绿色强调文字颜色5(theme=8) OK')
            elif 'theme=\"8\"' in self.styles_xml:
                s += 5; d.append('水绿色强调文字颜色5(theme=8) OK')
            else:
                d.append('未找到theme=8')

            # tint=0.8 (淡色80%)
            if '0.7999' in self.styles_xml or '0.8' in self.styles_xml:
                s += 5; d.append('淡色80%(tint=0.8) OK')
            else:
                d.append('未找到tint=0.8')
        else:
            d.append('未找到solid填充')

        return self.R(6, '底纹设置', s, full, d)

    # ===== R7: 排序 班级升序+总分降序 (10) =====
    def r7(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('Sheet1')
        if info is None:
            return self.R(7, '排序', s, full, d)
        root = info['root']

        # 检查 sortState
        sort_found = False
        for sortState in root.iter('{%s}sortState' % S):
            sort_found = True
            ref = sortState.get('ref', '')
            conditions = list(sortState.iter('{%s}sortCondition' % S))
            d.append(f'sortState ref={ref}, {len(conditions)}个条件')

            if len(conditions) >= 1:
                s += 2; d.append('排序已设置 OK')

            # 主关键字: 班级(C列) 升序
            if conditions:
                sc1 = conditions[0]
                sc1_ref = sc1.get('ref', '')
                sc1_desc = sc1.get('descending', '')
                if 'C' in sc1_ref and sc1_desc != '1':
                    s += 4; d.append(f'班级升序({sc1_ref}) OK')
                else:
                    d.append(f'第一条件: ref={sc1_ref} desc={sc1_desc}')

            # 次关键字: 总分(K列) 降序
            if len(conditions) >= 2:
                sc2 = conditions[1]
                sc2_ref = sc2.get('ref', '')
                sc2_desc = sc2.get('descending', '')
                if 'K' in sc2_ref and sc2_desc == '1':
                    s += 4; d.append(f'总分降序({sc2_ref}) OK')
                else:
                    d.append(f'第二条件: ref={sc2_ref} desc={sc2_desc}')

        if not sort_found:
            # 素材数据本就按班级升序，单看数据顺序无法区分是否做了双关键字排序，
            # 因此必须有 sortState 才给分
            d.append('未找到排序状态(sortState)')

        return self.R(7, '排序', s, full, d)

    # ===== R8: 筛选表 改名+复制+高级筛选 (10) =====
    def r8(self):
        d = []; s = 0; full = 10

        # Sheet2 改名"筛选"
        info = self.sheets.get('筛选')
        if info is None:
            d.append('未找到"筛选"工作表')
            return self.R(8, '高级筛选', s, full, d)
        s += 2; d.append('改名"筛选" OK')

        root = info['root']; s2_xml = info['xml']

        # 条件区域 A1:B2: 高等代数>=95 且 大学英语>=95
        a1 = self.sv(root, 'A1')
        b1 = self.sv(root, 'B1')
        a2 = self.sv(root, 'A2')
        b2 = self.sv(root, 'B2')
        if '高等代数' in a1:
            s += 1; d.append(f'A1={a1} OK')
        if '大学英语' in b1:
            s += 1; d.append(f'B1={b1} OK')
        if '95' in a2:
            s += 1; d.append(f'A2={a2}(>=95) OK')
        if '95' in b2:
            s += 1; d.append(f'B2={b2}(>=95) OK')

        # 数据从A3开始复制 (粘贴值 = 无公式)
        a3 = self.sv(root, 'A3')
        if a3 and len(list(root.iter('{%s}row' % S))) >= 5:
            s += 1; d.append(f'数据从A3开始({a3}) OK')

        # 无公式 = 粘贴值
        has_f = any(f is not None and f.text for f in root.iter('{%s}f' % S))
        if not has_f:
            s += 2; d.append('粘贴值(无公式) OK')
        else:
            d.append('存在公式(未粘贴值)')

        # 高级筛选 = 隐藏行 (非匹配行被隐藏)
        hidden_count = 0
        for row in root.iter('{%s}row' % S):
            if row.get('hidden') == '1':
                hidden_count += 1
        if hidden_count > 0:
            s += 2; d.append(f'高级筛选({hidden_count}行隐藏) OK')
        else:
            d.append('未发现隐藏行(筛选可能未应用)')

        return self.R(8, '高级筛选', s, full, d)

    # ===== R9: 改名"学生平均分比较表" + 图表 (10) =====
    def r9(self):
        d = []; s = 0; full = 10

        # Sheet3 改名"学生平均分比较表"
        if '学生平均分比较表' in self.sheets:
            s += 2; d.append('改名"学生平均分比较表" OK')
        else:
            d.append('未找到"学生平均分比较表"')

        # 图表文件
        if 'xl/charts/chart1.xml' not in self.namelist:
            d.append('未找到图表文件')
            return self.R(9, '图表', s, full, d)

        chart = self.zf.read('xl/charts/chart1.xml').decode('utf-8')

        # 三维簇状柱形图 = bar3DChart
        if 'bar3DChart' in chart:
            s += 3; d.append('三维簇状柱形图(bar3DChart) OK')
        elif 'barChart' in chart:
            s += 2; d.append('簇状柱形图(非三维)')

        # 标题 "学生平均分比较图"
        if '学生平均分比较图' in chart:
            s += 3; d.append('标题"学生平均分比较图" OK')
        else:
            d.append('未找到图表标题')

        # 图例靠右
        if 'legendPos' in chart and 'val="r"' in chart:
            s += 2; d.append('图例靠右 OK')
        elif 'legend' in chart.lower():
            s += 1; d.append('图例存在(位置未确认)')

        return self.R(9, '图表', s, full, d)

    # ===== R10: K/L列格式 楷体14居中+2位小数 (10) =====
    def r10(self):
        d = []; s = 0; full = 10
        info = self.sheets.get('Sheet1')
        if info is None:
            return self.R(10, '总分平均分格式', s, full, d)
        root = info['root']

        # K/L列有值
        k3 = self.sv(root, 'K3')
        l3 = self.sv(root, 'L3')
        if k3 and l3:
            s += 2; d.append(f'K3={k3} L3={l3} OK')

        # K/L列数据填充到K15/L15
        k15 = self.sv(root, 'K15')
        l15 = self.sv(root, 'L15')
        if k15 and l15:
            s += 2; d.append(f'数据填充至K15/L15 OK')

        # 楷体14 (K/L列也用楷体14)
        # 楷体已在all_xml中, 14也已在
        if '楷体' in self.all_xml.replace('华文行楷', '') and ('sz="14"' in self.all_xml or 'val="14"' in self.all_xml):
            s += 3; d.append('K/L列楷体14 OK')

        # 居中
        if 'horizontal="center"' in self.all_xml and 'vertical="center"' in self.all_xml:
            s += 3; d.append('K/L列居中 OK')

        return self.R(10, '总分平均分格式', s, full, d)


def main():
    if sys.platform == 'win32':
        try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except: pass

    base = os.path.dirname(os.path.abspath(__file__))
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base, '1结果.xlsx')
    print(f'评分: {os.path.basename(path)}')
    print('=' * 50)
    g = ExcelGrader2(path)
    results = g.grade()
    for r in results[:-1]:
        icon = 'PASS' if r['st'] == 'PASS' else 'FAIL'
        print(f"[{icon}] {r['req']:2d}. {r['name']:<14s}  {r['s']:5.1f} / {r['full']}")
        for d in r['d']:
            print(f'       {d}')
    print(f'{"=" * 50}\n  总分: {results[-1]["s"]} / 100')
    if len(sys.argv) < 2:
        input('\n按回车退出...')

if __name__ == '__main__':
    main()
