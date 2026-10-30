"""Excel 2010 评分器 - 第6套（员工加班统计表）"""
import os
import re
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import ExcelBook, Report, run

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


class ExcelGrader6:
    def __init__(self, path):
        self.book = ExcelBook(path)

    def grade(self):
        report = Report('Excel 第6套 · 员工加班统计表')
        for fn in (self.r1, self.r2, self.r3, self.r4, self.r5,
                   self.r6, self.r7, self.r8, self.r9):
            report.add(*fn())
        return report

    def sheet(self, name):
        return self.book.sheet(name)

    def fmt(self, sheet, ref):
        return self.book.format_in(sheet.root, ref)

    # 1. 改名+插入行+标题 (8)
    def r1(self):
        notes, got = [], 0
        sheet = self.sheet('员工加班统计表')
        if sheet is None:
            return (1, '改名+插入行+标题', 0, 8, ['未找到"员工加班统计表"'])
        got += 3; notes.append('改名 OK')
        a1 = self.book.value_in(sheet.root, 'A1')
        if '各部门' in a1 and '加班' in a1:
            got += 2; notes.append('A1 标题 OK')
        else:
            notes.append('A1=%r' % a1)
        if self.book.value_in(sheet.root, 'A2') == '工号':
            got += 3; notes.append('插入空行 OK')
        else:
            notes.append('A2=%r（期望 工号）' % self.book.value_in(sheet.root, 'A2'))
        return (1, '改名+插入行+标题', got, 8, notes)

    # 2. 合并+楷体15+红色+垂直居中 (12)
    def r2(self):
        notes, got = [], 0
        sheet = self.sheet('员工加班统计表')
        if sheet is None:
            return (2, '标题格式', 0, 12, ['缺工作表'])
        if 'A1:H1' in sheet.merged():
            got += 3; notes.append('A1:H1 合并 OK')
        fmt = self.fmt(sheet, 'A1')
        font = fmt.get('font', {})
        if font.get('name') == '楷体':
            got += 3; notes.append('楷体 OK')
        else:
            notes.append('字体=%s' % font.get('name'))
        if font.get('size') == '15':
            got += 2; notes.append('字号15 OK')
        else:
            notes.append('字号=%s' % font.get('size'))
        if (font.get('color', {}).get('rgb') or '').upper() == 'FFFF0000':
            got += 2; notes.append('红色 OK')
        if fmt.get('alignment', {}).get('vertical') == 'center':
            got += 2; notes.append('垂直居中 OK')
        return (2, '标题格式', got, 12, notes)

    # 3. 工号补齐 + VLOOKUP 加班小时数 (12)
    def r3(self):
        notes, got = [], 0
        sheet = self.sheet('员工加班统计表')
        if sheet is None:
            return (3, '工号+VLOOKUP', 0, 12, ['缺工作表'])
        codes = [self.book.value_in(sheet.root, 'A%d' % r) for r in range(3, 22)]
        text_codes = [c for c in codes if c]
        if text_codes and all(re.fullmatch(r'\d{3}', c) for c in text_codes):
            got += 5; notes.append('工号三位补齐 OK（%s）' % text_codes[0])
        else:
            notes.append('工号格式异常: %s' % text_codes[:5])
        target = self.sheet('加班时间')
        formula = self.book.formula_in(sheet.root, 'E3')
        ref_ok = '加班时间' in formula or (target is not None and 'A1:B4' in formula.replace('$', ''))
        if 'VLOOKUP' in formula.upper() and ref_ok:
            got += 7; notes.append('VLOOKUP: %s' % formula)
        elif 'VLOOKUP' in formula.upper():
            got += 5; notes.append('VLOOKUP 存在但引用可疑: %s' % formula)
        else:
            notes.append('E3=%r' % formula)
        return (3, '工号+VLOOKUP', got, 12, notes)

    # 4. IF 加班工资 + SUM 合计 + 备注 (14)
    def r4(self):
        notes, got = [], 0
        sheet = self.sheet('员工加班统计表')
        if sheet is None:
            return (4, '工资计算', 0, 14, ['缺工作表'])
        f_f3 = self.book.formula_in(sheet.root, 'F3').upper()
        if 'IF' in f_f3 and all(x in f_f3 for x in ('30', '25', '20')):
            got += 6; notes.append('加班工资 IF 分档 OK')
        else:
            notes.append('F3=%r' % f_f3)
        g3 = self.book.formula_in(sheet.root, 'G3').upper()
        if 'SUM' in g3 or '+' in g3:
            got += 4; notes.append('合计 SUM OK')
        else:
            notes.append('G3=%r' % g3)
        h3 = self.book.formula_in(sheet.root, 'H3').upper()
        if 'IF' in h3 and '2500' in h3 and '优秀' in self.book.formula_in(sheet.root, 'H3'):
            got += 4; notes.append('备注 IF OK')
        else:
            notes.append('H3=%r' % h3)
        return (4, '工资计算', got, 14, notes)

    # 5. 表样式+字体+数据条 (14)
    def r5(self):
        notes, got = [], 0
        sheet = self.sheet('员工加班统计表')
        if sheet is None:
            return (5, '表样式+数据条', 0, 14, ['缺工作表'])
        table_style = ''
        for part in self.book.tables():
            hit = re.search(r'tableStyleInfo name="([^"]+)"', self.book.read(part))
            if hit:
                table_style = hit.group(1)
        if table_style == 'TableStyleLight5':
            got += 4; notes.append('表样式 浅色5 OK')
        else:
            notes.append('表样式=%r' % table_style)
        fmt = self.fmt(sheet, 'A3')
        if fmt.get('font', {}).get('name') == '华文中宋':
            got += 3; notes.append('华文中宋 OK')
        if fmt.get('font', {}).get('size') == '12':
            got += 2; notes.append('字号12 OK')
        align = fmt.get('alignment', {})
        if align.get('horizontal') == 'center' and align.get('vertical') == 'center':
            got += 2; notes.append('水平垂直居中 OK')
        else:
            notes.append('对齐=%s/%s' % (align.get('horizontal'), align.get('vertical')))
        bars = []
        for cf in sheet.root.iter('{%s}conditionalFormatting' % S):
            for rule in cf.iter('{%s}cfRule' % S):
                if rule.get('type') == 'dataBar':
                    bars.append(cf.get('sqref', ''))
        if any('G' in b for b in bars):
            got += 3; notes.append('数据条 %s OK' % bars[0])
        else:
            notes.append('未找到 G 列数据条（找到 %s）' % bars)
        return (5, '表样式+数据条', got, 14, notes)

    # 6. 复制值+排序+分类汇总 (12)
    def r6(self):
        notes, got = [], 0
        sheet = self.sheet('分类汇总')
        if sheet is None:
            return (6, '排序+分类汇总', 0, 12, ['未找到"分类汇总"'])
        got += 2; notes.append('工作表改名 OK')
        has_formula = any(
            (node.find('{%s}f' % S) is not None
             and 'SUBTOTAL' not in (node.find('{%s}f' % S).text or '').upper())
            for node in sheet.cells().values())
        if not has_formula:
            got += 2; notes.append('粘贴值(无公式) OK')
        else:
            notes.append('除汇总外仍含公式')
        conds = []
        for st in sheet.root.iter('{%s}sortState' % S):
            conds = [(sc.get('ref', ''), sc.get('descending')) for sc in st.iter('{%s}sortCondition' % S)]
        if conds and conds[0][0].startswith('C') and conds[0][1] != '1':
            got += 2; notes.append('部门升序 OK')
        else:
            notes.append('排序条件=%s' % (conds or '无'))
        if len(conds) > 1 and conds[1][0].startswith('B') and conds[1][1] == '1':
            got += 2; notes.append('姓名降序 OK')
        levels = {row.get('outlineLevel') for row in sheet.rows() if row.get('outlineLevel')}
        if levels:
            got += 2; notes.append('分级显示 OK')
        subs = [node.findtext('{%s}f' % S) or '' for node in sheet.cells().values()]
        avg = [f for f in subs if 'SUBTOTAL(1' in f.upper()]
        if avg:
            got += 2; notes.append('SUBTOTAL 平均值 OK（%s）' % avg[0])
        elif any('SUBTOTAL' in f.upper() for f in subs):
            notes.append('SUBTOTAL 存在但非平均值')
        two_dec = any(str(self.fmt(sheet, ref).get('numfmt', '')).startswith('0.00')
                      for ref in ('D8', 'G8', 'D16', 'G16', 'D23', 'G23'))
        if two_dec:
            got += 2; notes.append('平均值两位小数 OK')
        return (6, '排序+分类汇总', got, 12, notes)

    # 7. 筛选表 (10)
    def r7(self):
        notes, got = [], 0
        sheet = self.sheet('筛选')
        if sheet is None:
            return (7, '高级筛选', 0, 10, ['未找到"筛选"'])
        got += 2; notes.append('工作表 OK')
        if self.book.value_in(sheet.root, 'D1') == '基本工资':
            got += 2; notes.append('条件标题 OK')
        if '2500' in self.book.value_in(sheet.root, 'D2'):
            got += 2; notes.append('条件 >=2500 OK')
        hidden = sum(1 for row in sheet.rows() if row.get('hidden') == '1')
        if hidden:
            got += 4; notes.append('已筛选（%d 行隐藏）' % hidden)
        else:
            notes.append('没有隐藏行')
        return (7, '高级筛选', got, 10, notes)

    # 8. 图表 (10)
    def r8(self):
        notes, got = [], 0
        charts = self.book.charts()
        if not charts:
            return (8, '图表', 0, 10, ['没有图表'])
        xml = self.book.read(charts[0])
        if 'bar3DChart' in xml:
            got += 3; notes.append('三维簇状柱形 OK')
        titles = re.findall(r'<a:t>([^<]+)</a:t>', xml)
        if any('应发工资对比图' in t for t in titles):
            got += 3; notes.append('标题 OK')
        else:
            notes.append('标题=%s' % titles[:3])
        legend = re.search(r'<c:legendPos val="(\w+)"', xml)
        if legend and legend.group(1) == 'r':
            got += 2; notes.append('图例靠右 OK')
        refs = ' '.join(re.findall(r'<c:f>([^<]+)</c:f>', xml)).replace('$', '')
        if 'B3:B10' in refs and 'G3:G10' in refs:
            got += 2; notes.append('数据区域 OK')
        else:
            notes.append('数据区域=%s' % refs[:60])
        return (8, '图表', got, 10, notes)

    # 9. 数据透视表 (8)
    def r9(self):
        notes, got = [], 0
        sheet = self.sheet('数据透视表')
        if sheet is None:
            return (9, '数据透视表', 0, 8, ['未找到"数据透视表"'])
        got += 2; notes.append('工作表 OK')
        pivots = self.book.pivot_tables()
        if not pivots:
            return (9, '数据透视表', got, 8, notes + ['没有透视表'])
        cache = {}
        for part in self.book.parts(self.book._base + 'pivotCache/', exclude=('_rels/',)):
            fields = re.findall(r'<cacheField name="([^"]+)"', self.book.read(part))
            cache = dict(enumerate(fields))
            break
        xml = self.book.read(pivots[0])
        row_block = re.search(r'<rowFields.*?</rowFields>', xml, re.S)
        col_block = re.search(r'<colFields.*?</colFields>', xml, re.S)
        row_names = []
        if row_block:
            for idx in re.findall(r'<field x="(\d+)"', row_block.group(0)):
                row_names.append(cache.get(int(idx), '?'))
        col_names = []
        if col_block:
            for idx in re.findall(r'<field x="(\d+)"', col_block.group(0)):
                col_names.append(cache.get(int(idx), '?'))
        if col_names and col_names[0] == '部门':
            got += 2; notes.append('列标签=部门 OK')
        else:
            notes.append('列标签=%s' % col_names)
        if row_names and row_names[0] == '姓名':
            got += 2; notes.append('行标签=姓名 OK')
        else:
            notes.append('行标签=%s' % row_names)
        data = re.findall(r'<dataField name="([^"]+)"', xml)
        if data and '合计应发工资' in data[0]:
            got += 2; notes.append('求和项=%s OK' % data[0])
        else:
            notes.append('数据项=%s' % data)
        return (9, '数据透视表', got, 8, notes)


if __name__ == '__main__':
    sys.exit(run(ExcelGrader6, default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        'Excel 2010 操作题题库', '6', '1结果.xlsx')))
