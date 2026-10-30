"""Excel 2010 评分器 - 第8套（员工工资统计表）"""
import os
import re
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import (ExcelBook, Report, run, formula_has, as_float, close,
                   style_name_is)

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


class ExcelGrader8:
    def __init__(self, path):
        self.book = ExcelBook(path)

    def grade(self):
        report = Report('Excel 第8套 · 员工工资统计表')
        for fn in (self.r1, self.r2, self.r3, self.r4, self.r5,
                   self.r6, self.r7, self.r8, self.r9):
            report.add(*fn())
        return report

    def sheet(self, name):
        return self.book.sheet(name)

    def fmt(self, sheet, ref):
        return self.book.format_in(sheet.root, ref)

    # 1. 改名+插入行+标题 (10)
    def r1(self):
        notes, got = [], 0
        sheet = self.sheet('员工工资统计表')
        if sheet is None:
            return (1, '改名+插入行+标题', 0, 10, ['未找到"员工工资统计表"'])
        got += 3; notes.append('改名 OK')
        if self.book.value_in(sheet.root, 'A2') == '姓名':
            got += 3; notes.append('插入空行 OK')
        else:
            notes.append('A2=%r（期望 姓名）' % self.book.value_in(sheet.root, 'A2'))
        a1 = self.book.value_in(sheet.root, 'A1')
        if '2021' in a1 and '工资' in a1:
            got += 4; notes.append('A1 标题 OK')
        else:
            notes.append('A1=%r' % a1)
        return (1, '改名+插入行+标题', got, 10, notes)

    # 2. 标题+数据区格式 (12)
    def r2(self):
        notes, got = [], 0
        sheet = self.sheet('员工工资统计表')
        if sheet is None:
            return (2, '标题格式', 0, 12, ['缺工作表'])
        if 'A1:L1' in sheet.merged():
            got += 2; notes.append('A1:L1 合并 OK')
        fmt = self.fmt(sheet, 'A1')
        font = fmt.get('font', {})
        if font.get('name') == '微软雅黑':
            got += 1; notes.append('微软雅黑 OK')
        if font.get('bold'):
            got += 1; notes.append('加粗 OK')
        if font.get('size') == '18':
            got += 1; notes.append('字号18 OK')
        color = self.book.resolve_color(font.get('color'))
        if color == 'FF0000':
            got += 1; notes.append('红色 OK')
        else:
            notes.append('标题颜色=%s' % (color or '未设置'))
        row1 = [r for r in sheet.rows() if r.get('r') == '1']
        if row1 and close(row1[0].get('ht'), 25, 2):
            got += 1; notes.append('行高25 OK')
        align = fmt.get('alignment', {})
        if align.get('horizontal') == 'center' and align.get('vertical') == 'center':
            got += 1; notes.append('标题水平垂直居中 OK')
        f2 = self.fmt(sheet, 'A2').get('font', {})
        if f2.get('name') == '华文楷体' and f2.get('size') == '14':
            got += 2; notes.append('数据区 华文楷体14 OK')
        else:
            notes.append('数据区字体=%s/%s' % (f2.get('name'), f2.get('size')))
        a2 = self.fmt(sheet, 'A2').get('alignment', {})
        if a2.get('horizontal') == 'center' and a2.get('vertical') == 'center':
            got += 2; notes.append('数据区居中 OK')
        return (2, '标题格式', got, 12, notes)

    # 3. 编号判性别 (10)
    def r3(self):
        notes, got = [], 0
        sheet = self.sheet('员工工资统计表')
        if sheet is None:
            return (3, '编号判性别', 0, 10, ['缺工作表'])
        f = self.book.formula_in(sheet.root, 'C3')
        if formula_has(f, 'MID', 'RIGHT', 'LEFT') and formula_has(f, 'MOD', 'ISEVEN', 'ISODD'):
            got += 10; notes.append('性别公式: %s' % f)
        elif formula_has(f, 'MID', 'RIGHT'):
            got += 6; notes.append('取位方式可行但缺奇偶判断: %s' % f)
        else:
            notes.append('C3=%r（要求 IF+MID/MOD）' % f)
        return (3, '编号判性别', got, 10, notes)

    # 4. 职务 + 补贴 (12)
    def r4(self):
        notes, got = [], 0
        sheet = self.sheet('员工工资统计表')
        if sheet is None:
            return (4, '职务+补贴', 0, 12, ['缺工作表'])
        e3 = self.book.formula_in(sheet.root, 'E3')
        if formula_has(e3, 'IF') and '3200' in e3 and '4200' in e3 and '5200' in e3:
            got += 6; notes.append('职务 IF OK')
        elif formula_has(e3, 'IF') and '3200' in e3 and '4200' in e3 and '部门经理' in e3:
            got += 6; notes.append('职务 IF OK（部门经理为否则分支）')
        else:
            notes.append('E3=%r' % e3[:70])
        g3 = self.book.formula_in(sheet.root, 'G3')
        if formula_has(g3, 'IF') and all(k in g3 for k in ('1000', '1500', '2000')):
            got += 6; notes.append('补贴 IF OK')
        else:
            notes.append('G3=%r' % g3[:70])
        return (4, '职务+补贴', got, 12, notes)

    # 5. 生活补贴+应发/实发+备注 (12)
    def r5(self):
        notes, got = [], 0
        sheet = self.sheet('员工工资统计表')
        if sheet is None:
            return (5, '工资计算', 0, 12, ['缺工作表'])
        lives = [as_float(self.book.value_in(sheet.root, 'H%d' % r)) for r in range(3, 15)]
        if lives and all(v == 1500 for v in lives if v is not None):
            got += 2; notes.append('生活补贴 1500 OK')
        else:
            notes.append('生活补贴异常: %s' % [v for v in lives if v is not None][:4])
        i3 = self.book.formula_in(sheet.root, 'I3')
        if formula_has(i3, 'SUM', '+'):
            got += 3; notes.append('应发公式: %s' % i3)
        elif as_float(self.book.value_in(sheet.root, 'I3')) is not None:
            got += 2; notes.append('应发为粘贴值')
        k3 = self.book.formula_in(sheet.root, 'K3')
        if formula_has(k3, 'SUM', '+'):
            got += 3; notes.append('实发公式: %s' % k3)
        elif as_float(self.book.value_in(sheet.root, 'K3')) is not None:
            got += 2; notes.append('实发为粘贴值')
        l3 = self.book.formula_in(sheet.root, 'L3')
        if formula_has(l3, 'IF') and '350' in l3:
            got += 4; notes.append('备注 IF OK')
        else:
            notes.append('L3=%r' % l3[:60])
        return (5, '工资计算', got, 12, notes)

    # 6. 表样式+数据条 (10)
    def r6(self):
        notes, got = [], 0
        sheet = self.sheet('员工工资统计表')
        if sheet is None:
            return (6, '表样式+数据条', 0, 10, ['缺工作表'])
        style_hit = ''
        for part in self.book.tables():
            xml = self.book.read(part)
            style = re.search(r'tableStyleInfo name="([^"]+)"', xml)
            ref = re.search(r'<table [^>]*ref="([^"]+)"', xml)
            if style and (ref is None or ref.group(1).replace('$', '').startswith('A2')):
                style_hit = style.group(1)
        if style_hit == 'TableStyleLight2':
            got += 5; notes.append('表样式 浅色2 (A2起) OK')
        elif style_hit:
            got += 3; notes.append('表样式=%s（注意区域）' % style_hit)
        else:
            notes.append('未找到套用表样式')
        bar = re.search(r'<conditionalFormatting sqref="(K\d+:K\d+)"[^>]*>\s*<cfRule type="dataBar"',
                        self.book.read(sheet.part), re.S)
        if bar:
            got += 5; notes.append('实发工资列数据条 %s OK' % bar.group(1))
        else:
            notes.append('未找到 K 列数据条')
        return (6, '表样式+数据条', got, 10, notes)

    # 7. 筛选 (14)
    def r7(self):
        notes, got = [], 0
        sheet = self.sheet('筛选')
        if sheet is None:
            return (7, '高级筛选', 0, 14, ['未找到"筛选"'])
        got += 2; notes.append('改名筛选 OK')
        a4 = self.book.value_in(sheet.root, 'A4')
        if a4 == '姓名':
            got += 2; notes.append('数据从 A4 开始 OK')
        else:
            notes.append('A4=%r（期望 姓名）' % a4)
        if self.book.value_in(sheet.root, 'A1') == '奖金' and '350' in self.book.value_in(sheet.root, 'A2'):
            got += 2; notes.append('条件 奖金>=350 OK')
        dept = self.book.value_in(sheet.root, 'B2')
        if dept == '技术' or self.book.formula_in(sheet.root, 'B2') == '"技术"':
            got += 2; notes.append('条件 部门=技术 OK')
        else:
            notes.append('B2=%r' % dept)
        hidden = sum(1 for row in sheet.rows() if row.get('hidden') == '1')
        if hidden:
            got += 4; notes.append('已筛选（%d 行隐藏）' % hidden)
        else:
            notes.append('没有隐藏行')
        cols = [c for c in sheet.root.iter('{%s}col' % S)]
        if cols and all(c.get('bestFit') == '1' or c.get('customWidth') == '1' for c in cols):
            got += 2; notes.append('A-L 自动列宽 OK')
        else:
            notes.append('列宽未全部调整')
        return (7, '高级筛选', got, 14, notes)

    # 8. 分类汇总 (12)
    def r8(self):
        notes, got = [], 0
        sheet = self.sheet('分类汇总')
        if sheet is None:
            return (8, '分类汇总', 0, 12, ['未找到"分类汇总"'])
        got += 2; notes.append('工作表 OK')
        conds = []
        for st in sheet.root.iter('{%s}sortState' % S):
            conds = [(sc.get('ref', ''), sc.get('descending')) for sc in st.iter('{%s}sortCondition' % S)]
        if conds and conds[0][0].startswith('D') and conds[0][1] != '1':
            got += 2; notes.append('部门升序 OK')
        else:
            notes.append('主排序=%s' % (conds[:1] or '无'))
        if len(conds) > 1 and conds[1][0].startswith('A') and conds[1][1] == '1':
            got += 2; notes.append('姓名降序 OK')
        subs = [node.findtext('{%s}f' % S) or '' for node in sheet.cells().values()]
        if any('SUBTOTAL(9' in f.upper() for f in subs):
            got += 3; notes.append('SUBTOTAL 求和 OK')
        elif any('SUBTOTAL' in f.upper() for f in subs):
            notes.append('有 SUBTOTAL 但不是求和')
        else:
            notes.append('没有汇总公式')
        total_formulas = sum(1 for f in subs if 'SUBTOTAL' in f.upper())
        if total_formulas >= 6:
            got += 1; notes.append('三项 × 多组汇总 OK')
        cols = [c for c in sheet.root.iter('{%s}col' % S)]
        if cols:
            got += 1; notes.append('列宽已调整 OK')
        align = self.fmt(sheet, 'A2').get('alignment', {})
        if align.get('horizontal') == 'center' and align.get('vertical') == 'center':
            got += 1; notes.append('水平垂直居中 OK')
        return (8, '分类汇总', got, 12, notes)

    # 9. 图表 (8)
    def r9(self):
        notes, got = [], 0
        sheet = self.sheet('图表')
        if sheet is None:
            return (9, '图表', 0, 8, ['未找到"图表"'])
        got += 2; notes.append('新建图表工作表 OK')
        charts = self.book.charts()
        if not charts:
            return (9, '图表', got, 8, notes + ['没有图表'])
        xml = self.book.read(charts[0])
        if 'bar3DChart' in xml:
            got += 1; notes.append('三维簇状柱形 OK')
        titles = re.findall(r'<a:t>([^<]+)</a:t>', xml)
        if any('员工工资对比图' in t for t in titles):
            got += 2; notes.append('标题在上方 OK')
        else:
            notes.append('标题=%s' % titles[:2])
        legend = re.search(r'<c:legendPos val="(\w+)"', xml)
        if legend and legend.group(1) == 'r':
            got += 1; notes.append('图例靠右 OK')
        refs = ' '.join(re.findall(r'<c:f>([^<]+)</c:f>', xml)).replace('$', '')
        ok_a = 'A3:A14' in refs
        ok_i = 'I3:I14' in refs
        ok_k = 'K3:K14' in refs
        if ok_a and ok_i and ok_k:
            got += 2; notes.append('系列=编号/应发/实发 OK')
        else:
            notes.append('数据区域=%s' % refs[:80])
        return (9, '图表', got, 8, notes)


if __name__ == '__main__':
    sys.exit(run(ExcelGrader8, default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        'Excel 2010 操作题题库', '8', '1结果.xlsx')))
