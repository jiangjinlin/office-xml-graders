"""Excel 2010 评分器 - 第7套（期末成绩表）"""
import os
import re
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import (ExcelBook, Report, run, formula_has, func_name,
                   style_name_is, as_float, close, font_color, fill_color,
                   has_border)

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


class ExcelGrader7:
    def __init__(self, path):
        self.book = ExcelBook(path)

    def grade(self):
        report = Report('Excel 第7套 · 期末成绩表')
        for fn in (self.r1, self.r2, self.r3, self.r4, self.r5,
                   self.r6, self.r7, self.r8, self.r9, self.r10):
            report.add(*fn())
        return report

    def sheet(self, name):
        return self.book.sheet(name)

    def fmt(self, sheet, ref):
        return self.book.format_in(sheet.root, ref)

    # 1. 改名+合并+标题格式+数据区格式 (12)
    def r1(self):
        notes, got = [], 0
        sheet = self.sheet('期末成绩表')
        if sheet is None:
            return (1, '改名+标题格式', 0, 12, ['未找到"期末成绩表"'])
        got += 2; notes.append('改名 OK')
        if 'A1:K1' in sheet.merged():
            got += 2; notes.append('A1:K1 合并 OK')
        fmt = self.fmt(sheet, 'A1')
        font = fmt.get('font', {})
        if font.get('name') == '微软雅黑':
            got += 1; notes.append('微软雅黑 OK')
        if font.get('bold'):
            got += 1; notes.append('加粗 OK')
        if font.get('size') == '16':
            got += 1; notes.append('字号16 OK')
        if font_color(fmt) == 'FF0000':
            got += 1; notes.append('红色 OK')
        row1 = [r for r in sheet.rows() if r.get('r') == '1']
        if row1 and close(row1[0].get('ht'), 30, 2):
            got += 1; notes.append('行高30 OK')
        if fmt.get('alignment', {}).get('vertical') == 'center':
            got += 1; notes.append('垂直居中 OK')
        fmt2 = self.fmt(sheet, 'A3')
        f2 = fmt2.get('font', {})
        if f2.get('name') == '华文楷体' and f2.get('size') == '12':
            got += 1; notes.append('数据区 华文楷体12 OK')
        else:
            notes.append('数据区字体=%s/%s' % (f2.get('name'), f2.get('size')))
        al2 = fmt2.get('alignment', {})
        if al2.get('horizontal') == 'center' and al2.get('vertical') == 'center':
            got += 1; notes.append('数据区居中 OK')
        return (1, '改名+标题格式', got, 12, notes)

    # 2. VLOOKUP 姓名 (10)
    def r2(self):
        notes, got = [], 0
        sheet = self.sheet('期末成绩表')
        if sheet is None:
            return (2, 'VLOOKUP姓名', 0, 10, ['缺工作表'])
        f = self.book.formula_in(sheet.root, 'B3')
        if formula_has(f, 'VLOOKUP', 'INDEX', 'XLOOKUP') and '学生信息表' in f:
            got += 10; notes.append('查找公式: %s' % f)
        elif formula_has(f, 'VLOOKUP', 'INDEX'):
            got += 7; notes.append('有查找函数但未引用学生信息表: %s' % f)
        else:
            notes.append('B3=%r（要求 VLOOKUP 等查找函数）' % f)
        return (2, 'VLOOKUP姓名', got, 10, notes)

    # 3. 总分 + 各科平均两位小数 (10)
    def r3(self):
        notes, got = [], 0
        sheet = self.sheet('期末成绩表')
        if sheet is None:
            return (3, '总分+平均值', 0, 10, ['缺工作表'])
        j3 = self.book.formula_in(sheet.root, 'J3')
        if formula_has(j3, 'SUM', 'SUBTOTAL', '+'):
            got += 4; notes.append('总分公式: %s' % j3)
        elif as_float(self.book.value_in(sheet.root, 'J3')) is not None:
            got += 3; notes.append('总分为粘贴值（公式更稳）')
        else:
            notes.append('J3 无总分值')

        avg_formula = any(formula_has(self.book.formula_in(sheet.root, '%s47' % c),
                                      'AVERAGE', 'ROUND(', 'SUM(')
                          for c in 'CDEFGHI')
        filled = 0
        formatted = 0
        for c in 'CDEFGHI':
            ref = '%s47' % c
            if as_float(self.book.value_in(sheet.root, ref)) is not None:
                filled += 1
                if str(self.fmt(sheet, ref).get('numfmt', '')).startswith('0.00'):
                    formatted += 1
        if avg_formula:
            got += 2; notes.append('平均值使用 AVERAGE 公式 OK')
        else:
            notes.append('平均值行为粘贴值（缺公式）')
        if filled == 7 and formatted == 7:
            got += 4; notes.append('七科平均且两位小数 OK')
        elif filled:
            got += 1; notes.append('平均值 %d/7 格，%d/7 两位小数' % (filled, formatted))
        else:
            notes.append('第47行没有平均值')
        return (3, '总分+平均值', got, 10, notes)

    # 4. RANK 名次 (10)
    def r4(self):
        notes, got = [], 0
        sheet = self.sheet('期末成绩表')
        if sheet is None:
            return (4, '总分名次', 0, 10, ['缺工作表'])
        k3 = self.book.formula_in(sheet.root, 'K3')
        if formula_has(k3, 'RANK', 'RANK.EQ', 'RANKX') and '第' in k3 and '名' in k3:
            got += 10; notes.append('名次公式: %s' % k3)
        elif formula_has(k3, 'RANK', 'RANK.EQ'):
            got += 8; notes.append('有 RANK 但缺"第/名"拼接: %s' % k3)
        else:
            v = self.book.value_in(sheet.root, 'K3')
            notes.append('K3=%r（要求 第N名）' % v)
            if re.match(r'第\d+名', str(v)):
                got += 5; notes.append('结果是名次文本（缺 RANK 公式）')
        return (4, '总分名次', got, 10, notes)

    # 5. 条件格式：最高分 + 前15名 (12)
    def r5(self):
        notes, got = [], 0
        sheet = self.sheet('期末成绩表')
        if sheet is None:
            return (5, '条件格式', 0, 12, ['缺工作表'])
        st = self.book.read('xl/styles.xml')
        dxf_m = re.search(r'<dxfs[^>]*>(.*?)</dxfs>', st, re.S)
        dxfs = re.findall(r'<dxf>.*?</dxf>', dxf_m.group(1), re.S) if dxf_m else []

        def dxf_summary(idx):
            if not dxfs or int(idx or -1) >= len(dxfs):
                return ''
            xml = dxfs[int(idx)]
            color = re.search(r'<color rgb="([0-9A-Fa-f]+)"', xml)
            fill = re.search(r'<(?:bg|fg)Color rgb="([0-9A-Fa-f]+)"', xml)
            return (color.group(1)[-6:].upper() if color else '',
                    fill.group(1)[-6:].upper() if fill else '')

        rules = []
        for cf in sheet.root.iter('{%s}conditionalFormatting' % S):
            for rule in cf.iter('{%s}cfRule' % S):
                rules.append((cf.get('sqref', ''), rule.get('rank', ''), dxf_summary(rule.get('dxfId'))))
        top1 = [r for r in rules if r[1] == '1' and r[2] and not r[0].startswith('J')]
        top15 = [r for r in rules if r[1] == '15']
        if top1 and top1[0][2] == ('C00000', '00B0F0'):
            got += 5; notes.append('最高分：深红文本+浅蓝填充 OK')
        elif top1:
            got += 4; notes.append('最高分规则存在，颜色=%s' % (top1[0][2],))
        else:
            notes.append('未找到各科最高分规则（rank=1）')
        if top15 and top15[0][2] and top15[0][2][0] == '7030A0':
            got += 5; notes.append('前15名：紫色文本 OK')
        elif top15:
            got += 4; notes.append('前15名规则存在，颜色=%s' % (top15[0][2],))
        else:
            notes.append('未找到总分前15名规则（rank=15）')
        if top15 and top15[0][2] and top15[0][2][1] == 'FFC000':
            got += 2; notes.append('前15名橙色填充 OK')
        if not any(r[1] == '15' for r in rules) and not top1:
            notes.append('都没有 → 条件格式未实现')
        return (5, '条件格式', got, 12, notes)

    # 6. 高级筛选 (12)
    def r6(self):
        notes, got = [], 0
        sheet = self.sheet('筛选')
        if sheet is None:
            return (6, '高级筛选', 0, 12, ['未找到"筛选"工作表'])
        got += 2; notes.append('改名筛选 OK')
        if self.book.value_in(sheet.root, 'D1') == '数学':
            got += 2; notes.append('条件"数学" OK')
        if '100' in self.book.value_in(sheet.root, 'D2'):
            got += 2; notes.append('条件>=100 OK')
        hidden = sum(1 for row in sheet.rows() if row.get('hidden') == '1')
        if hidden:
            got += 2; notes.append('已筛选（%d 行隐藏）' % hidden)
        else:
            notes.append('没有隐藏行')
        thin_red = double_red = False
        for node in sheet.cells().values():
            fmt = self.book.styles.get(int(node.get('s', '0') or 0))
            for _side, style, color in [(k, v['style'], (v.get('color') or {}).get('rgb') or '')
                                        for k, v in (fmt.get('border') or {}).items() if v.get('style')]:
                if color.endswith('FF0000'):
                    if style == 'thin':
                        thin_red = True
                    if style == 'double':
                        double_red = True
        if thin_red:
            got += 2; notes.append('红色单实线外框 OK')
        else:
            notes.append('未找到红色单实线')
        if double_red:
            got += 2; notes.append('红色双实线内框 OK')
        else:
            notes.append('未找到红色双实线')
        return (6, '高级筛选', got, 12, notes)

    # 7. 性别 IF+MID (10)
    def r7(self):
        notes, got = [], 0
        sheet = self.sheet('学生信息表')
        if sheet is None:
            return (7, '身份证判性别', 0, 10, ['未找到"学生信息表"'])
        f = self.book.formula_in(sheet.root, 'D2')
        if formula_has(f, 'MID', 'LEFT', 'MIDB') and '17' in f:
            got += 10; notes.append('性别公式: %s' % f)
        elif formula_has(f, 'MID'):
            got += 8; notes.append('有 MID 但没取第17位: %s' % f)
        else:
            notes.append('D2=%r（要求 IF+MID）' % f)
        return (7, '身份证判性别', got, 10, notes)

    # 8. 图表 (8)
    def r8(self):
        notes, got = [], 0
        charts = self.book.charts()
        if not charts:
            return (8, '图表', 0, 8, ['没有图表'])
        xml = self.book.read(charts[0])
        if 'bar3DChart' in xml:
            got += 3; notes.append('三维簇状柱形 OK')
        legend = re.search(r'<c:legendPos val="(\w+)"', xml)
        if legend and legend.group(1) == 'r':
            got += 2; notes.append('图例靠右 OK')
        refs = ' '.join(re.findall(r'<c:f>([^<]+)</c:f>', xml)).replace('$', '')
        if 'B3:B8' in refs and 'J3:J8' in refs:
            got += 2; notes.append('数据区域 B/J OK')
        else:
            notes.append('数据区域=%s' % refs[:60])
            return (8, '图表', got, 8, notes)
        sheet = self.sheet('期末成绩表')
        for rel_type, target in self.book.rels(sheet.part).values():
            if rel_type.endswith('/drawing'):
                anchor = self.book.read(target)
                cols = [int(c) for c in re.findall(r'<xdr:col>(\d+)</xdr:col>', anchor)]
                if any(c >= 11 for c in cols):
                    got += 1; notes.append('图表位于 M 列以后 OK')
                else:
                    notes.append('图表锚点列=%s' % cols[:4])
                break
        return (8, '图表', got, 8, notes)

    # 9. 单元格样式 (8)
    def r9(self):
        notes, got = [], 0
        sheet = self.sheet('学生信息表')
        if sheet is None:
            return (9, '单元格样式', 0, 8, ['缺工作表'])
        s1 = self.fmt(sheet, 'A1')
        if style_name_is(s1, '标题 1', '标题1'):
            got += 3; notes.append('A1:E1 标题1 OK')
        else:
            notes.append('A1 样式=%r' % (s1.get('style_name') or '无'))
        s2 = self.fmt(sheet, 'A2')
        if style_name_is(s2, '40% - 强调文字颜色 1', '40%-强调文字颜色1'):
            got += 3; notes.append('A2:E56 40%强调色1 OK')
        else:
            notes.append('A2 样式=%r' % (s2.get('style_name') or '无'))
        # 自动列宽：按覆盖列数统计（min..max 可能合并），限定 A-E
        covered = 0
        for c in sheet.root.iter('{%s}col' % S):
            if c.get('bestFit') == '1' or c.get('customWidth') == '1':
                lo, hi = int(c.get('min', '99')), int(c.get('max', '0'))
                covered += max(0, min(hi, 5) - max(lo, 1) + 1)
        if covered >= 5:
            got += 2; notes.append('A-E 自动列宽 OK')
        else:
            notes.append('A-E 列宽未调整（覆盖 %d 列）' % covered)
        return (9, '单元格样式', got, 8, notes)

    # 10. 页眉页脚 (8)
    def r10(self):
        notes, got = [], 0
        sheet = self.sheet('学生信息表')
        if sheet is None:
            return (10, '页眉页脚', 0, 8, ['缺工作表'])
        hf = sheet.root.find('{%s}headerFooter' % S)
        if hf is None:
            return (10, '页眉页脚', 0, 8, ['没有页眉页脚'])
        header = hf.findtext('{%s}oddHeader' % S) or ''
        footer = hf.findtext('{%s}oddFooter' % S) or ''
        if '学生信息' in header and header.startswith('&C'):
            got += 4; notes.append('页眉居中"学生信息" OK')
        elif '学生信息' in header:
            got += 3; notes.append('页眉内容对但位置=%r' % header)
        else:
            notes.append('页眉=%r' % header)
        if '&R&P' in footer or '&P' in footer.replace('&R', ''):
            got += 4; notes.append('页脚靠右页码 OK')
        else:
            notes.append('页脚=%r' % footer)
        return (10, '页眉页脚', got, 8, notes)


if __name__ == '__main__':
    sys.exit(run(ExcelGrader7, default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        '2022年准易：Excel 2010 操作题--素材', '7', '1结果.xlsx')))
