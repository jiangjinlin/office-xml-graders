"""Excel 2010 评分器 - 第10套（体检表）"""
import os
import re
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import (ExcelBook, Report, run, formula_has, style_name_is, close)

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


class ExcelGrader10:
    def __init__(self, path):
        self.book = ExcelBook(path)

    def grade(self):
        report = Report('Excel 第10套 · 体检表')
        for fn in (self.r1, self.r2, self.r3, self.r4, self.r5,
                   self.r6, self.r7, self.r8, self.r9, self.r10):
            report.add(*fn())
        return report

    def sheet(self, name):
        return self.book.sheet(name)

    def fmt(self, sheet, ref):
        return self.book.format_in(sheet.root, ref)

    # 1. 改名+区域格式 (10)
    def r1(self):
        notes, got = [], 0
        sheet = self.sheet('体检表')
        if sheet is None:
            return (1, '改名+区域格式', 0, 10, ['未找到"体检表"'])
        got += 3; notes.append('改名 OK')
        fmt = self.fmt(sheet, 'A1')
        font = fmt.get('font', {})
        if font.get('name') == '华文楷体' and font.get('size') == '12':
            got += 4; notes.append('华文楷体12 OK')
        else:
            notes.append('字体=%s/%s' % (font.get('name'), font.get('size')))
        align = fmt.get('alignment', {})
        if align.get('horizontal') == 'center' and align.get('vertical') == 'center':
            got += 3; notes.append('水平垂直居中 OK')
        end = self.fmt(sheet, 'G66')
        fe = end.get('font', {})
        if fe.get('name') == '华文楷体' and fe.get('size') == '12':
            got += 0
        else:
            notes.append('G66 字体=%s/%s（区域未全覆盖）' % (fe.get('name'), fe.get('size')))
            got -= 0
        return (1, '改名+区域格式', got, 10, notes)

    # 2. 出生日期 MID + 自动列宽 (10)
    def r2(self):
        notes, got = [], 0
        sheet = self.sheet('体检表')
        if sheet is None:
            return (2, '出生日期+列宽', 0, 10, ['缺工作表'])
        f = self.book.formula_in(sheet.root, 'B2')
        if formula_has(f, 'MID') and '年' in f and '月' in f and '日' in f:
            got += 7; notes.append('出生日期公式: %s' % f[:60])
        elif formula_has(f, 'MID'):
            got += 4; notes.append('有 MID 但缺年月日拼接: %s' % f[:50])
        else:
            v = self.book.value_in(sheet.root, 'B2')
            if re.match(r'\d{4}年\d{2}月\d{2}日', str(v)):
                got += 4; notes.append('出生日期为粘贴值 %s（缺公式）' % v)
            else:
                notes.append('B2=%r（要求 MID 拼 x年xx月xx日）' % v)
        cols = [c for c in sheet.root.iter('{%s}col' % S)]
        if cols and all(c.get('bestFit') == '1' or c.get('customWidth') == '1' for c in cols):
            got += 3; notes.append('A-G 自动列宽 OK')
        else:
            notes.append('列宽未调整')
        return (2, '出生日期+列宽', got, 10, notes)

    # 3. 性别 MID17 (10)
    def r3(self):
        notes, got = [], 0
        sheet = self.sheet('体检表')
        if sheet is None:
            return (3, '性别判断', 0, 10, ['缺工作表'])
        f = self.book.formula_in(sheet.root, 'C2')
        if formula_has(f, 'MID') and '17' in f and formula_has(f, 'MOD', 'ISEVEN', 'ISODD'):
            got += 10; notes.append('性别公式: %s' % f[:60])
        elif formula_has(f, 'MID'):
            got += 6; notes.append('有 MID 但缺第17位奇偶: %s' % f[:50])
        else:
            notes.append('C2=%r' % f)
        return (3, '性别判断', got, 10, notes)

    # 4. COUNTIF 人数 + 输出样式 (10)
    def r4(self):
        notes, got = [], 0
        sheet = self.sheet('体检表')
        if sheet is None:
            return (4, '人数统计', 0, 10, ['缺工作表'])
        i2 = self.book.formula_in(sheet.root, 'I2')
        i4 = self.book.formula_in(sheet.root, 'I4')
        if formula_has(i2, 'COUNTIF') and formula_has(i4, 'COUNTIF'):
            got += 4; notes.append('COUNTIF 男女 OK（I2/I4）')
        elif formula_has(i2, 'COUNTIF') or formula_has(i4, 'COUNTIF'):
            got += 2; notes.append('只有一个 COUNTIF')
        else:
            notes.append('I2=%r I4=%r' % (i2, i4))
        named = all(style_name_is(self.fmt(sheet, ref), '输出') for ref in ('I1', 'I2', 'I3', 'I4'))
        if named:
            got += 6; notes.append('I1:I4 样式"输出" OK')
        else:
            names = [self.fmt(sheet, ref).get('style_name') for ref in ('I1', 'I2', 'I3', 'I4')]
            notes.append('样式=%s' % names)
        return (4, '人数统计', got, 10, notes)

    # 5. 页眉页脚 (10)
    def r5(self):
        notes, got = [], 0
        sheet = self.sheet('体检表')
        if sheet is None:
            return (5, '页眉页脚', 0, 10, ['缺工作表'])
        hf = sheet.root.find('{%s}headerFooter' % S)
        if hf is None:
            return (5, '页眉页脚', 0, 10, ['没有页眉页脚'])
        header = hf.findtext('{%s}oddHeader' % S) or ''
        footer = hf.findtext('{%s}oddFooter' % S) or ''
        if header.startswith('&R') and '体检表' in header and '&10' in header:
            got += 4; notes.append('页眉靠右10号 OK')
        elif '体检表' in header:
            got += 3; notes.append('页眉内容对但格式=%r' % header)
        else:
            notes.append('页眉=%r' % header)
        if footer.startswith('&R') and '&P' in footer:
            got += 3; notes.append('页脚靠右页码 OK')
        else:
            notes.append('页脚=%r' % footer)
        ps = sheet.root.find('{%s}pageSetup' % S)
        if ps is not None and ps.get('useFirstPageNumber') == '1':
            got += 3; notes.append('起始页码=1 OK')
        else:
            notes.append('未设置起始页码')
        return (5, '页眉页脚', got, 10, notes)

    # 6. 边框 (10)
    def r6(self):
        notes, got = [], 0
        sheet = self.sheet('体检表')
        if sheet is None:
            return (6, '边框', 0, 10, ['缺工作表'])
        thick_red = double_blue = False
        for node in sheet.cells().values():
            fmt = self.book.styles.get(int(node.get('s', '0') or 0))
            for _side, style, color in [(k, v['style'], (v.get('color') or {}).get('rgb') or '')
                                        for k, v in (fmt.get('border') or {}).items() if v.get('style')]:
                if color.endswith('FF0000') and style == 'thick':
                    thick_red = True
                if color.endswith('0070C0') and style == 'double':
                    double_blue = True
        if thick_red:
            got += 5; notes.append('红色最粗外边框 OK')
        else:
            notes.append('未找到红色 thick 外框')
        if double_blue:
            got += 5; notes.append('蓝色双实线内边框 OK')
        else:
            notes.append('未找到蓝色 double 内框')
        return (6, '边框', got, 10, notes)

    # 7. 高级筛选 (12)
    def r7(self):
        notes, got = [], 0
        sheet = self.sheet('筛选')
        if sheet is None:
            return (7, '高级筛选', 0, 12, ['未找到"筛选"'])
        crit = [self.book.value_in(sheet.root, '%s%d' % (c, r)) for r in (1, 2) for c in 'DE']
        if crit[0] == '左眼视力' and crit[1] == '右眼视力':
            got += 3; notes.append('条件标题 OK')
        else:
            notes.append('条件标题=%s' % crit[:2])
        if '5.0' in crit[2] and '5.0' in crit[3]:
            got += 3; notes.append('条件 >=5.0 OK')
        else:
            notes.append('条件值=%s' % crit[2:])
        hidden = sum(1 for row in sheet.rows() if row.get('hidden') == '1')
        if hidden:
            got += 6; notes.append('已筛选（%d 行隐藏）' % hidden)
        else:
            notes.append('没有隐藏行（素材本身带筛选表结构，必须有隐藏行才给分）')
        return (7, '高级筛选', got, 12, notes)

    # 8. 筛选后排序+边框 (6)
    def r8(self):
        notes, got = [], 0
        sheet = self.sheet('筛选')
        if sheet is None:
            return (8, '排序+边框', 0, 6, ['缺"筛选"工作表'])
        conds = []
        for st in sheet.root.iter('{%s}sortState' % S):
            conds = [(sc.get('ref', ''), sc.get('descending')) for sc in st.iter('{%s}sortCondition' % S)]
        if conds and conds[0][0].startswith('A') and conds[0][1] != '1':
            got += 3; notes.append('姓名升序 OK')
        else:
            names = [self.book.value_in(sheet.root, 'A%d' % r) for r in range(4, 12)]
            names = [n for n in names if n]
            if names == sorted(names):
                got += 2; notes.append('数据已按姓名升序（无 sortState）')
            else:
                notes.append('未排序且无 sortState')
        if any(True for _ in [1]):
            pass
        has_double_dark = False
        for node in sheet.cells().values():
            fmt = self.book.styles.get(int(node.get('s', '0') or 0))
            for _side, style, color in [(k, v['style'], (v.get('color') or {}).get('rgb') or '')
                                        for k, v in (fmt.get('border') or {}).items() if v.get('style')]:
                if style == 'double' and color.endswith(('002060', '1F497D', '0070C0')):
                    has_double_dark = True
        if has_double_dark:
            got += 3; notes.append('深蓝双实线边框 OK')
        else:
            notes.append('未找到深蓝 double 边框')
        return (8, '排序+边框', got, 6, notes)

    # 9. 分类汇总 (12)
    def r9(self):
        notes, got = [], 0
        sheet = self.sheet('分类汇总')
        if sheet is None:
            return (9, '分类汇总', 0, 12, ['未找到"分类汇总"'])
        conds = []
        for st in sheet.root.iter('{%s}sortState' % S):
            conds = [(sc.get('ref', ''), sc.get('descending')) for sc in st.iter('{%s}sortCondition' % S)]
        if conds and conds[0][0].startswith('C') and conds[0][1] == '1':
            got += 3; notes.append('性别降序 OK')
        else:
            notes.append('主排序=%s' % (conds[:1] or '无'))
        if len(conds) > 1 and conds[1][0].startswith('A') and conds[1][1] == '1':
            got += 3; notes.append('姓名降序 OK')
        subs = [node.findtext('{%s}f' % S) or '' for node in sheet.cells().values()]
        if any('SUBTOTAL(1' in f.upper() for f in subs):
            got += 4; notes.append('SUBTOTAL 平均值 OK')
        elif any('SUBTOTAL' in f.upper() for f in subs):
            notes.append('有 SUBTOTAL 但不是平均值')
        else:
            notes.append('没有汇总公式')
        cols = [c for c in sheet.root.iter('{%s}col' % S)]
        if cols:
            got += 2; notes.append('保留源列宽 OK')
        else:
            notes.append('列宽未保留')
        return (9, '分类汇总', got, 12, notes)

    # 10. SmartArt (10)
    def r10(self):
        notes, got = [], 0
        sheet = self.sheet('体检流程表')
        if sheet is None:
            return (10, 'SmartArt流程图', 0, 10, ['未找到"体检流程表"'])
        got += 3; notes.append('工作表 OK')
        data = [n for n in self.book.names if n.startswith('xl/diagrams/data')]
        colors = [n for n in self.book.names if n.startswith('xl/diagrams/colors')]
        style = [n for n in self.book.names if n.startswith('xl/diagrams/quickStyle')]
        if not data:
            return (10, 'SmartArt流程图', got, 10, notes + ['没有 SmartArt 数据'])
        xml = self.book.read(data[0])
        texts = re.findall(r'<a:t>([^<]+)</a:t>', xml)
        need = ['登记领表', '排队体检', '交表，健康咨询']
        hits = sum(1 for n in need if any(n in t for t in texts))
        if hits == 3:
            got += 4; notes.append('流程内容 3/3 OK')
        else:
            got += hits; notes.append('流程内容 %d/3 OK' % hits)
        if colors:
            cx = self.book.read(colors[0])
            if re.search(r'colorful4|accent4|accent5', cx, re.I):
                got += 2; notes.append('配色=彩色范围-强调文字颜色4-5 OK')
            else:
                notes.append('配色未识别（可能自定义）')
        if style:
            sx = self.book.read(style[0])
            if '3d' in sx.lower():
                got += 1; notes.append('三维样式 OK')
            else:
                notes.append('quickStyle 非 3d')
        return (10, 'SmartArt流程图', got, 10, notes)


if __name__ == '__main__':
    sys.exit(run(ExcelGrader10, default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        'Excel 2010 操作题题库', '10', '1结果.xlsx')))
