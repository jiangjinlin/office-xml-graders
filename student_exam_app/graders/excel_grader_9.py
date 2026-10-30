"""Excel 2010 评分器 - 第9套（商品销售情况统计表）"""
import os
import re
import sys
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import (ExcelBook, Report, run, formula_has, close, has_border)

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


class ExcelGrader9:
    def __init__(self, path):
        self.book = ExcelBook(path)

    def grade(self):
        report = Report('Excel 第9套 · 商品销售情况统计表')
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
        sheet = self.sheet('商品销售情况统计表')
        if sheet is None:
            return (1, '改名+插入行+标题', 0, 10, ['未找到工作表'])
        got += 3; notes.append('改名 OK')
        if self.book.value_in(sheet.root, 'A2') == '型号':
            got += 3; notes.append('插入空行 OK')
        else:
            notes.append('A2=%r（期望 型号）' % self.book.value_in(sheet.root, 'A2'))
        a1 = self.book.value_in(sheet.root, 'A1')
        if '2020' in a1 and '商品销售' in a1:
            got += 4; notes.append('A1 标题 OK')
        else:
            notes.append('A1=%r' % a1)
        return (1, '改名+插入行+标题', got, 10, notes)

    # 2. 合并+格式+列宽 (14)
    def r2(self):
        notes, got = [], 0
        sheet = self.sheet('商品销售情况统计表')
        if sheet is None:
            return (2, '标题格式+列宽', 0, 14, ['缺工作表'])
        if 'A1:F1' in sheet.merged():
            got += 2; notes.append('A1:F1 合并 OK')
        fmt = self.fmt(sheet, 'A1')
        font = fmt.get('font', {})
        if font.get('name') == '微软雅黑':
            got += 1; notes.append('微软雅黑 OK')
        if font.get('bold'):
            got += 1; notes.append('加粗 OK')
        if font.get('size') == '14':
            got += 1; notes.append('字号14 OK')
        color = self.book.resolve_color(font.get('color'))
        if color == '0070C0':
            got += 1; notes.append('蓝色 OK')
        else:
            notes.append('标题颜色=%s（要求 0070C0）' % (color or '未设置'))
        row1 = [r for r in sheet.rows() if r.get('r') == '1']
        if row1 and close(row1[0].get('ht'), 28, 2):
            got += 1; notes.append('行高28 OK')
        if fmt.get('alignment', {}).get('vertical') == 'center':
            got += 1; notes.append('标题垂直居中 OK')
        f2 = self.fmt(sheet, 'A2').get('font', {})
        if f2.get('name') == '华文中宋' and f2.get('size') == '12':
            got += 2; notes.append('数据区 华文中宋12 OK')
        else:
            notes.append('数据区字体=%s/%s' % (f2.get('name'), f2.get('size')))
        cols = list(sheet.root.iter('{%s}col' % S))
        wide = [c for c in cols
                if c.get('min') in ('1', '2', '3', '4', '5', '6')
                and close(float(c.get('width', '0')), 15, 1.5)]
        if wide:
            got += 2; notes.append('A-F 列宽15 OK')
        else:
            notes.append('A-F 列宽=%s' % [(c.get('min'), c.get('width')) for c in cols[:3]])
        align2 = self.fmt(sheet, 'A2').get('alignment', {})
        if align2.get('vertical') == 'center':
            got += 2; notes.append('数据区垂直居中 OK')
        return (2, '标题格式+列宽', got, 14, notes)

    # 3. 总价公式 + 页面居中 (10)
    def r3(self):
        notes, got = [], 0
        sheet = self.sheet('商品销售情况统计表')
        if sheet is None:
            return (3, '总价+页面设置', 0, 10, ['缺工作表'])
        f3 = self.book.formula_in(sheet.root, 'F3')
        if formula_has(f3, '*', 'PRODUCT'):
            got += 5; notes.append('总价公式: %s' % f3)
        else:
            notes.append('F3=%r' % f3)
        po = sheet.root.find('{%s}printOptions' % S)
        h = po is not None and po.get('horizontalCentered') == '1'
        v = po is not None and po.get('verticalCentered') == '1'
        if h and v:
            got += 5; notes.append('页面水平垂直居中 OK')
        else:
            notes.append('打印居中=%s/%s' % (h, v))
        return (3, '总价+页面设置', got, 10, notes)

    # 4. MID 生产年份 (10)
    def r4(self):
        notes, got = [], 0
        sheet = self.sheet('商品销售情况统计表')
        if sheet is None:
            return (4, 'MID年份', 0, 10, ['缺工作表'])
        f = self.book.formula_in(sheet.root, 'B3')
        if formula_has(f, 'MID') and ('3' in f and '4' in f):
            got += 10; notes.append('年份公式: %s' % f)
        elif formula_has(f, 'MID'):
            got += 7; notes.append('有 MID 但取位参数可疑: %s' % f)
        else:
            notes.append('B3=%r（要求 MID）' % f)
        return (4, 'MID年份', got, 10, notes)

    # 5. 页眉页脚+起始页码 (10)
    def r5(self):
        notes, got = [], 0
        sheet = self.sheet('商品销售情况统计表')
        if sheet is None:
            return (5, '页眉页脚', 0, 10, ['缺工作表'])
        hf = sheet.root.find('{%s}headerFooter' % S)
        if hf is None:
            return (5, '页眉页脚', 0, 10, ['没有页眉页脚'])
        header = hf.findtext('{%s}oddHeader' % S) or ''
        footer = hf.findtext('{%s}oddFooter' % S) or ''
        if header.startswith('&R') and '商品销售情况统计表' in header and '&10' in header:
            got += 4; notes.append('页眉靠右10号 OK')
        elif '商品销售情况统计表' in header:
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

    # 6. 边框 (12)
    def r6(self):
        notes, got = [], 0
        sheet = self.sheet('商品销售情况统计表')
        if sheet is None:
            return (6, '边框', 0, 12, ['缺工作表'])
        thick_red = double_purple = False
        for node in sheet.cells().values():
            fmt = self.book.styles.get(int(node.get('s', '0') or 0))
            for _side, style, color in [(k, v['style'], (v.get('color') or {}).get('rgb') or '')
                                        for k, v in (fmt.get('border') or {}).items() if v.get('style')]:
                if color.endswith('FF0000') and style == 'thick':
                    thick_red = True
                if color.endswith('7030A0') and style == 'double':
                    double_purple = True
        if thick_red:
            got += 6; notes.append('红色最粗外边框 OK')
        else:
            notes.append('未找到红色 thick 外边框')
        if double_purple:
            got += 6; notes.append('紫色双实线内边框 OK')
        else:
            notes.append('未找到紫色 double 内边框')
        return (6, '边框', got, 12, notes)

    # 7. 数据透视表 (12)
    def r7(self):
        notes, got = [], 0
        sheet = self.sheet('商品销售情况统计表')
        if sheet is None:
            return (7, '数据透视表', 0, 12, ['缺工作表'])
        pivots = self.book.pivot_tables()
        if not pivots:
            return (7, '数据透视表', 0, 12, ['没有透视表'])
        xml = self.book.read(pivots[0])
        loc = re.search(r'<location ref="([A-Z]+\d+:[A-Z]+\d+)"', xml)
        if loc and loc.group(1) == 'H1:K5':
            got += 3; notes.append('位置 H1:K5 OK')
        elif loc:
            got += 1; notes.append('位置=%s（要求 H1:K5）' % loc.group(1))
        else:
            notes.append('没有 location')
        cache = {}
        for part in self.book.parts(self.book._base + 'pivotCache/', exclude=('_rels/',)):
            cache = dict(enumerate(re.findall(r'<cacheField name="([^"]+)"', self.book.read(part))))
            break
        row_block = re.search(r'<rowFields.*?</rowFields>', xml, re.S)
        col_block = re.search(r'<colFields.*?</colFields>', xml, re.S)
        rows = [cache.get(int(i), '?') for i in re.findall(r'<field x="(\d+)"', row_block.group(0))] if row_block else []
        cols = [cache.get(int(i), '?') for i in re.findall(r'<field x="(\d+)"', col_block.group(0))] if col_block else []
        if rows == ['生产年份']:
            got += 2; notes.append('行标签=生产年份 OK')
        else:
            notes.append('行标签=%s' % rows)
        if cols == ['型号']:
            got += 2; notes.append('列标签=型号 OK')
        else:
            notes.append('列标签=%s' % cols)
        data = re.findall(r'<dataField name="([^"]+)"', xml)
        if data and '总价' in data[0]:
            got += 2; notes.append('求和项=%s OK' % data[0])
        else:
            notes.append('数据项=%s' % data)
        data_area = re.search(r'<pivotTableStyleInfo|<cacheSource[^>]*worksheetSource ref="([^"]+)"', xml)
        if data_area:
            got += 1; notes.append('数据源 OK')
        if loc and loc.group(1) == 'H1:K5' and rows == ['生产年份'] and cols == ['型号'] and data:
            pass
        if self.book.value_in(sheet.root, 'H1'):
            got += 2; notes.append('透视表落在主表 H1 OK')
        else:
            notes.append('主表 H1 无透视表内容')
        return (7, '数据透视表', got, 12, notes)

    # 8. 分类汇总 (12)
    def r8(self):
        notes, got = [], 0
        sheet = self.sheet('分类汇总')
        if sheet is None:
            return (8, '分类汇总', 0, 12, ['未找到"分类汇总"'])
        conds = []
        for st in sheet.root.iter('{%s}sortState' % S):
            conds = [(sc.get('ref', ''), sc.get('descending')) for sc in st.iter('{%s}sortCondition' % S)]
        if conds and conds[0][0].startswith('C') and conds[0][1] != '1':
            got += 3; notes.append('类别升序 OK')
        else:
            notes.append('主排序=%s' % (conds[:1] or '无'))
        if len(conds) > 1 and conds[1][0].startswith('A') and conds[1][1] != '1':
            got += 2; notes.append('型号升序 OK')
        subs = [node.findtext('{%s}f' % S) or '' for node in sheet.cells().values()]
        if any('SUBTOTAL(9' in f.upper() for f in subs):
            got += 4; notes.append('SUBTOTAL 求和 OK')
        elif any('SUBTOTAL' in f.upper() for f in subs):
            notes.append('有 SUBTOTAL 但不是求和')
        else:
            notes.append('没有汇总公式')
        total = sum(1 for f in subs if 'SUBTOTAL' in f.upper())
        if total >= 4:
            got += 1; notes.append('数量/总价两组汇总 OK')
        outline = any(row.get('outlineLevel') for row in sheet.rows())
        if outline:
            got += 2; notes.append('分级显示 OK')
        cols = [c for c in sheet.root.iter('{%s}col' % S)
                if c.get('min') in ('1', '2', '3', '4', '5', '6')
                and close(float(c.get('width', '0')), 15, 1.5)]
        if cols:
            got += 2; notes.append('A-F 列宽15 OK')
        else:
            notes.append('列宽未设 15')
        return (8, '分类汇总', got, 12, notes)

    # 9. 图表 (10)
    def r9(self):
        notes, got = [], 0
        sheet = self.sheet('图表')
        if sheet is None:
            return (9, '图表', 0, 10, ['未找到"图表"'])
        got += 2; notes.append('图表工作表 OK')
        charts = self.book.charts()
        if not charts:
            return (9, '图表', got, 10, notes + ['没有图表'])
        xml = self.book.read(charts[0])
        if 'bar3DChart' in xml:
            got += 2; notes.append('三维簇状柱形 OK')
        title_m = re.search(r'<c:title>(.*?)</c:title>', xml, re.S)
        if title_m:
            t = title_m.group(1)
            ok_title = '单价对比图' in t
            ok_font = '楷体' in t
            ok_size = 'sz="1200"' in t
            ok_bold = 'b="1"' in t
            score = sum((ok_title, ok_font, ok_size, ok_bold))
            got += score
            notes.append('标题: 文字%s 楷体%s 12号%s 加粗%s' % tuple('OK' if x else 'NG' for x in (ok_title, ok_font, ok_size, ok_bold)))
        else:
            notes.append('没有图表标题')
        legend = re.search(r'<c:legendPos val="(\w+)"', xml)
        if legend and legend.group(1) == 'r':
            got += 1; notes.append('图例靠右 OK')
        refs = ' '.join(re.findall(r'<c:f>([^<]+)</c:f>', xml)).replace('$', '')
        if 'A3:A10' in refs and 'E3:E10' in refs:
            got += 1; notes.append('系列区域 OK')
        else:
            notes.append('数据区域=%s' % refs[:60])
        for rel_type, target in self.book.rels(sheet.part).values():
            if rel_type.endswith('/drawing'):
                got += 1; notes.append('图表放在"图表"工作表 OK')
                break
        return (9, '图表', got, 10, notes)


if __name__ == '__main__':
    sys.exit(run(ExcelGrader9, default=os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        'Excel 2010 操作题题库', '9', '1结果.xlsx')))
