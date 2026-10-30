# -*- coding: utf-8 -*-
"""考点覆盖自检：题干里出现的考点，评分器有没有对应的检测代码。

用法：python tools/coverage_check.py [word|excel|ppt|all]
输出每个考点的「题干命中套数 / 是否有检测代码」，缺检测的会标出来。
"""
import glob
import os
import re
import sys
import xml.etree.ElementTree as ET
import zipfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# 考点 -> (题干关键词, 代码里应有的特征串)
POINTS = {
    'word': {
        '邮件合并': (['邮件合并'], r'MERGEFIELD|mailMerge'),
        '图片': (['图片', '图像'], r'pictures|w:drawing|extent|imagedata'),
        '环绕方式': (['环绕'], r'wrapSquare|wrapTopAndBottom|wrapNone'),
        '艺术字': (['艺术字'], r'textpath|prstTxWarp'),
        '文本框': (['文本框'], r'txbxContent'),
        '形状': (['形状', '椭圆', '箭头'], r'shape|oval|prstGeom'),
        'SmartArt': (['SmartArt'], r'diagram|dgm'),
        '公式': (['公式'], r'oMath|instrText'),
        '题注': (['题注'], r'caption|SEQ '),
        '水印': (['水印'], r'Watermark|textpath'),
        '脚注': (['脚注'], r'footnote'),
        '分栏': (['分栏'], r'cols'),
        '页面边框': (['页面边框'], r'pgBorders'),
        '页眉页脚': (['页眉', '页脚'], r'header|footer'),
        '页码': (['页码'], r'has_field|PAGE'),
        '超链接': (['超链接'], r'hyperlink'),
        '表格': (['表格'], r'w:tbl|tblPr'),
        '缩进行距': (['缩进', '行距'], r'ind|spacing'),
        '字符间距': (['字符间距'], r'spacing|kern'),
        '纸张页边距': (['纸张', '页边距'], r'pgSz|pgMar'),
    },
    'excel': {
        '图表': (['图表'], r'chart'),
        '数据透视表': (['数据透视表'], r'pivot'),
        '条件格式': (['条件格式'], r'conditionalFormatting'),
        '数据条': (['数据条'], r'dataBar'),
        '迷你图': (['迷你图'], r'sparkline'),
        '筛选': (['筛选'], r'hidden|autoFilter|filterColumn'),
        '排序': (['排序'], r'sortState|sortCondition'),
        '分类汇总': (['分类汇总'], r'SUBTOTAL|outlineLevel'),
        '合并单元格': (['合并'], r'mergeCell|centerContinuous'),
        '数字格式': (['小数', '货币', '百分比'], r'numFmt|formatCode'),
        '表样式': (['表样式'], r'tableStyleInfo|tableStyle'),
        '页面设置': (['页面设置', '纸张', '页边距'], r'pageSetup|pageMargins'),
        '页眉页脚': (['页眉', '页脚'], r'oddHeader|oddFooter'),
        '行高列宽': (['行高', '列宽'], r'trHeight|ht=|cols|bestFit'),
    },
    'ppt': {
        '主题': (['主题'], r'clrScheme|theme'),
        '版式': (['版式'], r'slideLayout|layout'),
        '动画': (['动画'], r'animEffect|timing'),
        '切换': (['切换'], r'transition|blinds|dissolve'),
        '放映': (['放映'], r'showPr|showType|browse'),
        '备注': (['备注'], r'notesSlide'),
        '图片': (['图片'], r'p:pic|blip|image'),
        '图片样式': (['图片样式', '柔化'], r'softEdge|prstGeom|effectLst'),
        '艺术字': (['艺术字'], r'prstTxWarp|gradFill'),
        '形状': (['形状'], r'spPr|prstGeom'),
        '超链接/动作': (['超链接', '动作按钮'], r'hlinkClick|actionButton|sldJmp'),
        '编号项目符号': (['编号', '项目符号'], r'buAutoNum|buChar'),
        '页脚': (['页脚'], r'ftr|footer'),
        '行距': (['行距'], r'lnSpc|spcPct'),
    },
}


def question_text(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read('word/document.xml'))
    return ' '.join(''.join(t.text or '' for t in p.iter(W + 't')) for p in root.iter(W + 'p'))


def grader_source(subject):
    buf = []
    for f in glob.glob(os.path.join(BASE, 'graders', '%s_grader_*.py' % subject)):
        buf.append(open(f, encoding='utf-8', errors='replace').read())
    return '\n'.join(buf)


def main():
    which = (sys.argv[1] if len(sys.argv) > 1 else 'all').lower()
    subs = list(POINTS) if which == 'all' else [which]
    missing_total = 0
    for subject in subs:
        text = []
        for n in range(1, 11):
            p = os.path.join(BASE, 'materials', subject, str(n), '3题干.docx')
            if os.path.exists(p):
                text.append((n, question_text(p)))
        src = grader_source(subject)
        print('=' * 62)
        print('%s（%d 套题干，评分器 %d 字符）' % (subject, len(text), len(src)))
        print('%-14s %-8s %s' % ('考点', '题干', '检测代码'))
        for name, (kws, pattern) in POINTS[subject].items():
            hits = [n for n, t in text if any(k in t for k in kws)]
            has = re.search(pattern, src) is not None
            flag = ''
            if hits and not has:
                flag = '  ← 缺'
                missing_total += 1
            print('%-14s %-8s %s%s' % (name, '%d套' % len(hits) if hits else '-',
                                       '有' if has else '无', flag))
    print('=' * 62)
    print('题干考到但评分器没检测的考点：%d 个' % missing_total)


if __name__ == '__main__':
    main()
