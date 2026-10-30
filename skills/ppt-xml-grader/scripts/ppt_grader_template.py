#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PowerPoint 2010 操作题 XML 评分器 — 通用模板
使用此模板创建新的 PPT 评分器: cp this.py ppt_grader_N.py → 填充 req 函数 → 测试

核心知识:
  - 主题: theme1.xml name 属性
  - 版式: slide rels → slideLayout Target
  - 字体: a:rPr typeface 或 theme 默认字体
  - 动画: animEffect + nodeType (afterEffect/withEffect/clickEffect)
  - 切换: transition 子元素 + spd (slow/med/fast)
  - 放映: presProps.xml (browse/kiosk/present)
  - 页脚: ph type="ftr" → a:t 文本
  - 备注: notesSlides/notesSlideN.xml
"""
import zipfile, os, sys, io, re
from pathlib import Path
import xml.etree.ElementTree as ET

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

class PPTGrader:
    """PPT XML 底层评分器"""
    def __init__(self, path):
        self.zf = zipfile.ZipFile(path)
        self.nl = self.zf.namelist()
        self.slides = {}
        self.theme_name = ''
        self.pres_props_xml = ''
        self.results = []

    def load(self):
        # slides
        for f in self.nl:
            if f.startswith('ppt/slides/slide') and f.endswith('.xml') and '_rels' not in f:
                xml = self.zf.read(f).decode('utf-8')
                self.slides[f] = {'root': ET.fromstring(xml), 'xml': xml}
        # theme
        if 'ppt/theme/theme1.xml' in self.nl:
            self.theme_name = ET.fromstring(
                self.zf.read('ppt/theme/theme1.xml').decode('utf-8')
            ).get('name', '')
        # presProps
        if 'ppt/presProps.xml' in self.nl:
            self.pres_props_xml = self.zf.read('ppt/presProps.xml').decode('utf-8')

    def get_slide_count(self):
        return len(self.slides)

    def get_slide_layout(self, fname):
        name = fname.split('/')[-1]
        rf = f'ppt/slides/_rels/{name}.rels'
        if rf in self.nl:
            for rel in ET.fromstring(self.zf.read(rf).decode('utf-8')).iter(f'{{{PKG}}}Relationship'):
                if 'slideLayout' in rel.get('Type', ''):
                    return rel.get('Target', '').split('/')[-1].replace('.xml', '')
        return ''

    def has_image(self, fname):
        name = fname.split('/')[-1]
        rf = f'ppt/slides/_rels/{name}.rels'
        if rf in self.nl:
            for rel in ET.fromstring(self.zf.read(rf).decode('utf-8')).iter(f'{{{PKG}}}Relationship'):
                if 'image' in rel.get('Type', ''):
                    return rel.get('Target', '')
        return ''

    def R(self, req, name, score, full, detail):
        self.results.append({'req': req, 'name': name, 'score': min(score, full),
                              'full': full, 'detail': detail})

    def grade(self):
        self.load()
        # ===== 填充你的 req 函数 =====
        self.req1()
        # self.req2()
        # ...

        total = sum(r['score'] for r in self.results)
        full = sum(r['full'] for r in self.results)
        print(f'\n{"="*56}\n  总分: {total}/{full} = {total/full*100:.1f}%\n{"="*56}')

    # ===== 示例: 主题+标题幻灯片 =====
    def req1(self):
        d = []; s = 0
        if self.theme_name == '奥斯汀':
            s += 5; d.append('主题"奥斯汀"')
        else:
            d.append(f'NG 主题={self.theme_name}')

        if 'ppt/slides/slide1.xml' not in self.slides:
            self.R(1, '主题+标题', s, 15, d); return

        if self.get_slide_layout('ppt/slides/slide1.xml') == 'slideLayout1':
            s += 3; d.append('标题幻灯片')

        s1 = self.slides['ppt/slides/slide1.xml']['xml']
        if '你的标题' in s1:
            s += 3; d.append('标题文本')
        # 字号检测 (DrawingML 字号 ×100)
        for sz in re.findall(r'sz="(\d+)"', s1):
            if sz == '2800':  # 28pt × 100
                s += 2; d.append(f'字号28pt'); break

        self.R(1, '主题+标题', s, 15, d)


if __name__ == '__main__':
    docx = sys.argv[1] if len(sys.argv) > 1 else None
    if docx is None:
        base = os.path.dirname(os.path.abspath(__file__))
        for f in Path(base).glob('1结果.pptx'): docx = str(f); break
    if not docx:
        print('用法: python ppt_grader_template.py <文件.pptx>'); sys.exit(1)
    print(f'评分: {os.path.basename(docx)}')
    g = PPTGrader(docx)
    g.grade()
