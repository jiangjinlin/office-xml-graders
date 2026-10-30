#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 2010 操作题 XML 评分器 — 通用模板（非修订版）
直接解析答案文档XML结构进行评分，不依赖修订标记

用法: cp this.py word_grader_N.py → 填充 req_N() → 测试
"""
import zipfile, os, sys, io, re, shutil
from pathlib import Path
import xml.etree.ElementTree as ET

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
V = 'urn:schemas-microsoft-com:vml'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

def T(tag):
    if ':' not in tag: return tag
    p, l = tag.split(':', 1)
    ns = {'w': W, 'wp': WP, 'a': A, 'v': V, 'r': R, 'pkg': PKG}
    return '{' + ns[p] + '}' + l if p in ns else tag

def GA(elem, attr, default=None):
    if elem is None: return default
    v = elem.get(T(attr))
    if v is not None: return v
    return elem.get(attr.split(':')[1] if ':' in attr else attr, default)

def GT(elem):
    if elem is None: return ''
    return ''.join(t.text or '' for t in elem.iter(T('w:t')))

class WordGraderTemplate:
    """Word XML 评分器 — 直接解析"""
    def __init__(self, path):
        self.docx = Path(path)
        self.zf = zipfile.ZipFile(path)
        self.nl = self.zf.namelist()
        self.doc_root = None
        self.all_xml = ''
        self.results = []

    def unzip_to_temp(self):
        d = self.docx.parent / '_xml_work'
        if d.exists(): shutil.rmtree(d)
        with zipfile.ZipFile(self.docx, 'r') as z: z.extractall(d)
        return d

    def load_all(self):
        d = self.unzip_to_temp()
        parts = []
        for f in d.rglob('*.xml'):
            try:
                parts.append(f.read_text(encoding='utf-8'))
            except: pass
        self.all_xml = '\n'.join(parts)
        self.doc_root = ET.parse(str(d / 'word/document.xml')).getroot()

    def R(self, req, name, score, full, detail):
        s = min(score, full)
        self.results.append({'req': req, 'name': name, 'score': s, 'full': full, 'detail': detail})
        print(f'\n[{("PASS" if s>=full*.6 else "FAIL")}] {req:2d}. {name:<14s} {float(s):5.1f} / {full}')
        for d in detail:
            print(f'      {d[:100]}')

    def grade(self):
        self.load_all()
        # ===== 填充你的 req 函数 =====
        self.req1()
        # self.req2()
        # ...

        total = sum(r['score'] for r in self.results)
        full = sum(r['full'] for r in self.results)
        print(f'\n{"="*56}\n  总分: {total}/{full} = {total/full*100:.1f}%\n{"="*56}')

    # ===== 示例: 页面设置 =====
    def req1(self):
        d = []; s = 0
        sp = next(self.doc_root.iter(T('w:sectPr')), None)
        if sp is None: self.R(1, '页面设置', 0, 10, ['未找到sectPr']); return
        pgSz = sp.find(T('w:pgSz'))
        if pgSz is not None:
            code = GA(pgSz, 'w:code', '')
            if code == '9': s += 3; d.append(f'A4纸张')
            else: d.append(f'NG 非A4(code={code})')
        pgMar = sp.find(T('w:pgMar'))
        if pgMar is not None:
            top = int(GA(pgMar, 'w:top', '0'))
            if abs(top - 1134) < 50: s += 2; d.append(f'上边距{top/567:.1f}cm')
            else: d.append(f'NG 上边距={top}')
        self.R(1, '页面设置', s, 10, d)


if __name__ == '__main__':
    docx = sys.argv[1] if len(sys.argv) > 1 else None
    if docx is None:
        base = os.path.dirname(os.path.abspath(__file__))
        for f in Path(base).glob('1结果.docx'): docx = str(f); break
    if not docx:
        print('用法: python word_grader_template.py <文件.docx>'); sys.exit(1)
    print(f'评分: {os.path.basename(docx)}')
    g = WordGraderTemplate(docx)
    g.grade()
