#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 2010 操作题 XML底层精准评分器 - 第3套题 v2
基于系统性XML分析，针对文件夹3的11项要求
"""

import sys, os, io, zipfile, datetime, shutil
from pathlib import Path
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import extract_word_parts

pass  # stdout redirect removed for web

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WPNS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
ANS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PICNS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
VNS = 'urn:schemas-microsoft-com:vml'
W14NS = 'http://schemas.microsoft.com/office/word/2010/wordml'

def T(tag):
    if ':' not in tag: return tag
    p, l = tag.split(':', 1)
    ns_map = {'w': WNS, 'wp': WPNS, 'a': ANS, 'pic': PICNS, 'r': RNS, 'v': VNS, 'w14': W14NS}
    if p in ns_map:
        return '{' + ns_map[p] + '}' + l
    return tag

def X(parent, path):
    if parent is None: return None
    if '//' in path:
        pre, post = path.split('//', 1)
        if pre in ('.', ''):
            ctx = parent
        else:
            ctx = X(parent, pre)
        if ctx is None: return None
        return ctx.find('.//' + T(post))
    parts = path.split('/')
    node = parent
    for p in parts[:-1]:
        if node is None: return None
        node = node.find(T(p))
    if node is None: return None
    return node.find(T(parts[-1]))

def XA(parent, path):
    if parent is None: return []
    if '//' in path:
        pre, post = path.split('//', 1)
        if pre in ('.', ''):
            ctx = parent
        else:
            ctx = X(parent, pre)
        if ctx is None: return []
        return ctx.findall('.//' + T(post))
    parts = path.split('/')
    node = parent
    for p in parts[:-1]:
        if node is None: return []
        node = node.find(T(p))
    if node is None: return []
    return node.findall(T(parts[-1]))

def GA(elem, attr, default=None):
    if elem is None: return default
    val = elem.get(T(attr))
    if val is not None: return val
    return elem.get(attr.split(':')[1] if ':' in attr else attr, default)

def GT(elem):
    if elem is None: return ""
    parts = []
    for t in elem.findall('.//{' + WNS + '}t'):
        if t.text: parts.append(t.text)
    return ''.join(parts)


class XMLGrader:
    def __init__(self, docx_path):
        self.docx = Path(docx_path)
        self.xml_dir = None
        self.results = []
        self.doc_root = None
        self.settings_root = None
        self.styles_root = None
        self.headers = {}
        self.footers = {}
        self.rels = {}
        self.main_rels = {}
        self.media = []

    def unzip(self):
        self.xml_dir = self.docx.parent / ('_xml_' + self.docx.stem)
        if self.xml_dir.exists():
            shutil.rmtree(self.xml_dir)
        extract_word_parts(self.docx, self.xml_dir)
        print('  解压完成:', self.xml_dir.name)

    def load_all(self):
        # 加载主文档
        doc_xml = self.xml_dir / 'word' / 'document.xml'
        if doc_xml.exists():
            self.doc_root = ET.parse(str(doc_xml)).getroot()
            print('  加载 document.xml')

        # 加载样式
        styles_xml = self.xml_dir / 'word' / 'styles.xml'
        if styles_xml.exists():
            self.styles_root = ET.parse(str(styles_xml)).getroot()
            print('  加载 styles.xml')

        # 加载设置
        settings_xml = self.xml_dir / 'word' / 'settings.xml'
        if settings_xml.exists():
            self.settings_root = ET.parse(str(settings_xml)).getroot()
            print('  加载 settings.xml')

        # 加载关系文件
        rels_xml = self.xml_dir / 'word' / '_rels' / 'document.xml.rels'
        if rels_xml.exists():
            tree = ET.parse(str(rels_xml))
            for rel in tree.getroot().findall('.//{' + RNS + '}Relationship'):
                self.main_rels[rel.get('Id')] = rel.get('Target')

        # 加载页眉页脚
        word_dir = self.xml_dir / 'word'
        for f in word_dir.glob('header*.xml'):
            tree = ET.parse(str(f))
            self.headers[f.name] = tree.getroot()
            print('  加载', f.name)

        for f in word_dir.glob('footer*.xml'):
            tree = ET.parse(str(f))
            self.footers[f.name] = tree.getroot()
            print('  加载', f.name)

        # 加载图片列表
        media_dir = self.xml_dir / 'word' / 'media'
        if media_dir.exists():
            self.media = [f.name for f in media_dir.glob('*')]

    def grade(self):
        self.results = []
        for i in range(1, 12):
            req_method = getattr(self, 'req%d' % i, None)
            if req_method:
                result = req_method()
                self.results.append(result)
                status = 'OK' if result['score'] >= result['full'] * 0.7 else ('PART' if result['score'] > 0 else 'NG')
                print('  [%d] %s: %d/%d (%s)' % (result['req'], result['name'], result['score'], result['full'], status))

    def report(self):
        total = sum(r['score'] for r in self.results)
        full = sum(r['full'] for r in self.results)
        pct = (total / full * 100) if full > 0 else 0

        html = '''<html><head><meta charset="utf-8"><title>评分报告</title>
<style>body{font-family:微软雅黑,sans-serif;margin:40px}
h1{color:#2c3e50}.item{background:#f8f9fa;padding:15px;margin:10px 0;border-radius:8px}
.ok{color:green} .part{color:orange} .ng{color:red}
pre{background:#eee;padding:10px;overflow-x:auto}
</style></head><body>
<h1>Word 2010 操作题评分报告</h1>
<p>文件: %s</p>
<p>总分: <b>%d/%d = %.1f%%</b></p>
<hr>
''' % (self.docx.name, total, full, pct)

        for r in self.results:
            cls = 'ok' if r['score'] >= r['full'] * 0.7 else ('part' if r['score'] > 0 else 'ng')
            html += '<div class="item">\n'
            html += '<h3>[%d] %s: <span class="%s">%d/%d (%s)</span></h3>\n' % (r['req'], r['name'], cls, r['score'], r['full'], r['status'])
            html += '<pre>\n'
            for line in r['detail']:
                html += line + '\n'
            html += '</pre></div>\n'

        html += '</body></html>'

        report_path = self.docx.parent / ('report_' + self.docx.stem + '.html')
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html)
        print('\n报告已生成:', report_path.name)
        return report_path

    # ------------------------------------------------------------------
    # req1: 替换文字
    # ------------------------------------------------------------------
    def req1(self):
        log = []; score = 0
        if self.doc_root is not None:
            xs = ET.tostring(self.doc_root, encoding='unicode')
            if '动完' not in xs and '动物' in xs:
                score = 5; log.append('  OK 已替换"动完"为"动物" (+5)')
            elif '动完' in xs:
                log.append('  NG 仍存在"动完"')
            else:
                log.append('  NG 未找到"动物"')

        return {'req': 1, 'name': '替换文字', 'score': score, 'full': 5, 'detail': log, 'status': '通过' if score >= 4 else '未通过'}

    # ------------------------------------------------------------------
    # req2: 页面设置
    # ------------------------------------------------------------------
    def req2(self):
        log = []; score = 0
        if self.doc_root is not None:
            for sp in self.doc_root.findall('.//{' + WNS + '}sectPr'):
                # 页边距（A4为素材默认，不计分）
                pgMar = sp.find('{' + WNS + '}pgMar')
                if pgMar is not None:
                    top = GA(pgMar, 'w:top', '0')
                    bottom = GA(pgMar, 'w:bottom', '0')
                    left = GA(pgMar, 'w:left', '0')
                    right = GA(pgMar, 'w:right', '0')
                    gutter = GA(pgMar, 'w:gutter', '0')
                    log.append('  pgMar: top=%s bottom=%s left=%s right=%s gutter=%s' % (top, bottom, left, right, gutter))
                    # 3cm=1701缇, 3.5cm=1985缇, 0.6cm=340缇
                    if abs(int(top) - 1701) < 50 and abs(int(bottom) - 1701) < 50:
                        score += 3; log.append('  OK 上下3cm (+3)')
                    if abs(int(left) - 1985) < 50 and abs(int(right) - 1985) < 50:
                        score += 3; log.append('  OK 左右3.5cm (+3)')
                    if abs(int(gutter) - 340) < 20:
                        score += 2; log.append('  OK 装订线0.6cm左侧 (+2)')

                # 文档网格
                docGrid = sp.find('{' + WNS + '}docGrid')
                if docGrid is not None:
                    linePitch = GA(docGrid, 'w:linePitch', '0')
                    charSpace = GA(docGrid, 'w:charSpace', '0')
                    log.append('  docGrid: linePitch=%s charSpace=%s' % (linePitch, charSpace))
                    # 每页42行: linePitch=319, 每行40字符: charSpace=-4285
                    if linePitch and int(linePitch) == 319:
                        score += 1; log.append('  OK 每页42行 (+1)')
                    if charSpace and abs(int(charSpace) + 4285) < 200:
                        score += 1; log.append('  OK 每行40字符 (+1)')
                break

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 2, 'name': '页面设置', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req3: 标题格式
    # ------------------------------------------------------------------
    def req3(self):
        log = []; score = 0
        if self.doc_root is not None:
            paras = self.doc_root.findall('.//{' + WNS + '}p')
            if paras:
                title = paras[0]
                txt = GT(title)
                log.append('  标题: ' + txt[:20])

                # 检查字体、加粗
                pPr = title.find('{' + WNS + '}pPr')
                if pPr is not None:
                    rPr = pPr.find('{' + WNS + '}rPr')
                    if rPr is not None:
                        rf = rPr.find('{' + WNS + '}rFonts')
                        if rf is not None:
                            ea = GA(rf, 'w:eastAsia', '')
                            if '黑' in ea:
                                score += 2; log.append('  OK 黑体 (+2)')

                        if rPr.find('{' + WNS + '}b') is not None:
                            score += 2; log.append('  OK 加粗 (+2)')

                        sz = rPr.find('{' + WNS + '}sz')
                        if sz is not None and int(GA(sz, 'w:val', '0')) >= 40:
                            score += 1; log.append('  OK 一号及以上 (+1)')

                # 检查居中
                if pPr is not None:
                    jc = pPr.find('{' + WNS + '}jc')
                    if jc is not None and GA(jc, 'w:val', '') == 'center':
                        score += 2; log.append('  OK 居中 (+2)')
                    else:
                        # 检查样式
                        pStyle = pPr.find('{' + WNS + '}pStyle')
                        if pStyle is not None and self.styles_root is not None:
                            style_id = GA(pStyle, 'w:val', '')
                            for style in self.styles_root.findall('.//{' + WNS + '}style'):
                                if style.get('{' + WNS + '}styleId', '') == style_id:
                                    s_jc = style.find('.//{' + WNS + '}jc')
                                    if s_jc is not None and GA(s_jc, 'w:val', '') == 'center':
                                        score += 2; log.append('  OK 居中(样式) (+2)')
                                        break

                # 检查文本效果（艺术字）
                if pPr is not None:
                    rPr = pPr.find('{' + WNS + '}rPr')
                    if rPr is not None:
                        # 填充效果
                        textFill = rPr.find('.//{' + W14NS + '}textFill')
                        if textFill is not None:
                            score += 2; log.append('  OK 文本填充效果 (+2)')

                        # 轮廓效果
                        textOutline = rPr.find('.//{' + W14NS + '}textOutline')
                        if textOutline is not None:
                            score += 2; log.append('  OK 文本轮廓效果 (+2)')

                        # 阴影效果
                        shadow = rPr.find('.//{' + W14NS + '}shadow')
                        if shadow is not None:
                            score += 1; log.append('  OK 文本阴影效果 (+1)')

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 3, 'name': '标题格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req4: 正文格式
    # ------------------------------------------------------------------
    def req4(self):
        log = []; score = 0
        if self.doc_root is not None:
            paras = self.doc_root.findall('.//{' + WNS + '}p')
            if len(paras) > 1:
                body = paras[1]
                txt = GT(body)
                log.append('  正文: ' + txt[:30])

                pPr = body.find('{' + WNS + '}pPr')
                if pPr is not None:
                    # 字体（小四号来自素材既有的样式 a3，两边一致，不计分）
                    rPr = pPr.find('{' + WNS + '}rPr')
                    if rPr is not None:
                        rf = rPr.find('{' + WNS + '}rFonts')
                        if rf is not None:
                            ea = GA(rf, 'w:eastAsia', '')
                            if '华文行楷' in ea or '行楷' in ea:
                                score += 4; log.append('  OK 华文行楷 (+4)')

                    # 缩进
                    ind = pPr.find('{' + WNS + '}ind')
                    if ind is not None:
                        flc = GA(ind, 'w:firstLineChars', '0')
                        lc = GA(ind, 'w:leftChars', '0')
                        rc = GA(ind, 'w:rightChars', '0')
                        log.append('  缩进: firstLineChars=%s leftChars=%s rightChars=%s' % (flc, lc, rc))
                        # 首行缩进2字符=200/20, 左右缩进0.5字符=50/20
                        if abs(int(flc) - 200) < 20:
                            score += 2; log.append('  OK 首行缩进2字符 (+2)')
                        if abs(int(lc) - 50) < 10 and abs(int(rc) - 50) < 10:
                            score += 2; log.append('  OK 左右缩进0.5字符 (+2)')

                    # 间距
                    spacing = pPr.find('{' + WNS + '}spacing')
                    if spacing is not None:
                        bl = GA(spacing, 'w:beforeLines', '0')
                        al = GA(spacing, 'w:afterLines', '0')
                        log.append('  间距: beforeLines=%s afterLines=%s' % (bl, al))
                        # 段前1行=100/100, 段后0.5行=50/100
                        if abs(int(bl) - 100) < 10:
                            score += 1; log.append('  OK 段前1行 (+1)')
                        if abs(int(al) - 50) < 10:
                            score += 1; log.append('  OK 段后0.5行 (+1)')

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 4, 'name': '正文格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req5: 图片设置
    # ------------------------------------------------------------------
    def req5(self):
        log = []; score = 0
        if self.doc_root is not None:
            # 检查drawing中的图片
            found = False
            for dwg in self.doc_root.findall('.//{' + WNS + '}drawing'):
                blip = dwg.find('.//{' + ANS + '}blip')
                if blip is not None:
                    found = True
                    log.append('  找到图片drawing')
                    score += 2

                    # 尺寸
                    ext = dwg.find('.//{' + WPNS + '}extent')
                    if ext is not None:
                        cx = int(GA(ext, 'cx', '0')) / 914400 * 2.54
                        cy = int(GA(ext, 'cy', '0')) / 914400 * 2.54
                        log.append('  尺寸: %.1fcm x %.1fcm' % (cx, cy))
                        if abs(cx - 8) < 0.5 and abs(cy - 5) < 0.5:
                            score += 3; log.append('  OK 尺寸8x5cm (+3)')

                    # 环绕
                    if dwg.find('.//{' + WPNS + '}wrapSquare') is not None:
                        score += 2; log.append('  OK 四周型环绕 (+2)')

                    # 图片样式（柔化边缘椭圆）
                    pic = dwg.find('.//{' + PICNS + '}pic')
                    if pic is not None:
                        log.append('  找到pic元素')
                        score += 3; log.append('  OK 图片样式 (+3)')
                    break

            if not found:
                log.append('  NG 未找到图片')

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 5, 'name': '图片设置', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req6: 页眉页脚
    # ------------------------------------------------------------------
    def req6(self):
        log = []; score = 0
        # 页眉
        for nm, root in self.headers.items():
            txt = GT(root)
            if '保护野生动物' in txt:
                log.append('  页眉: ' + txt[:20])
                score += 3; log.append('  OK 页眉内容 (+3)')

                # 检查字体
                rPr = root.find('.//{' + WNS + '}rPr')
                if rPr is not None:
                    rf = rPr.find('{' + WNS + '}rFonts')
                    if rf is not None:
                        ea = GA(rf, 'w:eastAsia', '')
                        if '雅黑' in ea:
                            score += 2; log.append('  OK 微软雅黑 (+2)')

                    sz = rPr.find('{' + WNS + '}sz')
                    if sz is not None and int(GA(sz, 'w:val', '0')) == 21:
                        score += 2; log.append('  OK 五号 (+2)')

                    if rPr.find('{' + WNS + '}b') is not None:
                        score += 1; log.append('  OK 加粗 (+1)')
                    else:
                        # 检查样式链中的加粗
                        pStyle = root.find('.//{' + WNS + '}pStyle')
                        if pStyle is not None and self.styles_root is not None:
                            style_id = GA(pStyle, 'w:val', '')
                            for style in self.styles_root.findall('.//{' + WNS + '}style'):
                                if style.get('{' + WNS + '}styleId', '') == style_id:
                                    s_rPr = style.find('{' + WNS + '}rPr')
                                    if s_rPr is not None and s_rPr.find('{' + WNS + '}b') is not None:
                                        bval = s_rPr.find('{' + WNS + '}b').get('{' + WNS + '}val', '')
                                        if bval != '0':
                                            score += 1; log.append('  OK 加粗(样式) (+1)')
                                    # 也检查basedOn样式
                                    basedOn = style.find('{' + WNS + '}basedOn')
                                    if basedOn is not None:
                                        base_id = GA(basedOn, 'w:val', '')
                                        for bs in self.styles_root.findall('.//{' + WNS + '}style'):
                                            if bs.get('{' + WNS + '}styleId', '') == base_id:
                                                bs_rPr = bs.find('{' + WNS + '}rPr')
                                                if bs_rPr is not None and bs_rPr.find('{' + WNS + '}b') is not None:
                                                    bval = bs_rPr.find('{' + WNS + '}b').get('{' + WNS + '}val', '')
                                                    if bval != '0':
                                                        score += 1; log.append('  OK 加粗(基础样式%s) (+1)' % base_id)
                                                break
                                    break
                break

        # 页脚
        for nm, root in self.footers.items():
            txt = GT(root)
            if txt.strip():
                log.append('  页脚: ' + txt[:20])
                score += 2; log.append('  OK 页脚 (+2)')

                # 检查罗马数字
                xs = ET.tostring(root, encoding='unicode')
                if 'Ⅰ' in xs or 'II' in xs or 'upperRoman' in xs:
                    score += 2; log.append('  OK 罗马数字 (+2)')
                break

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 6, 'name': '页眉页脚', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req7: 第五段格式（正文第五段：红色字体、黄色1.5磅单实线阴影边框、浅蓝底纹）
    # 标准答案中第五段=段落6（跳过标题后的第5个正文段落）
    # ------------------------------------------------------------------
    def req7(self):
        log = []; score = 0
        if self.doc_root is not None:
            paras = self.doc_root.findall('.//{' + WNS + '}p')
            # 跳过标题(段落1)，找正文第五段 = 总段落6 (index 5)
            target_para = None
            if len(paras) >= 6:
                target_para = paras[5]
                txt = GT(target_para)
                log.append('  目标段落(第五段): %s...' % txt[:30])
            
            # 备选：如果段落6不包含"打击"，找包含"打击"且有红色/边框/底纹的段落
            if target_para is not None:
                t_txt = GT(target_para)
                if '打击' not in t_txt:
                    log.append('  段落6未含"打击"，搜索含"打击"且有格式的段落')
                    for i, p in enumerate(paras):
                        p_txt = GT(p)
                        if '打击' in p_txt:
                            pPr = p.find('{' + WNS + '}pPr')
                            if pPr is not None:
                                has_bdr = pPr.find('{' + WNS + '}pBdr') is not None
                                has_shd = pPr.find('{' + WNS + '}shd') is not None
                                rPr = pPr.find('{' + WNS + '}rPr')
                                has_red = False
                                if rPr is not None:
                                    c = rPr.find('{' + WNS + '}color')
                                    if c is not None and 'FF0000' in GA(c, 'w:val', ''):
                                        has_red = True
                                if has_bdr or has_shd or has_red:
                                    target_para = p
                                    log.append('  改用段落%d (含"打击"且有格式)' % (i+1))
                                    break
            
            if target_para is None:
                log.append('  NG 未找到目标段落')
                return {'req': 7, 'name': '第五段格式', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}
            
            pPr = target_para.find('{' + WNS + '}pPr')
            if pPr is not None:
                # 字体颜色（红色）- 检查pPr/rPr和run/rPr
                rPr = pPr.find('{' + WNS + '}rPr')
                color_found = False
                if rPr is not None:
                    color = rPr.find('{' + WNS + '}color')
                    if color is not None:
                        c = GA(color, 'w:val', '')
                        log.append('  pPr/rPr颜色: ' + c)
                        if 'FF0000' in c.upper():
                            score += 2; log.append('  OK 红色字体 (+2)')
                            color_found = True
                # 也检查run级别的颜色
                if not color_found:
                    for run in target_para.findall('.//{' + WNS + '}r'):
                        rrPr = run.find('{' + WNS + '}rPr')
                        if rrPr is not None:
                            rc = rrPr.find('{' + WNS + '}color')
                            if rc is not None and 'FF0000' in GA(rc, 'w:val', '').upper():
                                score += 2; log.append('  OK 红色字体(run) (+2)')
                                color_found = True
                                break
                
                # 段落边框
                pBdr = pPr.find('{' + WNS + '}pBdr')
                if pBdr is not None:
                    log.append('  找到pBdr')
                    
                    # 检查黄色边框
                    yellow_found = False
                    for side in ['top', 'left', 'bottom', 'right']:
                        border = pBdr.find('{' + WNS + '}' + side)
                        if border is not None:
                            c = GA(border, 'w:color', '')
                            sz = GA(border, 'w:sz', '0')
                            shadow = GA(border, 'w:shadow', '0')
                            val = GA(border, 'w:val', '')
                            log.append('    %s: val=%s sz=%s color=%s shadow=%s' % (side, val, sz, c, shadow))
                            if 'FFFF00' in c.upper():
                                yellow_found = True
                    
                    if yellow_found:
                        score += 2; log.append('  OK 黄色边框 (+2)')
                    
                    # 检查1.5磅单实线 (sz=12 = 1.5pt)
                    single_ok = True
                    for side in ['top', 'left', 'bottom', 'right']:
                        border = pBdr.find('{' + WNS + '}' + side)
                        if border is not None:
                            val = GA(border, 'w:val', '')
                            sz = GA(border, 'w:sz', '0')
                            if val != 'single' or abs(int(sz) - 12) > 2:
                                single_ok = False
                    if single_ok:
                        score += 2; log.append('  OK 1.5磅单实线 (+2)')
                    
                    # 检查阴影
                    has_shadow = False
                    for side in ['top', 'left', 'bottom', 'right']:
                        border = pBdr.find('{' + WNS + '}' + side)
                        if border is not None and GA(border, 'w:shadow', '0') == '1':
                            has_shadow = True
                            break
                    if has_shadow:
                        score += 2; log.append('  OK 阴影边框 (+2)')
                
            # 段落底纹 (浅蓝色 00B0F0)
            shd = pPr.find('{' + WNS + '}shd') if pPr is not None else None
            if shd is not None:
                fill = GA(shd, 'w:fill', '')
                log.append('  底纹fill: ' + fill)
                # 00B0F0是Word中的浅蓝色, CCE5FF也是浅蓝
                if fill and ('00B0F0' in fill.upper() or 'CCE5FF' in fill.upper() or 'B0DEFF' in fill.upper()):
                    score += 2; log.append('  OK 浅蓝色底纹 (+2)')
        
        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 7, 'name': '第五段格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req8: 表格边框
    # ------------------------------------------------------------------
    def req8(self):
        log = []; score = 0
        if self.doc_root is not None:
            tbls = self.doc_root.findall('.//{' + WNS + '}tbl')
            if tbls:
                tbl = tbls[0]
                tblPr = tbl.find('{' + WNS + '}tblPr')
                if tblPr is not None:
                    tblBorders = tblPr.find('{' + WNS + '}tblBorders')
                    if tblBorders is not None:
                        log.append('  找到tblBorders')

                        # 外边框（双实线）
                        outer_ok = True
                        for side in ['top', 'left', 'bottom', 'right']:
                            border = tblBorders.find('{' + WNS + '}' + side)
                            if border is not None:
                                val = GA(border, 'w:val', '')
                                sz = GA(border, 'w:sz', '0')
                                color = GA(border, 'w:color', '')
                                log.append('    外边框%s: val=%s sz=%s color=%s' % (side, val, sz, color))
                                if val != 'double' or abs(int(sz) - 12) > 2:
                                    outer_ok = False
                        if outer_ok:
                            score += 5; log.append('  OK 外边框1.5磅双实线 (+5)')

                        # 内边框（单实线）
                        inner_ok = True
                        for side in ['insideH', 'insideV']:
                            border = tblBorders.find('{' + WNS + '}' + side)
                            if border is not None:
                                val = GA(border, 'w:val', '')
                                sz = GA(border, 'w:sz', '0')
                                color = GA(border, 'w:color', '')
                                log.append('    内边框%s: val=%s sz=%s color=%s' % (side, val, sz, color))
                                if val != 'single' or abs(int(sz) - 24) > 2:
                                    inner_ok = False
                        if inner_ok:
                            score += 5; log.append('  OK 内边框3磅单实线 (+5)')

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 8, 'name': '表格边框', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req9: 表格内容格式（水平居中、自动调整、表格居中、楷体四号水绿色、升序排序）
    # 注意：楷体可能继承自默认样式"a"(Normal)，需检查样式链
    # ------------------------------------------------------------------
    def req9(self):
        log = []; score = 0
        if self.doc_root is not None:
            tbls = self.doc_root.findall('.//{' + WNS + '}tbl')
            if tbls:
                tbl = tbls[0]

                # 表格居中
                tblPr = tbl.find('{' + WNS + '}tblPr')
                if tblPr is not None:
                    jc = tblPr.find('{' + WNS + '}jc')
                    if jc is not None and GA(jc, 'w:val', '') == 'center':
                        score += 2; log.append('  OK 表格居中 (+2)')

                    # 自动调整 (tblW type=auto 且无 tblLayout type=fixed)
                    tblW = tblPr.find('{' + WNS + '}tblW')
                    tblLayout = tblPr.find('{' + WNS + '}tblLayout')
                    auto_fit = False
                    if tblW is not None and GA(tblW, 'w:type', '') == 'auto':
                        if tblLayout is None or GA(tblLayout, 'w:type', '') != 'fixed':
                            auto_fit = True
                    if auto_fit:
                        score += 1; log.append('  OK 根据内容自动调整 (+1)')

                # 获取默认样式的eastAsia字体（用于样式继承检查）
                default_font = ''
                if self.styles_root is not None:
                    for style in self.styles_root.findall('.//{' + WNS + '}style'):
                        sid = style.get('{' + WNS + '}styleId', '')
                        stype = style.get('{' + WNS + '}type', '')
                        sdef = style.get('{' + WNS + '}default', '')
                        if stype == 'paragraph' and sdef == '1':
                            rPr = style.find('{' + WNS + '}rPr')
                            if rPr is not None:
                                rf = rPr.find('{' + WNS + '}rFonts')
                                if rf is not None:
                                    default_font = GA(rf, 'w:eastAsia', '')
                                    log.append('  默认样式字体: %s' % default_font)

                # 单元格格式
                cells = tbl.findall('.//{' + WNS + '}tc')
                if cells:
                    # 检查第一个有文本的单元格
                    font_found = False
                    size_found = False
                    color_found = False
                    align_found = False
                    
                    for cell in cells:
                        for p in cell.findall('.//{' + WNS + '}p'):
                            # 检查段落对齐 (jc=center)
                            pPr = p.find('{' + WNS + '}pPr')
                            if pPr is not None:
                                jc = pPr.find('{' + WNS + '}jc')
                                if jc is not None and GA(jc, 'w:val', '') == 'center':
                                    if not align_found:
                                        align_found = True
                                        score += 2; log.append('  OK 水平居中 (+2)')
                            
                            # 检查run格式
                            for run in p.findall('.//{' + WNS + '}r'):
                                rPr = run.find('{' + WNS + '}rPr')
                                if rPr is not None:
                                    txt = ''.join(t.text or '' for t in run.findall('.//{' + WNS + '}t'))
                                    if not txt:
                                        continue
                                    
                                    # 字体（楷体）- 检查run级别和样式继承
                                    if not font_found:
                                        rf = rPr.find('{' + WNS + '}rFonts')
                                        run_font = ''
                                        if rf is not None:
                                            run_font = GA(rf, 'w:eastAsia', '')
                                        # 如果run没有明确字体，使用默认样式的字体
                                        if not run_font and default_font:
                                            run_font = default_font
                                        if '楷' in run_font:
                                            font_found = True
                                            score += 2; log.append('  OK 楷体 (%s) (+2)' % run_font)
                                    
                                    # 字号（四号=28）
                                    if not size_found:
                                        sz = rPr.find('{' + WNS + '}sz')
                                        if sz is not None and int(GA(sz, 'w:val', '0')) == 28:
                                            size_found = True
                                            score += 1; log.append('  OK 四号(28) (+1)')
                                    
                                    # 颜色（水绿色 4BACC6）
                                    if not color_found:
                                        color = rPr.find('{' + WNS + '}color')
                                        if color is not None:
                                            c = GA(color, 'w:val', '')
                                            log.append('  表格文字颜色: %s' % c)
                                            if '4BACC6' in c.upper() or '70AD47' in c.upper() or '4CAF50' in c.upper():
                                                color_found = True
                                                score += 2; log.append('  OK 水绿色 (+2)')
                                break  # 只检查第一个有文本的run
                        if font_found and size_found and color_found and align_found:
                            break

                    # 检查单元格垂直居中 (vAlign=center)
                    vAlign_found = False
                    for cell in cells[:3]:
                        tcPr = cell.find('{' + WNS + '}tcPr')
                        if tcPr is not None:
                            vAlign = tcPr.find('{' + WNS + '}vAlign')
                            if vAlign is not None and GA(vAlign, 'w:val', '') == 'center':
                                vAlign_found = True
                                break
                    if vAlign_found:
                        score += 1; log.append('  OK 垂直居中 (+1)')

                # 检查排序（灭绝时间升序）
                rows = tbl.findall('{' + WNS + '}tr')
                years = []
                for row in rows[1:]:  # 跳过表头
                    cells_in_row = row.findall('{' + WNS + '}tc')
                    if len(cells_in_row) >= 3:
                        year_text = GT(cells_in_row[2])
                        # 提取数字
                        import re
                        nums = re.findall(r'\d+', year_text)
                        if nums:
                            years.append(int(nums[0]))
                log.append('  灭绝时间列: %s' % years)
                if years == sorted(years) and len(years) > 0:
                    score += 1; log.append('  OK 升序排序 (+1)')
                else:
                    log.append('  NG 未升序排序')

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 9, 'name': '表格内容格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req10: 标题超链接
    # 标准答案使用域代码(W:instrText HYPERLINK)实现超链接，不是w:hyperlink元素
    # ------------------------------------------------------------------
    def req10(self):
        log = []; score = 0
        if self.doc_root is not None:
            # 方法1: 检查w:hyperlink元素
            hyperlinks = self.doc_root.findall('.//{' + WNS + '}hyperlink')
            if hyperlinks:
                log.append('  找到%d个w:hyperlink元素' % len(hyperlinks))
                score += 3; log.append('  OK 找到超链接元素 (+3)')
                for hl in hyperlinks:
                    rid = hl.get('{' + RNS + '}id', '')
                    if rid in self.main_rels:
                        target = self.main_rels[rid]
                        if target == 'http://cwca.org.cn/' or target == 'http://cwca.org.cn':
                            score += 2; log.append('  OK URL=%s (+2)' % target)
                            break

            # 方法2: 检查域代码 instrText 中的 HYPERLINK
            if score == 0:
                instr_texts = self.doc_root.findall('.//{' + WNS + '}instrText')
                for it in instr_texts:
                    if it.text and 'HYPERLINK' in it.text.upper():
                        log.append('  找到域代码超链接: %s' % it.text.strip()[:80])
                        score += 3; log.append('  OK 找到超链接域代码 (+3)')
                        if 'cwca.org.cn' in it.text:
                            score += 2; log.append('  OK URL正确 (+2)')
                        break

            # 方法3: 检查原始XML中的HYPERLINK字符串
            if score == 0:
                xs = ET.tostring(self.doc_root, encoding='unicode')
                if 'HYPERLINK' in xs.upper() and 'cwca.org.cn' in xs.lower():
                    log.append('  在XML中找到HYPERLINK和cwca.org.cn')
                    score = 5; log.append('  OK 超链接 (+5)')

            if score == 0:
                log.append('  NG 未找到超链接')

        st = '通过' if score >= 5 else ('部分通过' if score >= 2 else '未通过')
        return {'req': 10, 'name': '标题超链接', 'score': min(score, 5), 'full': 5, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req11: 页面边框
    # ------------------------------------------------------------------
    def req11(self):
        log = []; score = 0
        if self.doc_root is not None:
            for sp in self.doc_root.findall('.//{' + WNS + '}sectPr'):
                pgBorders = sp.find('{' + WNS + '}pgBorders')
                if pgBorders is not None:
                    log.append('  找到pgBorders')

                    # 检查边框样式
                    all_ok = True
                    for side in ['top', 'left', 'bottom', 'right']:
                        border = pgBorders.find('{' + WNS + '}' + side)
                        if border is not None:
                            val = GA(border, 'w:val', '')
                            sz = GA(border, 'w:sz', '0')
                            color = GA(border, 'w:color', '')
                            shadow = GA(border, 'w:shadow', '0')
                            log.append('    %s: val=%s sz=%s color=%s shadow=%s' % (side, val, sz, color, shadow))
                            if val != 'single' or abs(int(sz) - 24) > 2 or shadow != '1':
                                all_ok = False

                    if all_ok:
                        score = 10; log.append('  OK 蓝色3磅单实线阴影边框 (+10)')
                break

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 11, 'name': '页面边框', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}


def main():
    if len(sys.argv) < 2:
        # 自动找当前目录的.docx
        for f in Path('.').glob('*.docx'):
            docx = str(f)
            break
    else:
        docx = sys.argv[1]

    print('=' * 60)
    print('Word 2010 操作题评分器 - 第3套题')
    print('=' * 60)
    print('文件:', docx)

    grader = XMLGrader(docx)
    grader.unzip()
    grader.load_all()
    print('\n开始评分...')
    grader.grade()
    report = grader.report()

    total = sum(r['score'] for r in grader.results)
    full = sum(r['full'] for r in grader.results)
    pct = (total / full * 100) if full > 0 else 0

    print('\n' + '=' * 60)
    print('  总得分: %d/%d = %.1f%%' % (total, full, pct))
    print('=' * 60)

if __name__ == '__main__':
    main()
