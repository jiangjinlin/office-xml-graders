#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 2010 操作题 XML底层精准评分器 - 第2套题 v3
修复: req3竖排检测, req4边框填充, req6段落边框, req6正确段落, req10表格字体
"""

import sys, os, io, zipfile, datetime, shutil
from pathlib import Path
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import has_field, extract_word_parts

pass  # stdout redirect removed for web

WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WPNS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
ANS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PICNS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
VNS = 'urn:schemas-microsoft-com:vml'

def T(tag):
    if ':' not in tag: return tag
    p, l = tag.split(':', 1)
    ns_map = {'w': WNS, 'wp': WPNS, 'a': ANS, 'pic': PICNS, 'r': RNS, 'v': VNS}
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
        self.xmldir = None
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
        self.xmldir = self.docx.parent / ('_xml_' + self.docx.stem)
        if self.xmldir.exists():
            shutil.rmtree(self.xmldir)
        extract_word_parts(self.docx, self.xmldir)
        print('  解压完成:', self.xmldir.name)

    def load_all(self):
        d = self.xmldir / 'word'
        p = d / 'document.xml'
        if p.exists(): self.doc_root = ET.parse(str(p)).getroot()
        p = d / 'settings.xml'
        if p.exists(): self.settings_root = ET.parse(str(p)).getroot()
        p = d / 'styles.xml'
        if p.exists(): self.styles_root = ET.parse(str(p)).getroot()
        for p in d.glob('header*.xml'):
            self.headers[p.name] = ET.parse(str(p)).getroot()
        for p in d.glob('footer*.xml'):
            self.footers[p.name] = ET.parse(str(p)).getroot()
        rels_dir = d / '_rels'
        if rels_dir.exists():
            for p in rels_dir.glob('*.xml.rels'):
                key = p.stem.replace('.xml', '')
                rd = {}
                for rel in ET.parse(str(p)).getroot().findall(
                        '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rd[rel.get('Id', '')] = {'target': rel.get('Target', ''), 'type': rel.get('Type', '')}
                self.rels[key] = rd
        self.main_rels = self.rels.get('document', {})
        mdir = d / 'media'
        if mdir.exists():
            self.media = [f.name for f in mdir.iterdir() if f.is_file()]
        print('  已加载: document.xml, settings.xml, styles.xml')
        print('  %d 个页眉, %d 个页脚, %d 个媒体文件' % (len(self.headers), len(self.footers), len(self.media)))

    # ------------------------------------------------------------------
    # req1: 页面设置
    # ------------------------------------------------------------------
    def req1(self):
        log = []; score = 0
        for sp in XA(self.doc_root, 'w:body//w:sectPr'):
            pgMar = X(sp, 'w:pgMar')
            if pgMar is not None:
                left = int(GA(pgMar, 'w:left', '0'))
                right = int(GA(pgMar, 'w:right', '0'))
                log.append('  页边距: 左%.1fcm 右%.1fcm' % (left/567, right/567))
                if abs(left - 1701) < 30 and abs(right - 1701) < 30:
                    score += 4; log.append('  OK 左右3cm (+4)')
            docGrid = X(sp, 'w:docGrid')
            if docGrid is not None:
                lp = GA(docGrid, 'w:linePitch', '')
                cs = GA(docGrid, 'w:charSpace', '')
                log.append('  文档网格: 每页约%d行, 每行约%d字符' % (int(lp) if lp else 0, abs(int(cs)) if cs else 0))
                if lp and int(lp) == 333: score += 3; log.append('  OK 每页约42行 (+3)')
                if cs and int(cs) < -400: score += 3; log.append('  OK 每行约41字符 (+3)')
            break
        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 1, 'name': '页面设置', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req2: 艺术字标题
    # ------------------------------------------------------------------
    def req2(self):
        log = []; score = 0
        drawings = XA(self.doc_root, 'w:body//w:drawing')
        for i, dwg in enumerate(drawings):
            txt = GT(dwg)
            if '无尽的好奇' in txt:
                log.append('  OK 找到艺术字: ' + txt[:30])
                score += 3
                for sz_elem in dwg.findall('.//{' + WNS + '}sz'):
                    v = int(GA(sz_elem, 'w:val', '0'))
                    if v == 44: score += 2; log.append('  OK 二号字 (+2)')
                if dwg.find('.//{' + WPNS + '}wrapTopAndBottom') is not None or \
                   dwg.find('.//{' + WPNS + '}wrapNone') is not None:
                    score += 2; log.append('  OK 上下型环绕 (+2)')
                pH = dwg.find('.//{' + WPNS + '}positionH')
                if pH is not None:
                    al = pH.find('{' + WPNS + '}align')
                    if al is not None and al.text == 'center':
                        score += 2; log.append('  OK 相对于栏居中 (+2)')
                break
        if score == 0: log.append('  NG 未找到艺术字')
        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 2, 'name': '艺术字标题', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req3: 竖排文本框 - 检查txbxContent中的字体和对齐
    # ------------------------------------------------------------------
    def req3(self):
        log = []; score = 0
        # 找含"中国空间站"的txbxContent（不需要"侧记"）
        txbxs = XA(self.doc_root, 'w:body//w:txbxContent')
        target_txbx = None
        for txbx in txbxs:
            txt = GT(txbx)
            if '中国空间站' in txt:
                target_txbx = txbx
                log.append('  OK 找到文本框内容: ' + txt[:20])
                score += 3
                break

        if target_txbx is None:
            log.append('  NG 未找到竖排文本框内容')
            return {'req': 3, 'name': '竖排文本框', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}

        # 检查文字方向 - 在父级drawing中检查
        # 同时也检查w:pict（VML竖排）
        found_vertical = False
        picts = XA(self.doc_root, 'w:body//w:pict')
        for pict in picts:
            xs = ET.tostring(pict, encoding='unicode')
            if '中国空间站' in xs:
                # VML文本框，假设是竖排
                found_vertical = True
                score += 4; log.append('  OK 竖排文本框(VML) (+4)')
                break

        if not found_vertical:
            # 检查drawing的a:bodyPr
            for dwg in XA(self.doc_root, 'w:body//w:drawing'):
                t = GT(dwg)
                if '中国空间站' in t:
                    bodyPr = dwg.find('.//{' + ANS + '}bodyPr')
                    if bodyPr is not None:
                        vert = bodyPr.get('vert', '')
                        if vert:
                            found_vertical = True
                            score += 4; log.append('  OK 竖排文字(%s) (+4)' % vert)
                    break
            if not found_vertical:
                # 没找到vert属性，但文本框存在，给部分分
                score += 2; log.append('  OK 文本框存在(竖排需人工确认) (+2)')

        # 检查字体格式 - 从txbxContent的w:rPr
        rPr = target_txbx.find('.//{' + WNS + '}rPr')
        if rPr is not None:
            rf = X(rPr, 'w:rFonts')
            if rf is not None:
                ea = GA(rf, 'w:eastAsia', '')
                if '黑' in ea:
                    score += 1; log.append('  OK 黑体 (+1)')
            if X(rPr, 'w:b') is not None:
                score += 1; log.append('  OK 加粗 (+1)')
            col = X(rPr, 'w:color')
            if col is not None and GA(col, 'w:val', '').upper() in ['002060', '00']:
                score += 1; log.append('  OK 深蓝色 (+1)')
            sz = X(rPr, 'w:sz')
            if sz is not None and int(GA(sz, 'w:val', '0')) == 44:
                score += 1; log.append('  OK 二号 (+1)')
            # 检查居中
            jc = target_txbx.find('.//{' + WNS + '}jc')
            if jc is not None and GA(jc, 'w:val', '') == 'center':
                score += 1; log.append('  OK 居中对齐 (+1)')

        if score >= 9: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 3, 'name': '竖排文本框', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req4: 文本框格式 - 尺寸、边框、填充、环绕
    # 边框和填充在VML w:pict中
    # ------------------------------------------------------------------
    def req4(self):
        log = []; score = 0
        # 检查w:drawing中的尺寸和环绕
        target_dwg = None
        for dwg in XA(self.doc_root, 'w:body//w:drawing'):
            txt = GT(dwg)
            if '中国空间站' in txt:
                target_dwg = dwg
                break

        if target_dwg is None:
            log.append('  NG 未找到竖排文本框drawing')
            return {'req': 4, 'name': '文本框格式', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}

        log.append('  找到竖排文本框drawing')
        # 尺寸
        ext = target_dwg.find('.//{' + WPNS + '}extent')
        if ext is not None:
            cx = int(GA(ext, 'cx', '0')) / 914400 * 2.54
            cy = int(GA(ext, 'cy', '0')) / 914400 * 2.54
            log.append('  文本框尺寸: 宽%.1fcm × 高%.1fcm' % (cx, cy))
            if abs(cx - 1.5) < 0.5 and abs(cy - 11) < 1:
                score += 2; log.append('  OK 尺寸1.5x11cm (+2)')

        # 环绕方式
        if target_dwg.find('.//{' + WPNS + '}wrapSquare') is not None:
            score += 2; log.append('  OK 四周型环绕 (+2)')

        # 检查VML w:pict中的边框和填充
        picts = XA(self.doc_root, 'w:body//w:pict')
        found_format = False
        for pict in picts:
            xs = ET.tostring(pict, encoding='unicode')
            if '中国空间站' in xs or '文本框 2' in xs or 'PowerPlusWaterMarkObject' in xs:
                # 这是竖排文本框的VML
                log.append('  找到VML pict（竖排文本框）')
                # 检查VML shape元素的strokecolor和strokeweight属性
                for shape in pict.findall('.//{' + VNS + '}shape'):
                    sc = shape.get('strokecolor', '')
                    sw = shape.get('strokeweight', '')
                    log.append('  VML形状: 边框颜色=%s, 线条宽度=%s' % (sc, sw))
                    if sc and ('red' in sc.lower() or 'ff0000' in sc.lower() or 'c00000' in sc.lower()):
                        score += 2; log.append('  OK 红色边框 (+2)')
                    if sw and '2' in sw:
                        score += 1; log.append('  OK 2磅边框 (+1)')
                    break
                # 检查VML fill（渐变填充）
                for fill in pict.findall('.//{' + VNS + '}fill'):
                    typ = fill.get('type', '')
                    log.append('  VML填充: 类型=%s' % typ)
                    if 'gradient' in typ.lower():
                        score += 2; log.append('  OK 渐变填充(+2)')
                    break
                found_format = True
                break
            elif 'Watermark' in xs.lower() or '科技日报' in xs:
                # 水印，跳过
                continue

        if not found_format:
            # 兜底检查drawing中的a:gradFill和a:ln
            if target_dwg.find('.//{' + ANS + '}gradFill') is not None:
                score += 2; log.append('  OK 渐变填充(drawing) (+2)')
            ln = target_dwg.find('.//{' + ANS + '}ln')
            if ln is not None:
                score += 2; log.append('  OK 边框(drawing) (+2)')

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 4, 'name': '文本框格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req5: 文本框对齐
    # ------------------------------------------------------------------
    def req5(self):
        log = []; score = 0
        for dwg in XA(self.doc_root, 'w:body//w:drawing'):
            txt = GT(dwg)
            if '中国空间站' in txt:
                pH = dwg.find('.//{' + WPNS + '}positionH')
                if pH is not None:
                    rf = pH.get('relativeFrom', '')
                    al = pH.find('{' + WPNS + '}align')
                    log.append('  水平对齐: 参考=%s, 对齐方式=%s' % (rf, al.text if al is not None else '未知'))
                    if al is not None and al.text == 'right':
                        score += 3; log.append('  OK 水平右对齐 (+3)')
                pV = dwg.find('.//{' + WPNS + '}positionV')
                if pV is not None:
                    rf = pV.get('relativeFrom', '')
                    al = pV.find('{' + WPNS + '}align')
                    log.append('  垂直对齐: 参考=%s, 对齐方式=%s' % (rf, al.text if al is not None else '未知'))
                    if al is not None and al.text == 'top':
                        score += 2; log.append('  OK 垂直顶端对齐 (+2)')
                break
        if score == 0: log.append('  NG 未检测到文本框对齐')
        st = '通过' if score >= 4 else ('部分通过' if score >= 2 else '未通过')
        return {'req': 5, 'name': '文本框对齐', 'score': min(score, 5), 'full': 5, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req6: 文字方框+底纹 - 检测w:pPr中的w:pBdr和w:shd
    # ------------------------------------------------------------------
    def req6(self):
        log = []; score = 0
        paras = XA(self.doc_root, 'w:body/w:p')
        # 找含"水膜实验中"的段落
        target = None
        for i, p in enumerate(paras):
            txt = GT(p)
            if '水膜实验中' in txt:
                target = p
                log.append('  找到目标段落[%d]: %s' % (i+1, txt[:30]))
                break
        if target is None:
            # 兼容只含"水膜"的情况
            for i, p in enumerate(paras):
                txt = GT(p)
                if '水膜' in txt:
                    target = p
                    log.append('  找到"水膜"段落[%d]: %s' % (i+1, txt[:30]))
                    break

        if target is None:
            log.append('  NG 未找到目标段落')
            return {'req': 6, 'name': '文字方框底纹', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}

        # 检查段落边框 w:pBdr (在w:pPr中)
        pPr = X(target, 'w:pPr')
        if pPr is not None:
            pBdr = X(pPr, 'w:pBdr')
            if pBdr is not None:
                log.append('  OK 找到w:pBdr(段落边框)')
                score += 3
                # 检查边框细节
                for side in ['top', 'left', 'bottom', 'right']:
                    b = X(pBdr, 'w:' + side)
                    if b is not None:
                        sz = GA(b, 'w:sz', '0')
                        col = GA(b, 'w:color', '')
                        sh = GA(b, 'w:shadow', '0')
                        log.append('    %s边框: 线宽=%s, 颜色=%s, 阴影=%s' % (side, sz, col, '有' if sh=='1' else '无'))
                        if sh == '1': score += 1; log.append('  OK 带阴影 (%s) (+1)' % side)
                if score >= 4: score = 7  # 边框+阴影共7分
            # 检查底纹 w:shd
            shd = X(pPr, 'w:shd')
            if shd is not None:
                fill = GA(shd, 'w:fill', '')
                log.append('  段落底纹颜色: ' + fill)
                if fill.upper() == 'FFC000':
                    score += 3; log.append('  OK 橙色底纹 (+3)')

        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 6, 'name': '文字方框底纹', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req7: 图片设置
    # ------------------------------------------------------------------
    def req7(self):
        log = []; score = 0
        drawings = XA(self.doc_root, 'w:body//w:drawing')
        for i, dwg in enumerate(drawings):
            cNv = dwg.find('.//{' + PICNS + '}cNvPr')
            if cNv is not None:
                nm = GA(cNv, 'name', '')
                if 'tp' in nm.lower():
                    log.append('  OK 找到图片: ' + nm)
                    score += 2
                    ext = dwg.find('.//{' + WPNS + '}extent')
                    if ext is not None:
                        cx = int(GA(ext, 'cx', '0')) / 914400 * 2.54
                        cy = int(GA(ext, 'cy', '0')) / 914400 * 2.54
                        log.append('    图片尺寸: 宽%.1fcm × 高%.1fcm' % (cx, cy))
                        if abs(cx - 6) < 1 and abs(cy - 3) < 1:
                            score += 2; log.append('  OK 6x3cm (+2)')
                    if dwg.find('.//{' + WPNS + '}wrapSquare') is not None:
                        score += 2; log.append('  OK 四周型 (+2)')
                    pH = dwg.find('.//{' + WPNS + '}positionH')
                    if pH is not None:
                        al = pH.find('{' + WPNS + '}align')
                        if al is not None and al.text == 'center':
                            score += 1; log.append('  OK 居中 (+1)')
                    if dwg.find('.//{' + ANS + '}softEdge') is not None:
                        score += 2; log.append('  OK 柔化边缘 (+2)')
                    if dwg.find('.//{' + ANS + '}glow') is not None:
                        score += 1; log.append('  OK 发光效果 (+1)')
                    break
        if score == 0: log.append('  NG 未找到tp.jpg')
        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 7, 'name': '图片设置', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req8: 文字水印 + 分栏
    # ------------------------------------------------------------------
    def req8(self):
        log = []; score = 0
        found_watermark = False
        for nm, root in self.headers.items():
            for shape in root.findall('.//{' + VNS + '}shape'):
                sid = shape.get('id', '')
                if 'watermark' in sid.lower():
                    found_watermark = True
                    log.append('  OK 找到水印VML: ' + sid)
                    score += 3
                    xs = ET.tostring(shape, encoding='unicode')
                    if '科技日报' in xs: score += 2; log.append('  OK 水印文字=科技日报 (+2)')
                    break
            if not found_watermark:
                xs = ET.tostring(root, encoding='unicode')
                if '科技日报' in xs and 'watermark' in xs.lower():
                    found_watermark = True
                    log.append('  OK ' + nm + ' 含水印')
                    score += 3
            if found_watermark: break

        if not found_watermark:
            log.append('  NG 未检测到文字水印')

        # 分栏
        found_cols = False
        for sp in XA(self.doc_root, 'w:body//w:sectPr'):
            cols = X(sp, 'w:cols')
            if cols is not None:
                num = GA(cols, 'w:num', '1')
                sep = GA(cols, 'w:sep', '0')
                ew = GA(cols, 'w:equalWidth', '1')
                log.append('  分栏设置: %s栏, %s宽, %s分隔线' % (num, '等' if ew=='1' else '不等', '有' if sep=='1' else '无'))
                if num == '3':
                    score += 3; log.append('  OK 三栏 (+3)')
                    found_cols = True
                    if sep == '1': score += 1; log.append('  OK 分隔线 (+1)')
                    if ew == '1': score += 1; log.append('  OK 等宽 (+1)')
                    break
        if not found_cols:
            log.append('  NG 未找到三栏设置')
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 8, 'name': '文字水印+分栏', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req9: 页眉页脚
    # ------------------------------------------------------------------
    def req9(self):
        log = []; score = 0
        # 1. 检查奇偶页不同
        if self.settings_root is not None:
            if X(self.settings_root, 'w:evenAndOddHeaders') is not None:
                score += 3; log.append('  OK 启用奇偶页不同 (+3)')
        
        # 2. 检查页眉内容和居中
        header_center = False
        for nm, root in self.headers.items():
            txt = GT(root)
            log.append('  %s页眉内容: %s' % (nm, txt[:40]))
            if '科技日报' in txt: score += 3; log.append('  OK 页眉含"科技日报" (+3)')
            if '太空授课' in txt: score += 2; log.append('  OK 页眉含"太空授课" (+2)')
            
            # 检查页眉居中 - 方法1: 直接检查w:jc
            for p in root.findall('.//{' + WNS + '}p'):
                pPr = p.find('{' + WNS + '}pPr')
                if pPr is not None:
                    jc = pPr.find('{' + WNS + '}jc')
                    if jc is not None and GA(jc, 'w:val', '') == 'center':
                        header_center = True
                        break
                    # 方法2: 检查段落样式
                    pStyle = pPr.find('{' + WNS + '}pStyle')
                    if pStyle is not None:
                        style_id = GA(pStyle, 'w:val', '')
                        # 读取styles.xml检查样式是否定义居中
                        if self.styles_root is not None:
                            for style in self.styles_root.findall('.//{' + WNS + '}style'):
                                if style.get('{' + WNS + '}styleId', '') == style_id:
                                    style_jc = style.find('.//{' + WNS + '}jc')
                                    if style_jc is not None and GA(style_jc, 'w:val', '') == 'center':
                                        header_center = True
                                        break
                    if header_center:
                        break
        
        if header_center:
            score += 4; log.append('  OK 页眉居中 (+4)')
        else:
            log.append('  NG 页眉未居中')
        
        # 3. 检查页脚
        for nm, root in self.footers.items():
            xs = ET.tostring(root, encoding='unicode')
            txt = GT(root)
            log.append('  ' + nm + ': ' + txt[:30])
            if has_field(xs, 'PAGE'):
                score += 3; log.append('  OK 页脚含PAGE域 (+3)')
                # 检查页脚右对齐（题目要求）
                for p in root.findall('.//{' + WNS + '}p'):
                    jc = p.find('.//{' + WNS + '}jc')
                    if jc is not None and GA(jc, 'w:val', '') == 'right':
                        score += 2; log.append('  OK 页脚右对齐 (+2)')
                        break
                break
        
        st = '通过' if score >= 11 else ('部分通过' if score >= 5 else '未通过')
        return {'req': 9, 'name': '页眉页脚', 'score': min(score, 15), 'full': 15, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req10: 表格格式 - 黑体小四，中部两端对齐
    # ------------------------------------------------------------------
    def req10(self):
        log = []; score = 0
        tbls = XA(self.doc_root, 'w:body//w:tbl')
        if not tbls:
            log.append('  NG 未找到表格')
            return {'req': 10, 'name': '表格格式', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}
        log.append('  找到 %d 个表格' % len(tbls))
        tbl = tbls[0]
        tblPr = X(tbl, 'w:tblPr')
        if tblPr is not None:
            jc = X(tblPr, 'w:jc')
            if jc is not None and GA(jc, 'w:val', '') == 'center':
                score += 3; log.append('  OK 表格居中 (+3)')
            tblW = X(tblPr, 'w:tblW')
            if tblW is not None and GA(tblW, 'w:type', '') == 'auto':
                score += 2; log.append('  OK 根据内容自动调整 (+2)')
        # 检查字体 - 搜索整个表格XML
        tbl_xs = ET.tostring(tbl, encoding='unicode')
        if '黑体' in tbl_xs:
            score += 2; log.append('  OK 黑体 (+2)')
        if '24' in tbl_xs and '黑体' in tbl_xs:
            score += 2; log.append('  OK 小四号 (+2)')
        # 垂直对齐
        for va in tbl.findall('.//{' + WNS + '}vAlign'):
            if GA(va, 'w:val', '') == 'center':
                score += 1; log.append('  OK 中部对齐 (+1)')
                break
        # 水平对齐（两端对齐）
        for jc in tbl.findall('.//{' + WNS + '}jc'):
            if GA(jc, 'w:val', '') == 'both':
                score += 1; log.append('  OK 两端对齐 (+1)')
                break

        if score >= 8: score = 10
        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 10, 'name': '表格格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req11: 表格边框底纹
    # ------------------------------------------------------------------
    def req11(self):
        log = []; score = 0
        tbls = XA(self.doc_root, 'w:body//w:tbl')
        if not tbls:
            return {'req': 11, 'name': '表格边框底纹', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}
        tbl = tbls[0]
        tblPr = X(tbl, 'w:tblPr')
        if tblPr is not None:
            tblBorders = X(tblPr, 'w:tblBorders')
            if tblBorders is not None:
                log.append('  OK 找到tblBorders')
                score += 2
                # 外边框
                for side in ['top', 'left', 'bottom', 'right']:
                    b = X(tblBorders, 'w:' + side)
                    if b is not None:
                        val = GA(b, 'w:val', '')
                        col = GA(b, 'w:color', '')
                        sz = GA(b, 'w:sz', '0')
                        log.append('    %s边框: %s线, 颜色=%s, 线宽=%s' % (side, val, col, sz))
                # 内边框
                if X(tblBorders, 'w:insideH') is not None or X(tblBorders, 'w:insideV') is not None:
                    score += 2; log.append('  OK 内边框 (+2)')
                # 双实线
                top = X(tblBorders, 'w:top')
                if top is not None and GA(top, 'w:val', '') == 'double':
                    score += 3; log.append('  OK 双实线外边框 (+3)')
            # 第一行底纹
            rows = XA(tbl, 'w:tr')
            if rows:
                cells = XA(rows[0], 'w:tc')
                if cells:
                    tcPr = X(cells[0], 'w:tcPr')
                    if tcPr is not None:
                        shd = X(tcPr, 'w:shd')
                        if shd is not None:
                            fill = GA(shd, 'w:fill', '')
                            log.append('  表格第一行底纹颜色: ' + fill)
                            if fill and fill.upper() != 'AUTO':
                                score += 3; log.append('  OK 第一行底纹 (+3)')
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 11, 'name': '表格边框底纹', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # 评分主流程
    # ------------------------------------------------------------------
    def grade(self):
        print('\n' + '=' * 60)
        print('  Word 2010 XML底层精准评分器 - 第2套题 v3')
        print('=' * 60)
        self.unzip()
        print('\n加载XML...')
        self.load_all()
        print('\n' + '-' * 60)
        print('  逐项检测...')
        print('-' * 60)

        checks = [self.req1, self.req2, self.req3, self.req4, self.req5,
                  self.req6, self.req7, self.req8, self.req9, self.req10, self.req11]
        names = ['页面设置', '艺术字标题', '竖排文本框', '文本框格式', '文本框对齐',
                 '文字方框底纹', '图片设置', '文字水印+分栏', '页眉页脚', '表格格式', '表格边框底纹']

        for i, chk in enumerate(checks):
            r = chk()
            self.results.append(r)
            icon = {'通过': '[OK]', '部分通过': '[--]', '未通过': '[NG]', '检测异常': '[!]'}.get(r['status'], '?')
            print('\n[%02d] %s' % (i + 1, names[i]))
            print('  %s %d/%d (%s)' % (icon, r['score'], r['full'], r['status']))
            for d in r['detail']:
                print('    ' + d)

        total = sum(r['score'] for r in self.results)
        full = sum(r['full'] for r in self.results)
        rate = total / full * 100 if full > 0 else 0
        print('\n' + '=' * 60)
        print('  总得分: %d/%d = %.1f%%' % (total, full, rate))
        print('=' * 60)
        return self.results

    def report(self, path=None):
        if not self.results: return
        if path is None: path = self.docx.parent / 'report_2.html'
        total = sum(r['score'] for r in self.results)
        full = sum(r['full'] for r in self.results)
        rate = total / full * 100 if full > 0 else 0
        if rate >= 90: gr, gc = 'A 优秀', '#28a745'
        elif rate >= 75: gr, gc = 'B 良好', '#17a2b8'
        elif rate >= 60: gr, gc = 'C 及格', '#ffc107'
        else: gr, gc = 'D 不及格', '#dc3545'

        h = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
        h += '<title>Word评分报告 - 第2套题</title>'
        h += '<style>*{margin:0;padding:0;box-sizing:border-box;}'
        h += 'body{font-family:Microsoft YaHei,sans-serif;background:linear-gradient(135deg,#667eea,#764ba2);min-height:100vh;padding:20px;}'
        h += '.box{max-width:1000px;margin:0 auto;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden;}'
        h += '.hd{background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:30px;text-align:center;}'
        h += '.hd h1{font-size:26px;margin-bottom:10px;}'
        h += '.sc{display:flex;align-items:center;justify-content:center;padding:30px;gap:50px;background:#f8f9fa;}'
        h += '.num{font-size:56px;font-weight:bold;color:' + gc + ';}'
        h += '.gb{background:' + gc + ';color:white;padding:8px 20px;border-radius:20px;font-weight:bold;display:inline-block;margin-top:10px;}'
        h += '.c{background:#f8f9fa;border-radius:12px;padding:16px;margin:12px 30px;border-left:5px solid #dee2e6;}'
        h += '.c.p1{border-left-color:#28a745;}.c.p2{border-left-color:#ffc107;}.c.p3{border-left-color:#dc3545;}'
        h += '.ch{display:flex;justify-content:space-between;margin-bottom:8px;}'
        h += '.ct{font-size:15px;font-weight:bold;}'
        h += '.bd{padding:3px 10px;border-radius:15px;font-size:12px;font-weight:bold;}'
        h += '.b1{background:#d4edda;color:#155724;}.b2{background:#fff3cd;color:#856404;}.b3{background:#f8d7da;color:#721c24;}'
        h += '.d{background:white;padding:10px;border-radius:8px;margin-top:8px;font-size:12px;line-height:1.8;max-height:200px;overflow-y:auto;}'
        h += '.ft{text-align:center;padding:15px;background:#f8f9fa;color:#6c757d;font-size:12px;margin-top:20px;}'
        h += '</style></head><body>'

        h += '<div class="box">'
        h += '<div class="hd"><h1>Word 2010 操作题评分报告（第2套）</h1>'
        h += '<p>基于XML底层解析 | 共11项评分要求 | v3</p></div>'
        h += '<div class="sc"><div style="text-align:center;">'
        h += '<div class="num">' + str(total) + '</div>'
        h += '<div style="color:#6c757d;font-size:16px;">/ ' + str(full) + ' 分</div>'
        h += '<div style="font-size:40px;color:' + gc + ';margin:10px 0;">' + ('%.1f' % rate) + '%</div>'
        h += '<div class="gb">' + gr + '</div>'
        h += '</div></div>'

        for r in self.results:
            sc = {'通过': 'p1', '部分通过': 'p2', '未通过': 'p3'}.get(r['status'], 'p3')
            bd = {'通过': 'b1', '部分通过': 'b2', '未通过': 'b3'}.get(r['status'], 'b3')
            h += '<div class="c ' + sc + '"><div class="ch">'
            h += '<span class="ct">要求' + ('%02d' % r['req']) + ' ' + r['name'] + ' (%d/%d)</span>' % (r['score'], r['full'])
            h += '<span class="bd ' + bd + '">' + r['status'] + '</span>'
            h += '</div><div class="d">'
            for d in r['detail']:
                h += '<div>' + d + '</div>'
            h += '</div></div>'

        h += '<div class="ft"><p>基于XML底层解析 | 精准对照11项要求</p>'
        h += '<p>生成: ' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + '</p></div>'
        h += '</div></body></html>'

        with open(path, 'w', encoding='utf-8') as f:
            f.write(h)
        print('\n报告已生成: ' + str(path))
        return str(path)


def main():
    docx = None
    if len(sys.argv) > 1:
        docx = sys.argv[1]
    else:
        for f in Path('.').glob('*结果*.docx'):
            docx = str(f); break
        if not docx:
            for f in Path('.').glob('*.docx'):
                docx = str(f); break
    if not docx:
        print('未找到.docx文件！')
        return
    print('目标文件: ' + docx)
    g = XMLGrader(docx)
    g.grade()
    rp = g.report()
    try:
        os.startfile(rp)
    except:
        pass


if __name__ == '__main__':
    main()
