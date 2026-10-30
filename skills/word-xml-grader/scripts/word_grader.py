#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 2010 操作题 XML底层精准评分器 v7.0
修复: req9(题注居中-检查所有题注段落)
"""

import sys, os, io, zipfile, datetime, shutil
from pathlib import Path
import xml.etree.ElementTree as ET

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# 命名空间
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
    return ' + '.join(parts)


class XMLGrader:
    def __init__(self, docx_path):
        self.docx = Path(docx_path)
        self.xmldir = None
        self.results = []
        self.doc_root = None
        self.settings_root = None
        self.footnotes_root = None
        self.headers = {}
        self.footers = {}
        self.rels = {}
        self.main_rels = {}
        self.media = []

    def unzip(self):
        self.xmldir = self.docx.parent / ('_xml_' + self.docx.stem)
        if self.xmldir.exists():
            shutil.rmtree(self.xmldir)
        with zipfile.ZipFile(self.docx, 'r') as z:
            z.extractall(self.xmldir)
        print('  解压完成:', self.xmldir.name)

    def load_all(self):
        d = self.xmldir / 'word'
        p = d / 'document.xml'
        if p.exists(): self.doc_root = ET.parse(str(p)).getroot()
        p = d / 'settings.xml'
        if p.exists(): self.settings_root = ET.parse(str(p)).getroot()
        p = d / 'footnotes.xml'
        if p.exists(): self.footnotes_root = ET.parse(str(p)).getroot()
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
        print('  已加载: document.xml, settings.xml')
        print('  %d 个页眉, %d 个页脚, %d 个媒体文件' % (len(self.headers), len(self.footers), len(self.media)))

    def req1(self):
        log = []; score = 0
        paras = XA(self.doc_root, 'w:body/w:p')
        if not paras:
            return {'req': 1, 'name': '标题格式', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}
        title = paras[0]
        rPr = X(title, 'w:pPr/w:rPr')
        if rPr is None:
            runs = XA(title, 'w:r')
            if runs: rPr = X(runs[0], 'w:rPr')
        if rPr is None:
            log.append('  未找到rPr')
            return {'req': 1, 'name': '标题格式', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}
        rf = X(rPr, 'w:rFonts')
        if rf is not None:
            ea = GA(rf, 'w:eastAsia', '')
            log.append('  字体: ' + ea)
            if '行楷' in ea: score += 2; log.append('  OK 华文行楷 (+2)')
        if X(rPr, 'w:b') is not None: score += 2; log.append('  OK 加粗 (+2)')
        if X(rPr, 'w:i') is not None: score += 1; log.append('  OK 倾斜 (+1)')
        sz = X(rPr, 'w:sz')
        if sz is not None:
            v = int(GA(sz, 'w:val', '0'))
            log.append('  字号: ' + str(v) + ' (' + str(v // 2) + 'pt)')
            if v == 44: score += 2; log.append('  OK 二号22pt (+2)')
        co = X(rPr, 'w:color')
        if co is not None:
            v = GA(co, 'w:val', '')
            log.append('  颜色: ' + v)
            if v.upper() in ['C00000', '800000']: score += 2; log.append('  OK 深红色 (+2)')
        # 所有主要格式都正确，给满分10
        if score >= 9: score = 10
        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 1, 'name': '标题格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    def req2(self):
        log = []; score = 0
        paras = XA(self.doc_root, 'w:body/w:p')
        if not paras:
            return {'req': 2, 'name': '标题边框间距', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}
        title = paras[0]
        pPr = X(title, 'w:pPr')
        if pPr is None:
            return {'req': 2, 'name': '标题边框间距', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}
        rPr = X(pPr, 'w:rPr')
        bdr = None
        if rPr is not None: bdr = X(rPr, 'w:bdr')
        if bdr is None:
            for run in XA(title, 'w:r'):
                rp = X(run, 'w:rPr')
                if rp is not None:
                    bdr = X(rp, 'w:bdr')
                    if bdr is not None: break
        if bdr is not None:
            log.append('  OK 找到w:bdr(文字方框)')
            v = GA(bdr, 'w:val', '')
            sz = GA(bdr, 'w:sz', '0')
            col = GA(bdr, 'w:color', '')
            log.append('  样式=' + v + ' 线宽=' + sz + '(1/8pt) 颜色=' + col)
            if v in ['single', 'thick']: score += 2; log.append('  OK 单实线 (+2)')
            if abs(int(sz) - 12) < 4: score += 2; log.append('  OK 1.5磅 (+2)')
            if '00' in col.upper() and ('C0' in col.upper() or 'BLUE' in col.upper()):
                score += 2; log.append('  OK 蓝色 (+2)')
        else:
            log.append('  NG 未找到w:bdr')
        sp = X(pPr, 'w:spacing')
        if sp is not None:
            bl = GA(sp, 'w:beforeLines', '')
            al = GA(sp, 'w:afterLines', '')
            log.append('  段前: ' + bl + ' 段后: ' + al)
            if bl == '50' or (bl and int(bl) <= 60): score += 2; log.append('  OK 段前0.5行 (+2)')
            if al == '100' or (al and int(al) >= 80): score += 2; log.append('  OK 段后1行 (+2)')
        jc = X(pPr, 'w:jc')
        if jc is not None and GA(jc, 'w:val', '') == 'center':
            score += 2; log.append('  OK 居中 (+2)')
        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 2, 'name': '标题边框间距', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    def req3(self):
        log = []; score = 0
        paras = XA(self.doc_root, 'w:body/w:p')
        body = paras[1:] if len(paras) > 1 else []
        f_font = 0; f_sz = 0; f_ind = 0; f_ln = 0; tot = 0

        for para in body[:8]:
            tot += 1
            pPr = X(para, 'w:pPr')
            if pPr is None: continue
            # 字体 - 先检查pPr/rPr（段落默认）
            rp_p = X(pPr, 'w:rPr')
            if rp_p is not None:
                rf = X(rp_p, 'w:rFonts')
                if rf is not None and '中宋' in GA(rf, 'w:eastAsia', ''):
                    f_font += 1
            # 字号 - 先检查pPr/rPr
            if rp_p is not None:
                sz = X(rp_p, 'w:sz')
                if sz is not None and int(GA(sz, 'w:val', '0')) == 28:
                    f_sz += 1
            # 首行缩进
            ind = X(pPr, 'w:ind')
            if ind is not None:
                flc = GA(ind, 'w:firstLineChars', '0')
                if int(flc) >= 150: f_ind += 1
            # 行距
            sp = X(pPr, 'w:spacing')
            if sp is not None:
                ln = GA(sp, 'w:line', '0')
                rule = GA(sp, 'w:lineRule', '')
                if int(ln) == 400 and rule == 'exact': f_ln += 1

        log.append('  字体OK:' + str(f_font) + '/' + str(tot))
        log.append('  字号OK:' + str(f_sz) + '/' + str(tot))
        log.append('  缩进OK:' + str(f_ind) + '/' + str(tot))
        log.append('  行距OK:' + str(f_ln) + '/' + str(tot))

        if f_font > 0: score += 3; log.append('  OK 华文中宋 (+3)')
        if f_sz > 0: score += 2; log.append('  OK 四号 (+2)')
        if f_ind >= 3: score += 3; log.append('  OK 首行缩进 (+3)')
        elif f_ind > 0: score += 1
        if f_ln >= 3: score += 2; log.append('  OK 行距20磅 (+2)')
        elif f_ln > 0: score += 1

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 3, 'name': '正文格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    def req4(self):
        log = []; score = 0
        drawings = XA(self.doc_root, 'w:body//w:drawing')
        log.append('  找到drawing: ' + str(len(drawings)))

        for i, dwg in enumerate(drawings):
            ext = dwg.find('.//{' + WPNS + '}extent')
            if ext is not None:
                cx = int(GA(ext, 'cx', '0'))
                cy = int(GA(ext, 'cy', '0'))
                wcm = cx / 914400 * 2.54
                hcm = cy / 914400 * 2.54
                log.append('  图%d: %.1fcm x %.1fcm' % (i + 1, wcm, hcm))
                if abs(wcm - 14.5) < 1 and abs(hcm - 5) < 1:
                    score += 3; log.append('  OK 尺寸14.5x5cm (+3)')
                    # 检查环绕方式
                    if dwg.find('.//{' + WPNS + '}wrapTight') is not None:
                        score += 2; log.append('  OK 紧密型环绕 (+2)')
                    # 检查旋转
                    xfrm = dwg.find('.//{' + ANS + '}xfrm')
                    if xfrm is not None:
                        rot = int(GA(xfrm, 'rot', '0'))
                        deg = rot / 60000.0
                        log.append('  旋转: %.1f度' % deg)
                        if abs(deg - 180) < 10:
                            score += 2; log.append('  OK 旋转180度 (+2)')
                    # 检查水平对齐
                    pH = dwg.find('.//{' + WPNS + '}positionH')
                    if pH is not None:
                        rf = pH.get('relativeFrom', '')
                        al = pH.find('{' + WPNS + '}align')
                        if al is not None and rf == 'column' and al.text == 'center':
                            score += 1; log.append('  OK 相对于栏居中 (+1)')
                    break

        if score == 0: log.append('  NG 未找到符合要求的图片')
        # 所有主要属性都正确，给满分10
        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 4, 'name': '第3段图片', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    def req5(self):
        log = []; score = 0
        for sp in XA(self.doc_root, 'w:body//w:sectPr'):
            fp = X(sp, 'w:footnotePr')
            if fp is not None:
                nf = X(fp, 'w:numFmt')
                if nf is not None:
                    v = GA(nf, 'w:val', '')
                    log.append('  编号格式: ' + v)
                    if 'Circle' in v or 'enclosed' in v.lower():
                        score += 2; log.append('  OK 圆圈数字 (+2)')
        refs = XA(self.doc_root, 'w:body//w:footnoteReference')
        if refs: score += 1; log.append('  OK 有%d个脚注引用 (+1)' % len(refs))
        if self.footnotes_root is not None:
            score += 1; log.append('  OK 找到footnotes.xml')
            xs = ET.tostring(self.footnotes_root, encoding='unicode')
            if 'http' in xs.lower():
                score += 2; log.append('  OK 包含URL (+2)')
        st = '通过' if score >= 4 else ('部分通过' if score >= 2 else '未通过')
        return {'req': 5, 'name': '脚注设置', 'score': min(score, 5), 'full': 5, 'detail': log, 'status': st}

    def req6(self):
        log = []; score = 0
        found_watermark = False

        for nm, root in self.headers.items():
            log.append('  检查 ' + nm)
            for shape in root.findall('.//{' + VNS + '}shape'):
                sid = shape.get('id', '')
                if 'Watermark' in sid or 'watermark' in sid:
                    log.append('  OK 找到水印VML形状: ' + sid)
                    score += 3
                    found_watermark = True
                    for imgdata in shape.findall('{' + VNS + '}imagedata'):
                        blacklevel = imgdata.get('blacklevel', '')
                        gain = imgdata.get('gain', '')
                        if blacklevel or gain:
                            score += 2; log.append('  OK 冲蚀效果 (+2)')
                            break
                    break
            if found_watermark:
                break
            xs = ET.tostring(root, encoding='unicode')
            if not found_watermark and ('Watermark' in xs or 'watermark' in xs):
                log.append('  OK ' + nm + ' 含水印关键字')
                score += 3
                if 'blacklevel' in xs or 'gain=' in xs:
                    score += 2; log.append('  OK 冲蚀效果 (+2)')
                found_watermark = True
                break

        if not found_watermark:
            for m in self.media:
                if '水印' in m or 'watermark' in m.lower():
                    log.append('  OK 找到水印图片: ' + m)
                    score += 3
                    break

        if score == 0: log.append('  NG 未检测到图片水印')
        st = '通过' if score >= 4 else ('部分通过' if score >= 2 else '未通过')
        return {'req': 6, 'name': '图片水印', 'score': min(score, 5), 'full': 5, 'detail': log, 'status': st}

    def req7(self):
        log = []; score = 0
        paras = XA(self.doc_root, 'w:body/w:p')

        found_indent = False
        for idx in range(1, min(6, len(paras))):
            para = paras[idx]
            pPr = X(para, 'w:pPr')
            if pPr is None: continue
            ind = X(pPr, 'w:ind')
            if ind is not None:
                lc = GA(ind, 'w:leftChars', '0')
                rc = GA(ind, 'w:rightChars', '0')
                log.append('  段%d: leftChars=%s rightChars=%s' % (idx + 1, lc, rc))
                if int(lc) >= 150 and int(rc) >= 150:
                    found_indent = True
                    score += 4; log.append('  OK 左右缩进2字符 (+4)')
                    break

        if not found_indent:
            log.append('  NG 未找到含左右缩进的段落')

        found_cols = False
        for sp in XA(self.doc_root, 'w:body//w:sectPr'):
            cols = X(sp, 'w:cols')
            if cols is not None:
                num = GA(cols, 'w:num', '1')
                sep = GA(cols, 'w:sep', '0')
                ew = GA(cols, 'w:equalWidth', '1')
                log.append('  分栏: num=' + num + ' sep=' + sep + ' equal=' + ew)
                if num == '2':
                    score += 2; log.append('  OK 两栏 (+2)')
                    found_cols = True
                    if sep == '1':
                        score += 1; log.append('  OK 分隔线 (+1)')
                    if ew == '0':
                        score += 1; log.append('  OK 不等宽(偏左) (+1)')
                    for j, col in enumerate(XA(cols, 'w:col')):
                        w = int(GA(col, 'w:w', '0'))
                        log.append('    栏%d宽=%d' % (j + 1, w))
                        if j == 0 and abs(w - 2940) < 500:
                            score += 2; log.append('  OK 第1栏宽14字符 (+2)')
                    break

        if not found_cols:
            log.append('  NG 未找到分栏设置(num=2)')

        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 7, 'name': '第2段分栏', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    def req8(self):
        log = []; score = 0
        if self.settings_root is not None:
            if X(self.settings_root, 'w:evenAndOddHeaders') is not None:
                score += 3; log.append('  OK 启用奇偶页不同 (+3)')
        for nm, root in self.headers.items():
            txt = GT(root)
            log.append('  ' + nm + ': ' + (txt[:30] if len(txt) > 30 else txt))
            if '体育' in txt: score += 3; log.append('  OK 偶数页眉=体育精神 (+3)')
            if '冬奥' in txt: score += 3; log.append('  OK 奇数页眉=冬奥会 (+3)')
        for nm, root in self.footers.items():
            xs = ET.tostring(root, encoding='unicode')
            txt = GT(root)
            log.append('  ' + nm + ': ' + (txt[:30] if len(txt) > 30 else txt))
            if 'PAGE' in xs or 'page' in xs.lower():
                score += 3; log.append('  OK 页脚含PAGE域 (+3)')
            if '<w:b' in xs:
                score += 2; log.append('  OK 页码加粗 (+2)')
                break
        st = '通过' if score >= 11 else ('部分通过' if score >= 5 else '未通过')
        return {'req': 8, 'name': '页眉页脚', 'score': min(score, 15), 'full': 15, 'detail': log, 'status': st}

    def req9(self):
        log = []; score = 0
        drawings = XA(self.doc_root, 'w:body//w:drawing')

        found = False
        for i, dwg in enumerate(drawings):
            cNv = dwg.find('.//{' + PICNS + '}cNvPr')
            if cNv is not None:
                nm = GA(cNv, 'name', '')
                log.append('  图%d name=%s' % (i + 1, nm))
                if '冰墩' in nm or '雪容' in nm:
                    found = True; score += 3
                    log.append('  OK 冰墩墩图片 (+3)')
                    ext = dwg.find('.//{' + WPNS + '}extent')
                    if ext is not None:
                        cx = int(GA(ext, 'cx', '0')) / 914400 * 2.54
                        cy = int(GA(ext, 'cy', '0')) / 914400 * 2.54
                        if abs(cx - 7) < 1 and abs(cy - 5) < 1:
                            score += 2; log.append('  OK 尺寸5x7cm (+2)')
                    if dwg.find('.//{' + WPNS + '}wrapTight') is not None:
                        score += 1; log.append('  OK 紧密环绕 (+1)')
                    break

        if not found:
            log.append('  NG 未找到冰墩墩图片')

        # 题注 SEQ域 - 检查所有含SEQ和图片的段落，找居中
        found_caption = False
        for para in XA(self.doc_root, 'w:body//w:p'):
            xs = ET.tostring(para, encoding='unicode')
            if 'SEQ' in xs and '图片' in xs:
                found_caption = True
                score += 2; log.append('  OK 题注标签=图片 (+2)')
                # 检查居中 - 检查所有题注段落
                jc = X(para, 'w:pPr/w:jc')
                if jc is not None and GA(jc, 'w:val', '') == 'center':
                    score += 2; log.append('  OK 题注居中 (+2)')
                    break  # 找到居中的题注就停
                # 如果这段没居中，继续检查下一个题注段落

        if not found_caption:
            log.append('  NG 未找到题注')
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 9, 'name': '图片题注', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    def req10(self):
        log = []; score = 0
        for sp in XA(self.doc_root, 'w:body//w:sectPr'):
            pgb = X(sp, 'w:pgBorders')
            if pgb is not None:
                log.append('  OK 找到w:pgBorders')
                score += 3
                has_shadow = False
                for side in ['top', 'left', 'bottom', 'right']:
                    b = X(pgb, 'w:' + side)
                    if b is not None:
                        sz = GA(b, 'w:sz', '0')
                        col = GA(b, 'w:color', '')
                        sh = GA(b, 'w:shadow', '0')
                        log.append('    %s: sz=%s col=%s shadow=%s' % (side, sz, col, sh))
                        if sh == '1': has_shadow = True
                if has_shadow: score += 3; log.append('  OK 阴影边框 (+3)')
                break

        bg = X(self.doc_root, 'w:background')
        if bg is not None:
            score += 2; log.append('  OK 找到w:background')
            xs = ET.tostring(bg, encoding='unicode')
            if 'grad' in xs.lower() or 'fill' in xs:
                score += 2; log.append('  OK 渐变背景 (+2)')
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 10, 'name': '页面边框背景', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    def req11(self):
        log = []; score = 0
        hls = XA(self.doc_root, 'w:body//w:hyperlink')
        log.append('  超链接数: ' + str(len(hls)))
        for hl in hls:
            rid = GA(hl, 'r:id', '')
            txt = GT(hl)
            log.append('    文字=' + txt + ' rid=' + rid)
            if '丝绸' in txt: score += 2; log.append('  OK 丝绸之路 (+2)')
            if rid in self.main_rels:
                tgt = self.main_rels[rid]['target']
                log.append('    目标=' + tgt)
                if 'baidu' in tgt.lower() or 'baike' in tgt.lower():
                    score += 3; log.append('  OK 百度百科 (+3)')
        if not hls: log.append('  NG 未找到超链接')
        st = '通过' if score >= 4 else ('部分通过' if score >= 2 else '未通过')
        return {'req': 11, 'name': '超链接', 'score': min(score, 5), 'full': 5, 'detail': log, 'status': st}

    def grade(self):
        print('\n' + '=' * 60)
        print('  Word 2010 XML底层精准评分 v7.0')
        print('=' * 60)
        self.unzip()
        print('\n加载XML...')
        self.load_all()
        print('\n' + '-' * 60)
        print('  逐项检测...')
        print('-' * 60)

        checks = [self.req1, self.req2, self.req3, self.req4, self.req5,
                  self.req6, self.req7, self.req8, self.req9, self.req10, self.req11]
        names = ['标题格式', '标题边框间距', '正文格式', '第3段图片', '脚注设置',
                 '图片水印', '第2段分栏', '页眉页脚', '图片题注', '页面边框背景', '超链接']

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
        if path is None: path = self.docx.parent / 'report_xml_v7.html'
        total = sum(r['score'] for r in self.results)
        full = sum(r['full'] for r in self.results)
        rate = total / full * 100 if full > 0 else 0
        if rate >= 90: gr, gc = 'A 优秀', '#28a745'
        elif rate >= 75: gr, gc = 'B 良好', '#17a2b8'
        elif rate >= 60: gr, gc = 'C 及格', '#ffc107'
        else: gr, gc = 'D 不及格', '#dc3545'

        h = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
        h += '<title>XML精准评分报告 v7</title>'
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
        h += '<div class="hd"><h1>Word 2010 XML底层精准评分报告</h1>'
        h += '<p>直接解析docx内部XML | v7.0 | 修复题注居中检测</p></div>'
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
    import sys
    docx = None
    if len(sys.argv) > 1:
        docx = sys.argv[1]
    else:
        for f in Path('.').glob('结果*.docx'):
            docx = str(f); break
    if not docx:
        print('未找到结果.docx！')
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
