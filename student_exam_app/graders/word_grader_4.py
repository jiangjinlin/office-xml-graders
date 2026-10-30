#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 2010 操作题 XML底层精准评分器 - 第4套题 v7
方法: 结合DOM解析 + XML字符串搜索，确保100%检测标准答案
"""

import sys, os, io, zipfile, datetime, shutil
from pathlib import Path
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import has_field, extract_word_parts

pass  # stdout redirect removed for web

# 命名空间URI
W  = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
A  = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
PIC = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'
R  = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
V  = '{urn:schemas-microsoft-com:vml}'
W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'


def get_text(elem):
    """获取元素内所有文本"""
    if elem is None: return ""
    parts = []
    for t in elem.iter(W + 't'):
        if t.text: parts.append(t.text)
    return ' '.join(parts)


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
        self.doc_xml = ""  # 序列化后的XML字符串

    def unzip(self):
        self.xml_dir = self.docx.parent / ('_xml_' + self.docx.stem)
        if self.xml_dir.exists():
            shutil.rmtree(self.xml_dir)
        extract_word_parts(self.docx, self.xml_dir)
        print('  解压完成:', self.xml_dir.name)

    def load_all(self):
        d = self.xml_dir / 'word'
        # 加载document.xml
        p = d / 'document.xml'
        if p.exists():
            self.doc_root = ET.parse(str(p)).getroot()
            self.doc_xml = ET.tostring(self.doc_root, encoding='unicode')
        p = d / 'settings.xml'
        if p.exists(): self.settings_root = ET.parse(str(p)).getroot()
        p = d / 'styles.xml'
        if p.exists(): self.styles_root = ET.parse(str(p)).getroot()
        # 加载页眉页脚，并把XML加入doc_xml（用于搜索水印等）
        for p in d.glob('header*.xml'):
            self.headers[p.name] = ET.parse(str(p)).getroot()
            self.doc_xml += ET.tostring(self.headers[p.name], encoding='unicode')
        for p in d.glob('footer*.xml'):
            self.footers[p.name] = ET.parse(str(p)).getroot()
            self.doc_xml += ET.tostring(self.footers[p.name], encoding='unicode')
        # 加载关系
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
        # 加载媒体文件列表
        mdir = d / 'media'
        if mdir.exists():
            self.media = [f.name for f in mdir.iterdir() if f.is_file()]
        # 加载diagrams文件（SmartArt文本等）
        diag_dir = self.xml_dir / 'word' / 'diagrams'
        if diag_dir.exists():
            for p in diag_dir.glob('*.xml'):
                self.doc_xml += ET.tostring(ET.parse(str(p)).getroot(), encoding='unicode')
        print('  已加载: document.xml, settings.xml, styles.xml')
        print('  %d 个页眉, %d 个页脚, %d 个媒体文件' % (len(self.headers), len(self.footers), len(self.media)))

    # ------------------------------------------------------------------
    # req1: 页面设置
    # ------------------------------------------------------------------
    def req1(self):
        log = []; score = 0
        for sp in self.doc_root.iter(W + 'sectPr'):
            # 页面大小
            pgSz = sp.find(W + 'pgSz')
            if pgSz is not None:
                w = int(pgSz.get(W + 'w', '0'))
                h = int(pgSz.get(W + 'h', '0'))
                code = pgSz.get(W + 'code', '')
                log.append('  页面大小: w=%d h=%d code=%s' % (w, h, code))
                if code == '9' or (abs(w - 11906) < 200 and abs(h - 16838) < 500):
                    score += 4; log.append('  OK A4纸张 (+4)')
            # 页边距
            pgMar = sp.find(W + 'pgMar')
            if pgMar is not None:
                top = int(pgMar.get(W + 'top', '0'))
                bot = int(pgMar.get(W + 'bottom', '0'))
                left = int(pgMar.get(W + 'left', '0'))
                right = int(pgMar.get(W + 'right', '0'))
                log.append('  页边距: top=%d bottom=%d left=%d right=%d' % (top, bot, left, right))
            # 页边距（A4与上下2.5cm为素材默认，不计分）
            pgMar = sp.find(W + 'pgMar')
            if pgMar is not None:
                left = int(pgMar.get(W + 'left', '0'))
                right = int(pgMar.get(W + 'right', '0'))
                log.append('  页边距: left=%d right=%d' % (left, right))
                if abs(left - 1701) < 50 and abs(right - 1701) < 50:
                    score += 6; log.append('  OK 左右3cm (+6)')
            break
        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 1, 'name': '页面设置', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req2: 装订线+文档网格+正文格式
    # ------------------------------------------------------------------
    def req2(self):
        log = []; score = 0
        for sp in self.doc_root.iter(W + 'sectPr'):
            pgMar = sp.find(W + 'pgMar')
            if pgMar is not None:
                gutter = int(pgMar.get(W + 'gutter', '0'))
                log.append('  装订线: %d' % gutter)
                if abs(gutter - 397) < 30:
                    score += 2; log.append('  OK 装订线0.7cm (+2)')
            docGrid = sp.find(W + 'docGrid')
            if docGrid is not None:
                lp = docGrid.get(W + 'linePitch', '')
                cs = docGrid.get(W + 'charSpace', '')
                log.append('  文档网格: linePitch=%s charSpace=%s' % (lp, cs))
                if lp and int(lp) == 340:
                    score += 3; log.append('  OK 每页约40行 (+3)')
                if cs and int(cs) < -400:
                    score += 3; log.append('  OK 每行约40字符 (+3)')
            break

        # 正文首行缩进2字符 + 行距固定值18磅
        paras = list(self.doc_root.iter(W + 'p'))
        if len(paras) > 1:
            target = paras[1]  # 第2段是正文第1段
            pPr = target.find(W + 'pPr')
            if pPr is not None:
                ind = pPr.find(W + 'ind')
                if ind is not None:
                    flc = ind.get(W + 'firstLineChars', '0')
                    log.append('  首行缩进: %s 字符' % flc)
                    if abs(int(flc) - 200) < 50:
                        score += 1; log.append('  OK 首行缩进2字符 (+1)')
                spacing = pPr.find(W + 'spacing')
                if spacing is not None:
                    line = spacing.get(W + 'line', '0')
                    rule = spacing.get(W + 'lineRule', '')
                    log.append('  行距: line=%s rule=%s' % (line, rule))
                    if line and int(line) == 360:
                        score += 1; log.append('  OK 行距固定值18磅 (+1)')

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 2, 'name': '装订线+文档网格+正文格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req3: 标题格式
    # ------------------------------------------------------------------
    def req3(self):
        log = []; score = 0
        paras = list(self.doc_root.iter(W + 'p'))
        if not paras:
            log.append('  NG 未找到段落')
            return {'req': 3, 'name': '标题格式', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}
        title = paras[0]
        txt = get_text(title)
        log.append('  标题文字: ' + txt[:30])
        if '元宇宙是风口' in txt:
            score += 1; log.append('  OK 标题文字正确 (+1)')

        pPr = title.find(W + 'pPr')
        if pPr is not None:
            # 对齐
            jc = pPr.find(W + 'jc')
            if jc is not None and pPr.find(W + 'jc').get(W + 'val', '') == 'center':
                score += 1; log.append('  OK 居中 (+1)')
            # 段前段后
            spacing = pPr.find(W + 'spacing')
            if spacing is not None:
                before = spacing.get(W + 'beforeLines', '')
                after = spacing.get(W + 'afterLines', '')
                if before == '50' and after == '50':
                    score += 1; log.append('  OK 段前段后0.5行 (+1)')
            # rPr格式
            rPr = pPr.find(W + 'rPr')
            if rPr is not None:
                # 字体
                rf = rPr.find(W + 'rFonts')
                if rf is not None and '魏' in rf.get(W + 'eastAsia', ''):
                    score += 2; log.append('  OK 华文新魏 (+2)')
                # 字号
                sz = rPr.find(W + 'sz')
                if sz is not None and sz.get(W + 'val', '') == '44':
                    score += 1; log.append('  OK 二号 (+1)')
                # 加粗
                if rPr.find(W + 'b') is not None:
                    score += 1; log.append('  OK 加粗 (+1)')
                # 颜色
                color = rPr.find(W + 'color')
                if color is not None and color.get(W + 'val', '').upper() == '0070C0':
                    score += 1; log.append('  OK 蓝色 (+1)')
                # 字符缩放
                w_elem = rPr.find(W + 'w')
                if w_elem is not None and w_elem.get(W + 'val', '') == '90':
                    score += 1; log.append('  OK 字符缩放90%% (+1)')
                # 文本效果（阴影）
                shadow = rPr.find(W14 + 'shadow')
                if shadow is not None:
                    score += 1; log.append('  OK 文本效果阴影 (+1)')

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 3, 'name': '标题格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req4: 替换文字+正文字体
    # ------------------------------------------------------------------
    def req4(self):
        log = []; score = 0
        if '源于周' not in self.doc_xml and '元宇宙' in self.doc_xml:
            score += 4; log.append('  OK 已替换"源于周"为"元宇宙" (+4)')
        elif '源于周' in self.doc_xml:
            log.append('  NG 未替换"源于周"')
        else:
            score += 4; log.append('  OK 无"源于周" (+4)')

        # 检查正文字体（楷体）和字号（小四=24）
        found_kaiti = False
        found_xiaosi = False
        paras = list(self.doc_root.iter(W + 'p'))
        for i, p in enumerate(paras):
            if i == 0: continue  # 跳过标题
            txt = get_text(p)
            if not txt.strip(): continue
            for run in p.findall('.//' + W + 'r'):
                rPr = run.find(W + 'rPr')
                if rPr is not None:
                    rf = rPr.find(W + 'rFonts')
                    if rf is not None and '楷' in rf.get(W + 'eastAsia', ''):
                        found_kaiti = True
                    sz = rPr.find(W + 'sz')
                    if sz is not None and sz.get(W + 'val', '') == '24':
                        found_xiaosi = True
            if found_kaiti and found_xiaosi:
                break

        if found_kaiti:
            score += 3; log.append('  OK 正文字体楷体 (+3)')
        if found_xiaosi:
            score += 3; log.append('  OK 正文小四号 (+3)')

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 4, 'name': '替换文字+正文字体', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req5: 段落边框底纹+分栏
    # ------------------------------------------------------------------
    def req5(self):
        log = []; score = 0
        paras = list(self.doc_root.iter(W + 'p'))
        if len(paras) > 1:
            target = paras[1]  # 第2段
            pPr = target.find(W + 'pPr')
            if pPr is not None:
                # 边框
                pBdr = pPr.find(W + 'pBdr')
                if pBdr is not None:
                    log.append('  OK 找到w:pBdr')
                    score += 2
                    for side in ['top', 'left', 'bottom', 'right']:
                        b = pBdr.find(W + side)
                        if b is not None:
                            sz = b.get(W + 'sz', '0')
                            col = b.get(W + 'color', '')
                            if sz == '12':
                                score += 1; log.append('  OK %s边框1.5磅 (+1)' % side)
                            if col.upper() == 'FF0000':
                                score += 1; log.append('  OK %s边框红色 (+1)' % side)
                # 底纹
                shd = pPr.find(W + 'shd')
                if shd is not None:
                    fill = shd.get(W + 'fill', '')
                    val = shd.get(W + 'val', '')
                    log.append('  底纹: fill=%s val=%s' % (fill, val))
                    if fill.upper() == 'FFFF00':
                        score += 2; log.append('  OK 黄色底纹 (+2)')
                    if val == 'pct20':
                        score += 1; log.append('  OK 20%%图案样式 (+1)')

        # 分栏
        found_cols = False
        for sp in self.doc_root.iter(W + 'sectPr'):
            cols = sp.find(W + 'cols')
            if cols is not None:
                num = cols.get(W + 'num', '1')
                sep = cols.get(W + 'sep', '0')
                log.append('  分栏(文档级): num=%s sep=%s' % (num, sep))
                if num == '3':
                    score += 2; log.append('  OK 三栏 (+2)')
                    found_cols = True
                    if sep == '1':
                        score += 1; log.append('  OK 分隔线 (+1)')
                break

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 5, 'name': '段落边框底纹+分栏', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req6: 超链接
    # ------------------------------------------------------------------
    def req6(self):
        log = []; score = 0
        # 方法1: 检查hyperlink关系
        for hl in self.doc_root.findall('.//' + W + 'hyperlink'):
            rid = hl.get(R + 'id', '')
            if rid in self.main_rels:
                target = self.main_rels[rid]['target']
                if target == 'http://baike.baidu.com/' or target == 'http://baike.baidu.com':
                    score += 10; log.append('  OK 百度百科超链接 (+10)')
                    break
        # 方法2: 检查域代码
        if score == 0:
            if 'HYPERLINK' in self.doc_xml.upper() and 'baike.baidu.com/' in self.doc_xml:
                score += 10; log.append('  OK 百度百科超链接(域代码) (+10)')
        # 方法3: 至少找到URL
        if score == 0:
            if 'baike.baidu.com' in self.doc_xml:
                score += 6; log.append('  OK 找到百度百科URL (+6)')

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 6, 'name': '超链接', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req7: 图片设置
    # ------------------------------------------------------------------
    def req7(self):
        log = []; score = 0
        # 搜索XML字符串找到图片尺寸
        import re
        # 查找 wp:extent 的 cx=... cy=...
        matches = re.findall(r'cx="(\d+)"\s*cy="(\d+)"', self.doc_xml)
        for cx_str, cy_str in matches:
            cx = int(cx_str) / 914400 * 2.54
            cy = int(cy_str) / 914400 * 2.54
                # 图片.jpg: 5.1cm x 5.0cm
            if abs(cx - 5.1) < 0.2 and abs(cy - 5.0) < 0.2:
                log.append('  找到图片: %.2fcm x %.2fcm' % (cx, cy))
                score += 6; log.append('  OK 尺寸5.1x5.0cm (+6)')
                # 检查环绕方式（搜索XML）
                if 'wrapSquare' in self.doc_xml:
                    score += 4; log.append('  OK 四周型环绕 (+4)')
                break

        if score == 0:
            log.append('  NG 未找到图片(5.1x5.0cm,四周型)')
        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 7, 'name': '图片设置', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req8: 图片样式 (使用XML字符串搜索）
    # ------------------------------------------------------------------
    def req8(self):
        log = []; score = 0
        # 搜索椭圆形状
        if 'ellipse' in self.doc_xml:
            score += 2; log.append('  OK 椭圆形状 (+2)')
        # 搜索柔化边缘
        if 'softEdge' in self.doc_xml:
            score += 2; log.append('  OK 柔化边缘 (+2)')
        # 搜索发光效果
        if 'glow' in self.doc_xml:
            score += 1; log.append('  OK 发光效果 (+1)')

        if score == 0:
            log.append('  NG 未检测到图片样式')
        if score >= 4: score = 5
        st = '通过' if score >= 4 else ('部分通过' if score >= 2 else '未通过')
        return {'req': 8, 'name': '图片样式', 'score': min(score, 5), 'full': 5, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req9: 文字水印 (使用XML字符串搜索）
    # ------------------------------------------------------------------
    def req9(self):
        log = []; score = 0
        # 搜索水印文字
        if '科技日报' in self.doc_xml and 'watermark' in self.doc_xml.lower():
            score += 2; log.append('  OK 水印文字=科技日报 (+2)')
        # 搜索斜式
        if 'rotation:315' in self.doc_xml or 'rotation=315' in self.doc_xml:
            score += 1; log.append('  OK 斜式 (+1)')
        # 搜索字号90
        if 'font-size:90pt' in self.doc_xml or 'font-size:90' in self.doc_xml:
            score += 1; log.append('  OK 字号90 (+1)')
        # 搜索浅蓝色
        if '00b0f0' in self.doc_xml.lower():
            score += 1; log.append('  OK 浅蓝色 (+1)')

        if score == 0:
            log.append('  NG 未检测到文字水印')
        if score >= 4: score = 5
        st = '通过' if score >= 4 else ('部分通过' if score >= 2 else '未通过')
        return {'req': 9, 'name': '文字水印', 'score': min(score, 5), 'full': 5, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req10: 页眉页脚
    # ------------------------------------------------------------------
    def req10(self):
        log = []; score = 0
        # 检查奇偶页不同
        if self.settings_root is not None:
            if self.settings_root.find('.//' + W + 'evenAndOddHeaders') is not None:
                score += 2; log.append('  OK 启用奇偶页不同 (+2)')

        # 检查页眉内容
        header_center = False
        odd_header = False
        even_header = False
        for nm, root in self.headers.items():
            txt = get_text(root)
            if '元宇宙' in txt and '中国' not in txt:
                odd_header = True; log.append('  OK 奇数页页眉"元宇宙" (+2)')
            if '中国科技网' in txt:
                even_header = True; log.append('  OK 偶数页页眉"中国科技网" (+2)')
            # 检查居中
            for p in root.findall('.//' + W + 'p'):
                pPr = p.find(W + 'pPr')
                if pPr is not None:
                    jc = pPr.find(W + 'jc')
                    if jc is not None and jc.get(W + 'val', '') == 'center':
                        header_center = True
                        break

        if odd_header: score += 2
        if even_header: score += 2
        if header_center:
            score += 2; log.append('  OK 页眉居中 (+2)')

        # 检查页脚
        for nm, root in self.footers.items():
            txt = get_text(root)
            if txt.strip().isdigit() or has_field(self.doc_xml, 'PAGE'):
                score += 2; log.append('  OK 页脚含页码 (+2)')
                break

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 10, 'name': '页眉页脚', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # req11: SmartArt (使用XML字符串搜索）
    # ------------------------------------------------------------------
    def req11(self):
        log = []; score = 0
        # 检查diagram引用
        has_diagram = False
        for rid, info in self.main_rels.items():
            if 'diagram' in info.get('target', ''):
                has_diagram = True
                log.append('  找到diagram引用: ' + info['target'])
                break

        if has_diagram:
            score += 2; log.append('  OK 包含SmartArt (+2)')
            # 搜索SmartArt尺寸 (14cm x 9cm)
            import re
            matches = re.findall(r'cx="(\d+)"\s*cy="(\d+)"', self.doc_xml)
            for cx_str, cy_str in matches:
                cx = int(cx_str) / 914400 * 2.54
                cy = int(cy_str) / 914400 * 2.54
                if abs(cx - 14) < 1 and abs(cy - 9) < 1:
                    score += 2; log.append('  OK SmartArt尺寸14x9cm (+2)')
                    break

            # 搜索环绕方式
            if 'wrapTopAndBottom' in self.doc_xml:
                score += 1; log.append('  OK 上下型环绕 (+1)')
            # 搜索对齐
            if 'positionH' in self.doc_xml and 'center' in self.doc_xml:
                score += 1; log.append('  OK 水平居中 (+1)')
            if 'positionV' in self.doc_xml and 'top' in self.doc_xml:
                score += 1; log.append('  OK 顶端对齐 (+1)')
            # 搜索文本"Metaverse"
            if 'Metaverse' in self.doc_xml:
                score += 2; log.append('  OK 文本=Metaverse (+2)')
            elif 'metaverse' in self.doc_xml.lower():
                score += 1; log.append('  OK 找到metaverse文本 (+1)')

        if not has_diagram:
            log.append('  NG 未找到SmartArt')
        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 11, 'name': 'SmartArt', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ------------------------------------------------------------------
    # 评分主流程
    # ------------------------------------------------------------------
    def grade(self):
        print('\n' + '=' * 60)
        print('  Word 2010 XML底层精准评分器 - 第4套题 v7')
        print('=' * 60)
        self.unzip()
        print('\n加载XML...')
        self.load_all()
        print('\n' + '-' * 60)
        print('  逐项检测...')
        print('-' * 60)

        checks = [self.req1, self.req2, self.req3, self.req4, self.req5,
                  self.req6, self.req7, self.req8, self.req9, self.req10, self.req11]
        names = ['页面设置', '装订线+文档网格+正文格式', '标题格式', '替换文字+正文字体',
                 '段落边框底纹+分栏', '超链接', '图片设置', '图片样式',
                 '文字水印', '页眉页脚', 'SmartArt']

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
        if path is None: path = self.docx.parent / 'report_4.html'
        total = sum(r['score'] for r in self.results)
        full = sum(r['full'] for r in self.results)
        rate = total / full * 100 if full > 0 else 0
        if rate >= 90: gr, gc = 'A 优秀', '#28a745'
        elif rate >= 75: gr, gc = 'B 良好', '#17a2b8'
        elif rate >= 60: gr, gc = 'C 及格', '#ffc107'
        else: gr, gc = 'D 不及格', '#dc3545'

        h = '<!DOCTYPE html><html><head><meta charset="UTF-8">'
        h += '<title>Word评分报告 - 第4套题</title>'
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
        h += '<div class="hd"><h1>Word 2010 操作题评分报告（第4套）</h1>'
        h += '<p>基于XML底层解析 | 共11项评分要求 | v7</p></div>'
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
