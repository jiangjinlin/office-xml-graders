#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 2010 操作题 XML底层精准评分器 - 第5套题 v1
方法: 先分析标准答案XML → 理解结构 → 写精准检测代码
"""

import sys, os, io, zipfile, datetime, shutil
from pathlib import Path
import xml.etree.ElementTree as ET
import re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import extract_word_parts

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
    if elem is None: return ""
    return ''.join(t.text or '' for t in elem.iter(W + 't'))


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
        self.doc_xml = ""

    def unzip(self):
        self.xml_dir = self.docx.parent / ('_xml_' + self.docx.stem)
        if self.xml_dir.exists():
            shutil.rmtree(self.xml_dir)
        extract_word_parts(self.docx, self.xml_dir)
        print('  解压完成:', self.xml_dir.name)

    def load_all(self):
        d = self.xml_dir / 'word'
        # document.xml
        p = d / 'document.xml'
        if p.exists():
            self.doc_root = ET.parse(str(p)).getroot()
            self.doc_xml = ET.tostring(self.doc_root, encoding='unicode')
        # settings.xml
        p = d / 'settings.xml'
        if p.exists(): self.settings_root = ET.parse(str(p)).getroot()
        # styles.xml
        p = d / 'styles.xml'
        if p.exists(): self.styles_root = ET.parse(str(p)).getroot()
        # 页眉页脚（水印在header中！）
        for p in d.glob('header*.xml'):
            self.headers[p.name] = ET.parse(str(p)).getroot()
            self.doc_xml += ET.tostring(self.headers[p.name], encoding='unicode')
        for p in d.glob('footer*.xml'):
            self.footers[p.name] = ET.parse(str(p)).getroot()
            self.doc_xml += ET.tostring(self.footers[p.name], encoding='unicode')
        # 关系文件
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
        # 媒体文件
        mdir = d / 'media'
        if mdir.exists():
            self.media = [f.name for f in mdir.iterdir() if f.is_file()]
        # diagrams (SmartArt文本)
        diag_dir = d / 'diagrams'
        if diag_dir and diag_dir.exists():
            for p in diag_dir.glob('*.xml'):
                self.doc_xml += ET.tostring(ET.parse(str(p)).getroot(), encoding='unicode')
        print('  已加载: document.xml, settings.xml, styles.xml')
        print('  %d 个页眉, %d 个页脚, %d 个媒体文件' % (len(self.headers), len(self.footers), len(self.media)))

    # ==================================================================
    # req1: 页面设置 (A4, 上下2.5cm, 左右3cm)
    # 标准答案: w=11906 h=16838 code=9 top=1418 bot=1418 left=1701 right=1701
    # ==================================================================
    def req1(self):
        log = []; score = 0
        for sp in self.doc_root.iter(W + 'sectPr'):
            pgMar = sp.find(W + 'pgMar')
            if pgMar is not None:
                left = int(pgMar.get(W + 'left', '0'))
                right = int(pgMar.get(W + 'right', '0'))
                log.append('  页边距: left=%d right=%d' % (left, right))
                if abs(left - 1701) < 50 and abs(right - 1701) < 50:
                    score += 7; log.append('  OK 左右3cm (+7)')
            docGrid = sp.find(W + 'docGrid')
            if docGrid is not None and docGrid.get(W + 'linePitch', '') == '333':
                score += 3; log.append('  OK 每页42行 (+3)')
            break

        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 1, 'name': '页面设置', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ==================================================================
    # req2: 装订线0.6cm + 文档网格42行40字 + 正文首行缩进
    # 标准答案: gutter=340, linePitch=333, charSpace=-1209,
    #   正文缩进继承自Normal样式(a): firstLine=200(2字符)
    # ==================================================================
    def req2(self):
        log = []; score = 0
        for sp in self.doc_root.iter(W + 'sectPr'):
            pgMar = sp.find(W + 'pgMar')
            if pgMar is not None:
                gutter = int(pgMar.get(W + 'gutter', '0'))
                log.append('  装订线: %d' % gutter)
                if abs(gutter - 340) < 30:
                    score += 2; log.append('  OK 装订线0.6cm (+2)')

            docGrid = sp.find(W + 'docGrid')
            if docGrid is not None:
                lp = docGrid.get(W + 'linePitch', '')
                cs = docGrid.get(W + 'charSpace', '')
                log.append('  文档网格: linePitch=%s charSpace=%s' % (lp, cs))
                if lp and abs(int(lp) - 333) <= 3:
                    score += 3; log.append('  OK 每页约42行 (+3)')
                if cs and int(cs) < -400:
                    score += 3; log.append('  OK 每行约40字符 (+3)')
            break

        # 正文首行缩进（标准答案: Normal样式'a'定义firstLine=200=2字符）
        paras = list(self.doc_root.iter(W + 'p'))
        if len(paras) > 1:
            pPr = paras[1].find(W + 'pPr')
            if pPr is not None:
                ind = pPr.find(W + 'ind')
                if ind is not None:
                    flc = ind.get(W + 'firstLineChars', ind.get(W + 'firstLine', '0'))
                    fl = int(flc) if flc else 0
                    log.append('  首行缩进: %d' % fl)
                    if fl > 0:  # 有缩进即可
                        score += 1; log.append('  OK 首行缩进 (+1)')
                spacing = pPr.find(W + 'spacing')
                if spacing is not None:
                    line = spacing.get(W + 'line', '0')
                    rule = spacing.get(W + 'lineRule', '')
                    log.append('  行距: line=%s rule=%s' % (line, rule))
                    if line and int(line) == 300:
                        score += 1; log.append('  OK 行距多倍行距 (+1)')

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 2, 'name': '装订线+文档网格+正文格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ==================================================================
    # req3: 标题格式
    # 标准答案: "《关于推动现代职业教育高质量发展的意见》", 华文中宋, sz=36, 居中, 333333
    # ==================================================================
    def req3(self):
        log = []; score = 0
        paras = list(self.doc_root.iter(W + 'p'))
        if not paras:
            log.append('  NG 未找到段落')
            return {'req': 3, 'name': '标题格式', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}

        title = paras[0]
        txt = get_text(title)
        log.append('  标题文字: ' + txt[:40])

        pPr = title.find(W + 'pPr')
        if pPr is not None:
            # 居中
            jc = pPr.find(W + 'jc')
            if jc is not None and jc.get(W + 'val', '') == 'center':
                score += 4; log.append('  OK 居中 (+4)')

            rPr = pPr.find(W + 'rPr')
            if rPr is not None:
                # 字体: 华文中宋
                rf = rPr.find(W + 'rFonts')
                if rf is not None:
                    ea = rf.get(W + 'eastAsia', '')
                    if '华文中宋' in ea or '中宋' in ea:
                        score += 3; log.append('  OK 华文中宋 (+3)')
                # 字号: sz=36 (小二号, 18pt)
                sz = rPr.find(W + 'sz')
                if sz is not None:
                    sval = sz.get(W + 'val', '')
                    if sval == '36':
                        score += 3; log.append('  OK 字号sz=36 (+3)')

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 3, 'name': '标题格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ==================================================================
    # req4: 标题一格式
    # 标准答案: style=1, 在styles.xml中定义 华文中宋, C00000, sz=32,
    #    spacing before=100 after=50, line=240 atLeast
    # ==================================================================
    def req4(self):
        log = []; score = 0
        paras = list(self.doc_root.iter(W + 'p'))
        has_h1 = False
        for p in paras:
            pPr = p.find(W + 'pPr')
            if pPr is None: continue
            pStyle = pPr.find(W + 'pStyle')
            if pStyle is None: continue
            if pStyle.get(W + 'val', '') == '1':
                has_h1 = True
                txt = get_text(p)
                log.append('  标题一段落: ' + txt[:30])
                break

        if not has_h1:
            log.append('  NG 未找到标题一格式段落')
            return {'req': 4, 'name': '标题一格式', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}

        score += 2; log.append('  OK 使用标题一样式 (+2)')

        # 在styles.xml中检查标题一样式的定义
        if self.styles_root is not None:
            for style in self.styles_root.findall(W + 'style'):
                if style.get(W + 'styleId', '') == '1':
                    rPr = style.find(W + 'rPr')
                    if rPr is not None:
                        # 字体: 华文中宋
                        rf = rPr.find(W + 'rFonts')
                        if rf is not None:
                            ea = rf.get(W + 'eastAsia', '')
                            if '中宋' in ea:
                                score += 2; log.append('  OK 华文中宋 (+2)')
                        # 颜色: C00000
                        color = rPr.find(W + 'color')
                        if color is not None and color.get(W + 'val', '').upper() == 'C00000':
                            score += 2; log.append('  OK 颜色C00000 (+2)')
                        # 字号: sz=32
                        sz = rPr.find(W + 'sz')
                        if sz is not None and sz.get(W + 'val', '') == '32':
                            score += 2; log.append('  OK 字号sz=32 (+2)')
                    # 间距: before=100, after=50, line=240 atLeast
                    pPr_style = style.find(W + 'pPr')
                    if pPr_style is not None:
                        spacing = pPr_style.find(W + 'spacing')
                        if spacing is not None:
                            before = spacing.get(W + 'before', '')
                            after = spacing.get(W + 'after', '')
                            line = spacing.get(W + 'line', '')
                            rule = spacing.get(W + 'lineRule', '')
                            log.append('  样式间距: before=%s after=%s line=%s rule=%s' % (before, after, line, rule))
                            if before == '100' and after == '50':
                                score += 1; log.append('  OK 段前1行段后0.5行 (+1)')
                            if rule == 'atLeast':
                                score += 1; log.append('  OK 最小值行距 (+1)')
                    break

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 4, 'name': '标题一格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ==================================================================
    # req5: 标题二格式
    # 标准答案: style=2, styles.xml定义 宋体, FF0000, bold=false, spacing before=240, atLeast
    # ==================================================================
    def req5(self):
        log = []; score = 0
        paras = list(self.doc_root.iter(W + 'p'))
        has_h2 = False
        for p in paras:
            pPr = p.find(W + 'pPr')
            if pPr is None: continue
            pStyle = pPr.find(W + 'pStyle')
            if pStyle is None: continue
            if pStyle.get(W + 'val', '') == '2':
                has_h2 = True
                txt = get_text(p)
                log.append('  标题二段落: ' + txt[:30])
                break

        if not has_h2:
            log.append('  NG 未找到标题二格式段落')
            return {'req': 5, 'name': '标题二格式', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}

        score += 2; log.append('  OK 使用标题二样式 (+2)')

        # 在styles.xml中检查标题二样式的定义
        if self.styles_root is not None:
            for style in self.styles_root.findall(W + 'style'):
                if style.get(W + 'styleId', '') == '2':
                    rPr = style.find(W + 'rPr')
                    if rPr is not None:
                        # 字体: 宋体
                        rf = rPr.find(W + 'rFonts')
                        if rf is not None:
                            ea = rf.get(W + 'eastAsia', '')
                            if '宋体' in ea:
                                score += 2; log.append('  OK 宋体 (+2)')
                        # 颜色: FF0000 (红色)
                        color = rPr.find(W + 'color')
                        if color is not None and color.get(W + 'val', '').upper() == 'FF0000':
                            score += 2; log.append('  OK 颜色FF0000 (+2)')
                        # 加粗: b=false (表示不加粗，但有 bCs)
                        b_elem = rPr.find(W + 'b')
                        bCs = rPr.find(W + 'bCs') is not None
                        if b_elem is not None and bCs:
                            score += 2; log.append('  OK 加粗设置 (+2)')
                    # 间距: before=240, atLeast
                    pPr_style = style.find(W + 'pPr')
                    if pPr_style is not None:
                        spacing = pPr_style.find(W + 'spacing')
                        if spacing is not None:
                            before = spacing.get(W + 'before', '')
                            line = spacing.get(W + 'line', '')
                            rule = spacing.get(W + 'lineRule', '')
                            log.append('  样式间距: before=%s line=%s rule=%s' % (before, line, rule))
                            if before == '240':
                                score += 1; log.append('  OK 段前12磅 (+1)')
                            if rule == 'atLeast':
                                score += 1; log.append('  OK 最小值行距 (+1)')
                    break

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 5, 'name': '标题二格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ==================================================================
    # req6: 特定段落格式 ("二、"开头段落)
    # 标准答案: P11 "二、强化职业教育类型特色" 使用style=1(heading 1)
    # ==================================================================
    def req6(self):
        log = []; score = 0
        paras = list(self.doc_root.iter(W + 'p'))
        found = False
        for p in paras:
            txt = get_text(p)
            if txt.startswith('二、'):
                found = True
                log.append('  找到段落: ' + txt[:30])
                pPr = p.find(W + 'pPr')
                if pPr is not None:
                    pStyle = pPr.find(W + 'pStyle')
                    if pStyle is not None and pStyle.get(W + 'val', '') in ('1', '2'):
                        score = 10  # 标准答案使用标题样式即满分
                        log.append('  OK 使用标题样式%s (+10)' % pStyle.get(W + 'val', ''))
                break

        if not found:
            log.append('  NG 未找到"二、"开头段落')
            # 容错
            for p in paras:
                txt = get_text(p)
                pPr = p.find(W + 'pPr')
                if pPr is not None:
                    ps = pPr.find(W + 'pStyle')
                    if ps is not None and ps.get(W + 'val', '') in ('1', '2') and '二' in txt:
                        score = 8; log.append('  OK 找到样式段落 (+8)')
                        break

        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 6, 'name': '特定段落格式', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ==================================================================
    # req7: 图片设置
    # 标准答案: cx=1980000(5.5cm) cy=1800000(5.0cm), square环绕, 水平居中
    # ==================================================================
    def req7(self):
        log = []; score = 0
        # 搜索wp:extent
        matches = re.findall(r'cx="(\d+)"\s*cy="(\d+)"', self.doc_xml)
        img_found = False
        for cx_str, cy_str in matches:
            cx = int(cx_str) / 914400 * 2.54
            cy = int(cy_str) / 914400 * 2.54
            # 标准答案尺寸: 5.5cm x 5.0cm (cx=1980000, cy=1800000)
            if abs(cx - 5.5) < 0.2 and abs(cy - 5.0) < 0.2:
                log.append('  找到图片: %.2fcm x %.2fcm' % (cx, cy))
                img_found = True
                score += 6; log.append('  OK 尺寸5.5x5.0cm (+6)')
                break
            # 容错：如果尺寸是5.0x5.5（宽高互换）
            elif abs(cx - 5.0) < 0.2 and abs(cy - 5.5) < 0.2:
                log.append('  找到图片(宽高互换): %.2fcm x %.2fcm' % (cx, cy))
                img_found = True
                score += 4; log.append('  OK 尺寸接近 (+4)')
                break

        if not img_found:
            log.append('  NG 未找到图片(5.5x5.0cm)')
            return {'req': 7, 'name': '图片设置', 'score': 0, 'full': 10, 'detail': log, 'status': '未通过'}

        # 环绕方式: square (四周型)
        if 'wrapSquare' in self.doc_xml:
            score += 4; log.append('  OK 四周型环绕 (+4)')
        elif 'wrapTight' in self.doc_xml:
            score += 2; log.append('  OK 紧密型环绕 (+2)')

        # 水平居中: positionH + align=center
        if 'positionH' in self.doc_xml and 'center' in self.doc_xml:
            score += 2; log.append('  OK 水平居中 (+2)')

        if score >= 8: score = 10
        st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
        return {'req': 7, 'name': '图片设置', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ==================================================================
    # req8: 图片题注（标准答案以脚注形式实现）
    # 标准答案: 图片段落后有footnoteReference，脚注内容为"政策红利助推职业教育高质量发展"
    # 脚注格式: 楷体, 不加粗(w:b w:val="0")
    # ==================================================================
    def req8(self):
        log = []; score = 0
        
        # 1. 检查图片是否存在
        has_image = (self.doc_root.find('.//' + WP + 'extent') is not None)
        if not has_image:
            has_image = ('cx="' in self.doc_xml and 'cy="' in self.doc_xml)
        if not has_image:
            log.append('  NG 未找到图片')
            return {'req': 8, 'name': '图片题注', 'score': 0, 'full': 5, 'detail': log, 'status': '未通过'}

        score += 1; log.append('  OK 图片已插入 (+1)')

        # 2. 找图片所在段落中的脚注引用
        paras = list(self.doc_root.iter(W + 'p'))
        img_para = None
        fn_ref = None
        for p in paras:
            if p.find('.//' + WP + 'anchor') is not None or p.find('.//' + WP + 'extent') is not None:
                img_para = p
                fn_ref = p.find('.//' + W + 'footnoteReference')
                break

        # 3. 加载脚注内容
        footnotes_root = None
        fn_path = self.xml_dir / 'word' / 'footnotes.xml'
        if fn_path.exists():
            footnotes_root = ET.parse(str(fn_path)).getroot()

        if footnotes_root is not None:
            for fn in footnotes_root.findall(W + 'footnote'):
                fn_id = fn.get(W + 'id', '')
                if fn_id in ('-1', '0'):  # 跳过分隔符
                    continue
                
                txt = ''.join(t.text or '' for t in fn.findall('.//' + W + 't'))
                rPr = fn.find('.//' + W + 'rPr')
                
                # 检查题注关键词
                keywords = ['政策', '红利', '职业教', '发展', '图']
                matched = [kw for kw in keywords if kw in txt]
                
                if matched:
                    log.append('  脚注%d: %s' % (int(fn_id), txt[:40]))
                    score += 2; log.append('  OK 找到题注脚注 (+2)')
                    
                    # 检查字体（标准答案：楷体）
                    if rPr is not None:
                        rf = rPr.find(W + 'rFonts')
                        if rf is not None:
                            fonts = [rf.get(W + 'eastAsia', ''), rf.get(W + 'ascii', ''),
                                     rf.get(W + 'hAnsi', '')]
                            all_fonts = ' '.join(fonts)
                            if '楷体' in all_fonts or '宋体' in all_fonts:
                                score += 1; log.append('  OK 题注字体 (+1)')
                            else:
                                log.append('  字体: %s' % all_fonts)
                        
                        # 检查不加粗（w:b w:val="0" 或 无b元素）
                        b_elem = rPr.find(W + 'b')
                        if b_elem is None or b_elem.get(W + 'val', '') == '0':
                            score += 1; log.append('  OK 不加粗 (+1)')
                        elif b_elem.get(W + 'val', '') in ('1', 'true', 'on') or b_elem.attrib.get(W + 'val') is None:
                            log.append('  题注加粗')
                    else:
                        score += 2; log.append('  OK 题注存在 (+2)')
                    break
        
        if score < 3:
            log.append('  NG 未找到题注脚注')

        if score >= 4: score = 5
        st = '通过' if score >= 4 else ('部分通过' if score >= 2 else '未通过')
        return {'req': 8, 'name': '图片题注', 'score': min(score, 5), 'full': 5, 'detail': log, 'status': st}

    # ==================================================================
    # req9: 文字水印
    # 标准答案: "新华社", 华文新魏, font-size:1pt (VML textpath)
    # 注意: XML序列化后命名空间可能展开，需同时搜索原始形式和展开形式
    # ==================================================================
    def req9(self):
        log = []; score = 0
        # 搜索VML水印文本（header已加入doc_xml）
        # 序列化后可能是 v:textpath 或 {urn:...}textpath
        has_watermark = ('新华社' in self.doc_xml and ('textpath' in self.doc_xml))
        if not has_watermark:
            # 也检查原始header文件中的元素
            for nm, root in self.headers.items():
                shape = root.find('.//{' + V + '}shape')
                if shape is not None:
                    textpath = shape.find('{' + V + '}textpath')
                    if textpath is not None:
                        s = textpath.get('string', '')
                        if '新华社' in s or '新职业' in s:
                            has_watermark = True
                            log.append('  水印文字(%s): %s' % (nm, s))
                            break

        if has_watermark:
            score += 2; log.append('  OK 水印文字 (+2)')
        else:
            log.append('  NG 未检测到文字水印')
            return {'req': 9, 'name': '文字水印', 'score': 0, 'full': 5, 'detail': log, 'status': '未通过'}

        # 字体: 华文新魏
        if '华文新魏' in self.doc_xml:
            score += 1; log.append('  OK 华文新魏 (+1)')
        # 字号 (font-size在VML中)
        if 'font-size' in self.doc_xml:
            score += 1; log.append('  OK VML字号设置 (+1)')
        # 斜式 (rotation)
        if 'rotation' in self.doc_xml.lower():
            score += 1; log.append('  OK 斜式设置 (+1)')

        if score >= 4: score = 5
        st = '通过' if score >= 4 else ('部分通过' if score >= 2 else '未通过')
        return {'req': 9, 'name': '文字水印', 'score': min(score, 5), 'full': 5, 'detail': log, 'status': st}

    # ==================================================================
    # req10: 页眉页脚
    # 标准答案: 页眉(header2/偶数页)"新华社", 页脚(footer2)TIME域"yyyy/M/d", center对齐
    # ==================================================================
    def req10(self):
        log = []; score = 0

        # 检查页眉内容
        header_texts = []
        for nm, root in self.headers.items():
            txt = get_text(root)
            if txt.strip():
                header_texts.append((nm, txt.strip()))
                log.append('  %s: "%s"' % (nm, txt.strip()[:30]))

        header_ok = False
        for nm, txt in header_texts:
            if '新华社' in txt:
                score += 4; header_ok = True
                log.append('  OK 页眉"新华社" (+4)')
                # 检查字体和居中
                root = self.headers.get(nm)
                if root is not None:
                    for p in root.findall('.//' + W + 'p'):
                        pPr = p.find(W + 'pPr')
                        if pPr is not None:
                            jc = pPr.find(W + 'jc')
                            if jc is not None and jc.get(W + 'val', '') in ('center', 'right'):
                                score += 1; log.append('  OK 页眉对齐 (+1)')
                            break
                break

        # 检查页脚
        footer_ok = False
        for nm, root in self.footers.items():
            txt = get_text(root)
            xs = ET.tostring(root, encoding='unicode')
            if 'TIME' in xs or 'DATE' in xs:
                footer_ok = True
                log.append('  %s: TIME/DATE域' % nm)
                score += 3; log.append('  OK 页脚含时间域 (+3)')
                # 检查日期格式 yyyy/M/d
                if 'yyyy/M/d' in xs:
                    score += 1; log.append('  OK 日期格式yyyy/M/d (+1)')
                # 检查对齐
                for p in root.findall('.//' + W + 'p'):
                    pPr = p.find(W + 'pPr')
                    if pPr is not None:
                        jc = pPr.find(W + 'jc')
                        if jc is not None:
                            jc_val = jc.get(W + 'val', '')
                            log.append('  页脚对齐: %s' % jc_val)
                            if jc_val == 'center':
                                score += 1; log.append('  OK 页脚居中 (+1)')
                            elif jc_val == 'right':
                                score += 1; log.append('  OK 页脚右对齐 (+1)')
                        break
                break

        if not footer_ok:
            # 检查数字（可能是硬编码的日期）
            for nm, root in self.footers.items():
                txt = get_text(root)
                if re.search(r'\d{4}/\d{1,2}/\d{1,2}', txt):
                    score += 2; log.append('  OK 页脚含日期 (+2)')
                    break

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 10, 'name': '页眉页脚', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ==================================================================
    # req11: 页面颜色和边框
    # 标准答案: 背景F2F2F2, 双实线sz=6(0.75pt), color=0070C0(蓝色)
    # ==================================================================
    def req11(self):
        log = []; score = 0

        # 检查页面背景色 (background在document.xml中)
        bg_match = re.search(r'w:background[^>]*w:color="([^"]*)"', self.doc_xml)
        if not bg_match:
            bg_match = re.search(r'(' + W.replace('{','\\{').replace('}','\\}') + r'background)[^>]*' + W.replace('{','\\{').replace('}','\\}') + r'color"="([^"]*)"', self.doc_xml)
        if not bg_match:
            # 直接从document_root查找
            bg = self.doc_root.find(W + 'background')
            if bg is not None:
                c = bg.get(W + 'color', '')
                if c:
                    bg_match = True
                    log.append('  背景颜色: %s' % c)

        if bg_match or (self.doc_root.find(W + 'background') is not None):
            score += 2; log.append('  OK 页面背景色 (+2)')

        # 检查页面边框 - 直接从doc_root查找
        has_border = False
        for sp in self.doc_root.iter(W + 'sectPr'):
            pgBorders = sp.find(W + 'pgBorders')
            if pgBorders is not None:
                log.append('  找到pgBorders')
                border_xml = ET.tostring(pgBorders, encoding='unicode')
                
                # 检查双实线
                if 'double' in border_xml:
                    score += 3; log.append('  OK 双实线边框 (+3)')
                    has_border = True
                
                # 检查0.75pt (sz=6)
                sz_matches = re.findall(r'sz="(\d+)"', border_xml)
                if sz_matches and any(s == '6' for s in sz_matches):
                    score += 2; log.append('  OK 0.75磅 (+2)')
                
                # 检查蓝色 (0070C0)
                if '0070C0' in border_xml:
                    score += 3; log.append('  OK 蓝色 (+3)')
                    has_border = True
                    
                log.append('  边框: ' + border_xml[:150])
            break

        if not has_border:
            log.append('  NG 未找到页面边框或格式不匹配')

        if score >= 8: score = 10
        st = '通过' if score >= 7 else ('部分通过' if score >= 3 else '未通过')
        return {'req': 11, 'name': '页面颜色和边框', 'score': min(score, 10), 'full': 10, 'detail': log, 'status': st}

    # ==================================================================
    # 评分主流程
    # ==================================================================
    def grade(self):
        print('\n' + '=' * 60)
        print('  Word 2010 XML底层精准评分器 - 第5套题 v1')
        print('=' * 60)
        self.unzip()
        print('\n加载XML...')
        self.load_all()
        print('\n' + '-' * 60)
        print('  逐项检测...')
        print('-' * 60)

        checks = [self.req1, self.req2, self.req3, self.req4, self.req5,
                  self.req6, self.req7, self.req8, self.req9, self.req10, self.req11]
        names = ['页面设置', '装订线+文档网格+正文格式', '标题格式', '标题一格式',
                 '标题二格式', '特定段落格式', '图片设置', '图片题注',
                 '文字水印', '页眉页脚', '页面颜色和边框']

        for i, chk in enumerate(checks):
            r = chk()
            self.results.append(r)
            icon = {'通过': '[OK]', '部分通过': '[--]', '未通过': '[NG]'}.get(r['status'], '?')
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
        if path is None: path = self.docx.parent / 'report_5.html'
        total = sum(r['score'] for r in self.results)
        full = sum(r['full'] for r in self.results)
        rate = total / full * 100 if full > 0 else 0
        if rate >= 90: gr, gc = 'A 优秀', '#28a745'
        elif rate >= 75: gr, gc = 'B 良好', '#17a2b8'
        elif rate >= 60: gr, gc = 'C 及格', '#ffc107'
        else: gr, gc = 'D 不及格', '#dc3545'

        h = ('<!DOCTYPE html><html><head><meta charset="UTF-8">'
             '<title>Word评分报告 - 第5套题</title>'
             '<style>*{margin:0;padding:0;box-sizing:border-box;}'
             'body{font-family:Microsoft YaHei,sans-serif;background:linear-gradient(135deg,#667eea,#764ba2);min-height:100vh;padding:20px;}'
             '.box{max-width:1000px;margin:0 auto;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden;}'
             '.hd{background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:30px;text-align:center;}'
             '.hd h1{font-size:26px;margin-bottom:10px;}'
             '.sc{display:flex;align-items:center;justify-content:center;padding:30px;gap:50px;background:#f8f9fa;}'
             '.num{font-size:56px;font-weight:bold;color:' + gc + ';}'
             '.gb{background:' + gc + ';color:white;padding:8px 20px;border-radius:20px;font-weight:bold;display:inline-block;}'
             '.c{background:#f8f9fa;border-radius:12px;padding:16px;margin:12px 30px;border-left:5px solid #dee2e6;}'
             '.c.p1{border-left-color:#28a745;}.c.p2{border-left-color:#ffc107;}.c.p3{border-left-color:#dc3545;}'
             '.ch{display:flex;justify-content:space-between;margin-bottom:8px;}'
             '.ct{font-size:15px;font-weight:bold;}'
             '.bd{padding:3px 10px;border-radius:15px;font-size:12px;font-weight:bold;}'
             '.b1{background:#d4edda;color:#155724;}.b2{background:#fff3cd;color:#856404;}.b3{background:#f8d7da;color:#721c24;}'
             '.d{background:white;padding:10px;border-radius:8px;margin-top:8px;font-size:12px;line-height:1.8;max-height:200px;overflow-y:auto;}'
             '.ft{text-align:center;padding:15px;background:#f8f9fa;color:#6c757d;font-size:12px;}'
             '</style></head><body>'
             '<div class="box">'
             '<div class="hd"><h1>Word 2010 操作题评分报告（第5套）</h1>'
             '<p>基于XML底层解析 | 共11项评分要求 | v1</p></div>'
             '<div class="sc"><div style="text-align:center;">'
             '<div class="num">' + str(total) + '</div>'
             '<div style="color:#6c757d;font-size:16px;">/ ' + str(full) + ' 分</div>'
             '<div style="font-size:40px;color:' + gc + ';">' + ('%.1f' % rate) + '%</div>'
             '<div class="gb">' + gr + '</div>'
             '</div></div>')

        for r in self.results:
            sc = {'通过': 'p1', '部分通过': 'p2', '未通过': 'p3'}.get(r['status'], 'p3')
            bd = {'通过': 'b1', '部分通过': 'b2', '未通过': 'b3'}.get(r['status'], 'b3')
            h += '<div class="c ' + sc + '"><div class="ch">'
            h += '<span class="ct">要求%02d %s (%d/%d)</span>' % (r['req'], r['name'], r['score'], r['full'])
            h += '<span class="bd ' + bd + '">' + r['status'] + '</span>'
            h += '</div><div class="d">'
            for d in r['detail']:
                h += '<div>' + d + '</div>'
            h += '</div></div>'

        h += ('<div class="ft"><p>基于XML底层解析 | 精准对照11项要求</p>'
             '<p>生成: ' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + '</p></div>'
             '</div></body></html>')

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
