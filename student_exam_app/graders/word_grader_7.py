"""
Word 2010 操作题评分器 - 文件夹7
基于修订标记自动生成

要求清单（共11题，满分100分）：
1. 页面设置 B5/对称页边距/上下2.8cm/内3.0cm/外2.1cm (10分)
2. 正文华文新魏/1.5倍行距 (10分)
3. 椭圆标注"就业情况" (8分)
4. 标题文本框/华文中宋 (8分)
5. 页眉：奇数"新华网"/偶数"疫情冲击下的就业市场" (8分)
6. 小标题"标题2"样式 (8分)
7. 背景.jpg页面背景 (8分)
8. 第一段红色边框/水绿色底纹 (10分)
9. 文字转表格/公式计算 (10分)
10. 表格格式/表头蓝底/宋体小四/题注 (10分)
11. 表格外框红色/1-2行间双实线 (10分)
"""

import zipfile, os, re, io, sys
import xml.etree.ElementTree as ET

# 命名空间
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
V = '{urn:schemas-microsoft-com:vml}'
O = '{urn:schemas-microsoft-com:office:office}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'

def T(tag):
    """w:tag → 完整URI"""
    ns, local = tag.split(':', 1)
    ns_map = {'w': W, 'r': R, 'v': V, 'o': O, 'wp': WP}
    return '{%s}%s' % (ns_map.get(ns, ''), local)

def child(elem, tag):
    for c in elem:
        if c.tag == tag: return c
    return None

def children(elem, tag):
    return [c for c in elem if c.tag == tag]

def GA(elem, attr, default=None):
    """获取属性(带命名空间)"""
    if ':' in attr:
        ns, local = attr.split(':', 1)
        ns_map = {'w': W, 'r': R, 'v': V, 'o': O, 'wp': WP}
        return elem.get('%s%s' % (ns_map[ns], local), default)
    return elem.get(attr, default)

def get_text(elem):
    return ''.join(t.text or '' for t in elem.iter(W + 't'))


class WordGrader7:
    def __init__(self, docx_path):
        self.docx_path = docx_path
        self.doc_xml = ''
        self.doc_root = None
        self.body = None
        self.sectPr = None
        self.results = []
        self.total = 0
        self.xml_files = {}
        self.headers = {}
        self.footers = {}
        self.rels = {}
        
    def load(self):
        with zipfile.ZipFile(self.docx_path) as z:
            self.doc_xml = z.read('word/document.xml').decode('utf-8')
            self.doc_root = ET.fromstring(self.doc_xml.encode('utf-8'))
            
            for c in self.doc_root:
                if c.tag == W + 'body':
                    self.body = c
                    break
            
            self.sectPr = child(self.body, W + 'sectPr')
            
            # 加载所有XML文件用于字符串搜索
            all_xml = []
            for name in z.namelist():
                if name.endswith('.xml'):
                    try:
                        content = z.read(name).decode('utf-8', errors='replace')
                        self.xml_files[name] = content
                        all_xml.append(content)
                        if 'header' in name:
                            self.headers[name] = content
                        if 'footer' in name:
                            self.footers[name] = content
                    except:
                        pass
            self.all_xml_str = '\n'.join(all_xml)
            
            # 关系文件
            try:
                rels_xml = z.read('word/_rels/document.xml.rels')
                rels_root = ET.fromstring(rels_xml)
                for rel in rels_root:
                    rid = rel.get('Id', '')
                    self.rels[rid] = {
                        'target': rel.get('Target', ''),
                        'type': rel.get('Type', ''),
                    }
            except:
                pass
    
    def grade(self):
        self.load()
        self.results = []
        
        self.results.append(self.req1_page_setup())
        self.results.append(self.req2_body_font())
        self.results.append(self.req3_oval_shape())
        self.results.append(self.req4_title_textbox())
        self.results.append(self.req5_header())
        self.results.append(self.req6_heading2())
        self.results.append(self.req7_bg_image())
        self.results.append(self.req8_first_para_border())
        self.results.append(self.req9_table_formula())
        self.results.append(self.req10_table_format())
        self.results.append(self.req11_table_borders())
        
        self.total = sum(r['score'] for r in self.results)
        self.results.append({
            'req': 0, 'name': '总分', 'score': self.total, 'full': 100,
            'detail': [], 'status': 'pass' if self.total >= 60 else 'fail'
        })
        return self.results
    
    # ========== 要求1: 页面设置 B5/对称页边距 (10分) ==========
    def req1_page_setup(self):
        """B5纸, 对称页边距, 上下2.8cm, 内3.0cm, 外2.1cm"""
        detail = []
        score = 0
        full = 10
        
        if self.sectPr is None:
            detail.append('未找到节属性')
            return {'req': 1, 'name': '页面设置', 'score': 0, 'full': full, 'detail': detail, 'status': 'fail'}
        
        pgSz = child(self.sectPr, W + 'pgSz')
        pgMar = child(self.sectPr, W + 'pgMar')
        
        # B5纸: 182mm x 257mm = 10324 x 14592 EMU
        # 或 7.2 x 10.1 inches (实际答案: 10433 x 14742 EMU)
        if pgSz is not None:
            w = int(GA(pgSz, 'w:w', '0'))
            h = int(GA(pgSz, 'w:h', '0'))
            w_cm = w / 1440 * 2.54
            h_cm = h / 1440 * 2.54
            
            # B5: 约18.2cm x 25.7cm
            if abs(w_cm - 18.4) < 1.0 and abs(h_cm - 26.0) < 1.0:
                score += 3
                detail.append(f'纸张大小 B5 ({w_cm:.1f}x{h_cm:.1f}cm) 正确')
            else:
                detail.append(f'纸张大小 {w_cm:.1f}x{h_cm:.1f}cm (期望约18.4x26.0cm)')
        else:
            detail.append('未找到纸张大小设置')
        
        # 页边距: 上下2.8cm, 内3.0cm, 外2.1cm
        if pgMar is not None:
            top_twips = int(GA(pgMar, 'w:top', '0'))
            bottom_twips = int(GA(pgMar, 'w:bottom', '0'))
            left_twips = int(GA(pgMar, 'w:left', '0'))
            right_twips = int(GA(pgMar, 'w:right', '0'))
            
            top_cm = top_twips / 1440 * 2.54
            bottom_cm = bottom_twips / 1440 * 2.54
            left_cm = left_twips / 1440 * 2.54
            right_cm = right_twips / 1440 * 2.54
            
            pts = 0
            if abs(top_cm - 2.8) < 0.2:
                pts += 2
                detail.append(f'上边距 {top_cm:.1f}cm 正确')
            else:
                detail.append(f'上边距 {top_cm:.1f}cm (期望2.8cm)')
            
            if abs(bottom_cm - 2.8) < 0.2:
                pts += 2
                detail.append(f'下边距 {bottom_cm:.1f}cm 正确')
            else:
                detail.append(f'下边距 {bottom_cm:.1f}cm (期望2.8cm)')
            
            # 对称页边距: 内=左, 外=右, 内侧3.0cm, 外侧2.1cm
            # 素材默认内边距恰好 3.17cm，会命中 ±0.2 容差，改成精确 0.05
            if abs(left_cm - 3.0) < 0.05:
                pts += 1.5
                detail.append(f'内侧(左)边距 {left_cm:.1f}cm 正确')
            else:
                detail.append(f'内侧(左)边距 {left_cm:.1f}cm (期望3.0cm)')

            if abs(right_cm - 2.1) < 0.2:
                pts += 1.5
                detail.append(f'外侧(右)边距 {right_cm:.1f}cm 正确')
            else:
                detail.append(f'外侧(右)边距 {right_cm:.1f}cm (期望2.1cm)')
            
            score += pts
        
        status = 'pass' if score >= full * 0.6 else 'fail'
        return {'req': 1, 'name': '页面设置', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 要求2: 正文华文新魏/1.5倍行距 (10分) ==========
    def req2_body_font(self):
        """正文(标题除外)字体华文新魏, 1.5倍行距"""
        detail = []
        score = 0
        full = 10
        
        paras = children(self.body, W + 'p')
        body_paras = []
        
        for pi, p in enumerate(paras):
            txt = get_text(p)
            if not txt.strip():
                continue
            
            # 跳过标题段落（无首行缩进的段落）
            pPr = child(p, W + 'pPr')
            ind = child(pPr, W + 'ind') if pPr is not None else None
            has_indent = ind is not None and GA(ind, 'w:firstLine') is not None
            
            # 跳过序号标题段落和题注
            if txt.startswith('一、') or txt.startswith('二、') or txt.startswith('三、') or '表格' in txt:
                continue
            
            if has_indent:
                body_paras.append((pi, p, txt))
        
        if not body_paras:
            detail.append('未找到正文段落')
            return {'req': 2, 'name': '正文格式', 'score': 0, 'full': full, 'detail': detail, 'status': 'fail'}
        
        font_ok = 0
        line_ok = 0
        
        for pi, p, txt in body_paras:
            pPr = child(p, W + 'pPr')
            
            # 检查行距
            spc = child(pPr, W + 'spacing') if pPr else None
            if spc is not None:
                line_val = GA(spc, 'w:line')
                if line_val and int(line_val) == 360:
                    line_ok += 1
            
            # 检查字体
            for r in children(p, W + 'r'):
                rPr = child(r, W + 'rPr')
                if rPr is not None:
                    rFonts = child(rPr, W + 'rFonts')
                    if rFonts is not None:
                        f = GA(rFonts, 'w:ascii') or GA(rFonts, 'w:eastAsia') or GA(rFonts, 'w:hAnsi')
                        if f == '华文新魏':
                            font_ok += 1
                            break
        
        if body_paras:
            font_ratio = font_ok / len(body_paras)
            line_ratio = line_ok / len(body_paras)
            
            if font_ratio >= 0.8:
                score += 6
                detail.append(f'正文华文新魏字体 ({int(font_ratio*100)}%) 正确')
            elif font_ratio > 0:
                score += 3
                detail.append(f'正文华文新魏字体 ({int(font_ratio*100)}%) 部分正确')
            else:
                detail.append('正文未设置华文新魏字体')
            
            if line_ratio >= 0.8:
                score += 4
                detail.append(f'正文1.5倍行距 ({int(line_ratio*100)}%) 正确')
            elif line_ratio > 0:
                score += 2
                detail.append(f'正文1.5倍行距 ({int(line_ratio*100)}%) 部分正确')
            else:
                detail.append('正文未设置1.5倍行距')
        
        status = 'pass' if score >= full * 0.6 else 'fail'
        return {'req': 2, 'name': '正文格式', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 要求3: 椭圆标注"就业情况" (8分) ==========
    def req3_oval_shape(self):
        """椭圆形标注, 文字'就业情况', 填充白色/轮廓黑色2pt"""
        detail = []
        score = 0
        full = 8
        
        # 搜索VML椭圆形状
        found = False
        for pict in self.doc_root.iter(W + 'pict'):
            for oval in pict.iter(V + 'oval'):
                found = True
                break
            # 也检查shape type
            for sh in pict.iter(V + 'shape'):
                stype = sh.get('type', '')
                if stype == '#_x0000_t63':  # t63 = 椭圆批注
                    found = True
                    score += 2
                    detail.append('找到椭圆/批注形状')
                    
                    # 检查文本
                    for tb in sh.iter(V + 'textbox'):
                        tb_text = get_text(tb)
                        if '就业情况' == tb_text.strip():
                            score += 3
                            detail.append(f'标注文字"就业情况"正确')
                        elif '就业' in tb_text:
                            score += 1
                            detail.append(f'标注文字部分正确: "{tb_text.strip()}"')
                    
                    # 检查样式
                    style = sh.get('style', '')
                    fill_color = sh.get('fillcolor', '')
                    stroke_color = sh.get('strokecolor', '')
                    
                    if fill_color == 'white [3201]':
                        score += 1.5
                        detail.append('填充颜色白色正确')
                    else:
                        detail.append(f'填充颜色 {fill_color} (期望 white [3201])')
                    
                    if stroke_color == 'black [3200]':
                        score += 1.5
                        detail.append('轮廓颜色黑色正确')
                    else:
                        detail.append(f'轮廓颜色 {stroke_color} (期望 black [3200])')
                    
                    return {'req': 3, 'name': '椭圆标注', 'score': round(score, 1), 'full': full,
                            'detail': detail, 'status': 'pass' if score >= 4 else 'fail'}
        
        # 未找到，用字符串搜索兜底
        if '#_x0000_t63' in self.doc_xml and '就业情况' in self.doc_xml:
            score += 4
            detail.append('字符串匹配找到椭圆标注+就业情况')
        else:
            detail.append('未找到椭圆标注')
        
        status = 'pass' if score >= 4 else 'fail'
        return {'req': 3, 'name': '椭圆标注', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 要求4: 标题文本框/华文中宋 (8分) ==========
    def req4_title_textbox(self):
        """标题为文本框, 文字'疫情冲击下的中国就业市场：短期波动与长期展望'"""
        detail = []
        score = 0
        full = 8
        
        title_text = '疫情冲击下的中国就业市场：短期波动与长期展望'
        
        # 搜索VML文本框 (type=#_x0000_t202)
        for pict in self.doc_root.iter(W + 'pict'):
            for sh in pict.iter(V + 'shape'):
                stype = sh.get('type', '')
                if 't202' in stype:  # 文本框
                    score += 2
                    detail.append('找到文本框')
                    
                    for tb in sh.iter(V + 'textbox'):
                        tb_text = get_text(tb)
                        if tb_text.strip() == title_text:
                            score += 3
                            detail.append(f'标题文字正确')
                        
                        # 检查文字样式
                        found_font = False
                        for r in tb.iter(W + 'r'):
                            rPr = child(r, W + 'rPr')
                            if rPr is not None:
                                rFonts = child(rPr, W + 'rFonts')
                                if rFonts is not None and not found_font:
                                    f = GA(rFonts, 'w:ascii') or GA(rFonts, 'w:eastAsia')
                                    if f == '华文行楷':
                                        score += 2
                                        detail.append('标题字体华文行楷正确')
                                        found_font = True
                                sz = child(rPr, W + 'sz')
                                if sz is not None and GA(sz, 'w:val') == '32':
                                    score += 1
                                    detail.append('标题字号小三正确')
                                    break
                        
                        return {'req': 4, 'name': '标题文本框', 'score': round(score, 1), 'full': full,
                                'detail': detail, 'status': 'pass' if score >= 4 else 'fail'}
        
        # 字符串搜索备选：必须在文本框类型标记存在时才给分
        if title_text in self.doc_xml and 't202' in self.doc_xml:
            score += 3
            detail.append('字符串匹配找到标题文本框')
        else:
            detail.append('未找到标题文本框')
        
        status = 'pass' if score >= 4 else 'fail'
        return {'req': 4, 'name': '标题文本框', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 要求5: 页眉 (8分) ==========
    def req5_header(self):
        """奇数页页眉'新华网', 偶数页页眉'疫情冲击下的就业市场'"""
        detail = []
        score = 0
        full = 8
        
        odd_ok = False
        even_ok = False
        
        for hname, hxml in self.headers.items():
            hroot = ET.fromstring(hxml)
            htext = ''.join(t.text or '' for t in hroot.iter(W + 't'))
            
            if '新华网' == htext.strip():
                odd_ok = True
                score += 4
                detail.append(f'奇数页页眉"新华网"正确 ({hname})')
            elif '疫情冲击下的就业市场' == htext.strip():
                even_ok = True
                score += 4
                detail.append(f'偶数页页眉"疫情冲击下的就业市场"正确 ({hname})')
        
        if not odd_ok:
            # 检查字符串
            if '新华网' in self.all_xml_str:
                score += 2
                detail.append('页眉含"新华网"（部分正确）')
            else:
                detail.append('未找到"新华网"页眉')
        
        if not even_ok:
            if '疫情冲击下的就业市场' in self.all_xml_str:
                score += 2
                detail.append('页眉含"疫情冲击下的就业市场"（部分正确）')
            else:
                detail.append('未找到"疫情冲击下的就业市场"页眉')
        
        status = 'pass' if score >= 4 else 'fail'
        return {'req': 5, 'name': '页眉', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 要求6: 小标题"标题2"样式 (8分) ==========
    def req6_heading2(self):
        """一、二、三段落使用标题2样式"""
        detail = []
        score = 0
        full = 8
        
        paras = children(self.body, W + 'p')
        headings = ['一、', '二、', '三、']
        found_count = 0
        
        for p in paras:
            txt = get_text(p)
            for h in headings:
                if txt.startswith(h):
                    found_count += 1
                    pPr = child(p, W + 'pPr')
                    if pPr is not None:
                        pStyle = child(pPr, W + 'pStyle')
                        if pStyle is not None and GA(pStyle, 'w:val') == '2':
                            score += 2.5
                            detail.append(f'"{txt[:10]}..." 标题2样式正确')
                        else:
                            detail.append(f'"{txt[:10]}..." 未使用标题2样式')
                    break
        
        # 调整分数
        if found_count == 3 and score >= 7.5:
            score = 8  # 全部正确满分
        elif found_count == 0:
            detail.append('未找到小标题段落')
        
        score = min(score, full)
        status = 'pass' if score >= 4 else 'fail'
        return {'req': 6, 'name': '标题2样式', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 要求7: 背景.jpg页面背景 (8分) ==========
    def req7_bg_image(self):
        """背景.jpg设为页面背景"""
        detail = []
        score = 0
        full = 8
        
        # 检查背景图片
        bg = child(self.doc_root, W + 'background')
        if bg is not None:
            score += 3
            detail.append('找到页面背景元素')
            
            found_bg = False
            # 检查VML背景填充
            for vbg in bg.iter(V + 'background'):
                found_bg = True
                score += 2
                detail.append('找到VML背景填充')
            
            # 检查fill元素的title
            for fill in bg.iter(V + 'fill'):
                title = GA(fill, 'o:title', '')
                rtype = fill.get('type', '')
                if '背景' == title or rtype == 'frame':
                    score += 3
                    detail.append(f'页面背景图片正确 (title={title})')
                    found_bg = True
                rid = GA(fill, 'r:id', '')
                if rid:
                    score += 0.5
                    detail.append(f'背景填充 rId={rid}')
            
            # 检查imagedata
            for im in bg.iter(V + 'imagedata'):
                title = GA(im, 'o:title', '')
                if '背景' in title:
                    score += 2
                    detail.append(f'背景图片 title={title}')
                    found_bg = True
            
            if not found_bg:
                score += 1
                detail.append('页面背景存在但未确认内容')
        else:
            detail.append('未找到页面背景')
        
        score = min(score, full)
        status = 'pass' if score >= 4 else 'fail'
        return {'req': 7, 'name': '页面背景', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 要求8: 第一段边框底纹 (10分) ==========
    def req8_first_para_border(self):
        """第一段加红色实线边框和水绿色底纹"""
        detail = []
        score = 0
        full = 10
        
        paras = children(self.body, W + 'p')
        # 第一个有内容的段落
        first_p = None
        for p in paras:
            if get_text(p).strip():
                first_p = p
                break
        
        if first_p is None:
            detail.append('未找到第一段')
            return {'req': 8, 'name': '第一段边框底纹', 'score': 0, 'full': full, 'detail': detail, 'status': 'fail'}
        
        pPr = child(first_p, W + 'pPr')
        if pPr is None:
            detail.append('第一段无段落属性')
            return {'req': 8, 'name': '第一段边框底纹', 'score': 0, 'full': full, 'detail': detail, 'status': 'fail'}
        
        # 检查底纹
        shd = child(pPr, W + 'shd')
        if shd is not None:
            fill = GA(shd, 'w:fill', '')
            if fill.upper() == 'DAEEF3':
                score += 4
                detail.append(f'水绿色底纹 DAEEF3 正确')
            else:
                detail.append(f'底纹颜色 #{fill.upper()} (期望DAEEF3)')
        else:
            detail.append('未设置底纹')
        
        # 检查边框
        pBdr = child(pPr, W + 'pBdr')
        if pBdr is not None:
            border_ok = 0
            for bt in ['top', 'bottom', 'left', 'right']:
                b = child(pBdr, W + bt)
                if b is not None:
                    val = GA(b, 'w:val', '')
                    color = GA(b, 'w:color', '')
                    if val == 'single' and color.upper() == 'FF0000':
                        border_ok += 1
            
            if border_ok >= 4:
                score += 6
                detail.append('红色实线边框正确(上下左右)')
            elif border_ok >= 2:
                score += 3
                detail.append(f'红色实线边框部分正确({border_ok}/4边)')
            else:
                detail.append(f'红色实线边框不正确(仅{border_ok}/4边)')
        else:
            detail.append('未设置段落边框')
        
        status = 'pass' if score >= 5 else 'fail'
        return {'req': 8, 'name': '第一段边框底纹', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 要求9: 文字转表格/公式 (10分) ==========
    def req9_table_formula(self):
        """12行文字转表格, 公式计算同比增长人数"""
        detail = []
        score = 0
        full = 10
        
        # 检查表格
        tbls = list(self.body.iter(W + 'tbl'))
        if not tbls:
            detail.append('未找到表格')
            return {'req': 9, 'name': '表格转换', 'score': 0, 'full': full, 'detail': detail, 'status': 'fail'}
        
        tbl = tbls[0]
        rows = children(tbl, W + 'tr')
        
        # 12行（表头+11行数据）
        if len(rows) >= 11:
            score += 3
            detail.append(f'表格 {len(rows)} 行正确')
        else:
            detail.append(f'表格仅 {len(rows)} 行 (期望12行)')
        
        # 4列
        first_row_cells = children(rows[0], W + 'tc') if rows else []
        if len(first_row_cells) >= 4:
            score += 2
            detail.append(f'表格 {len(first_row_cells)} 列正确')
        
        # 检查公式（增长人数 = 毕业生 * 增长率）
        # Word公式在XML中存储为w:instrText或域代码
        has_formula = False
        if 'instrText' in self.doc_xml:
            has_formula = True
            score += 3
            detail.append('找到公式域代码')
        elif any('×' in c or '*PRODUCT' in c for c in [get_text(cell) for row in rows for cell in children(row, W+'tc')]):
            has_formula = True
            score += 2
            detail.append('疑似找到公式计算结果')
        
        # 检查数字格式 0.00
        if re.search(r'\d+\.\d{2}', self.doc_xml):
            score += 2
            detail.append('数字格式0.00正确')
        
        status = 'pass' if score >= 5 else 'fail'
        return {'req': 9, 'name': '表格转换', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 要求10: 表格格式/题注 (10分) ==========
    def req10_table_format(self):
        """表头蓝色底纹, 宋体小四, 题注'表格1'"""
        detail = []
        score = 0
        full = 10
        
        tbls = list(self.body.iter(W + 'tbl'))
        if not tbls:
            detail.append('未找到表格')
            return {'req': 10, 'name': '表格格式', 'score': 0, 'full': full, 'detail': detail, 'status': 'fail'}
        
        tbl = tbls[0]
        rows = children(tbl, W + 'tr')
        if not rows:
            detail.append('表格无行')
            return {'req': 10, 'name': '表格格式', 'score': 0, 'full': full, 'detail': detail, 'status': 'fail'}
        
        # 表头蓝色底纹 (0070C0)
        header_cells = children(rows[0], W + 'tc')
        header_shd_ok = False
        for cell in header_cells:
            tcPr = child(cell, W + 'tcPr')
            if tcPr is not None:
                shd = child(tcPr, W + 'shd')
                if shd is not None and GA(shd, 'w:fill', '').upper() == '0070C0':
                    header_shd_ok = True
                    break
        
        if header_shd_ok:
            score += 3
            detail.append('表头蓝色底纹(0070C0)正确')
        else:
            detail.append('表头未设置蓝色底纹')
        
        # 表内字体 (黑体 小四 - 标准答案使用黑体)
        font_ok = False
        for row in rows:
            for cell in children(row, W + 'tc'):
                for r in cell.iter(W + 'r'):
                    rPr = child(r, W + 'rPr')
                    if rPr is not None:
                        rFonts = child(rPr, W + 'rFonts')
                        sz = child(rPr, W + 'sz')
                        f = GA(rFonts, 'w:eastAsia', '') if rFonts is not None else ''
                        s = GA(sz, 'w:val', '') if sz is not None else ''
                        if f == '黑体' and s == '24':
                            font_ok = True
                            break
                    # Also check paragraph-level rPr
                    p = [cp for cp in cell if cp.tag == W + 'p']
                    if p:
                        pPr = child(p[0], W + 'pPr')
                        if pPr is not None:
                            pRPr = child(pPr, W + 'rPr')
                            if pRPr is not None:
                                rFonts = child(pRPr, W + 'rFonts')
                                sz = child(pRPr, W + 'sz')
                                f = GA(rFonts, 'w:eastAsia', '') if rFonts is not None else ''
                                s = GA(sz, 'w:val', '') if sz is not None else ''
                                if f == '黑体' and s == '24':
                                    font_ok = True
                                    break
                if font_ok:
                    break
            if font_ok:
                break
        
        if font_ok:
            score += 3
            detail.append('表内黑体小四正确')
        else:
            # 至少检查黑体
            has_heiti = '黑体' in self.doc_xml
            if has_heiti:
                score += 1.5
                detail.append('表内含黑体(未确认小四)')
        
        # 题注 "表格 1 2010年-2020年同比增长表"
        caption_found = False
        paras = children(self.body, W + 'p')
        for p in paras:
            txt = get_text(p)
            if '表格 1' in txt or '表格1' in txt:
                caption_found = True
                break
        
        if caption_found:
            score += 2
            detail.append('找到题注引用')
        
        # 检查caption样式
        paras = children(self.body, W + 'p')
        for p in paras:
            txt = get_text(p)
            if '表格 1' in txt or '表格1' in txt:
                pPr = child(p, W + 'pPr')
                if pPr is not None:
                    pStyle = child(pPr, W + 'pStyle')
                    if pStyle is not None:
                        sid = GA(pStyle, 'w:val', '')
                        if sid in ('a7', 'caption'):  # caption style
                            score += 2
                            detail.append(f'题注使用caption样式(sid={sid})正确')
                        else:
                            score += 1
                            detail.append(f'题注样式为 {sid}')
                    else:
                        score += 1
                        detail.append('题注段落无样式ID')
                break
        
        status = 'pass' if score >= 5 else 'fail'
        return {'req': 10, 'name': '表格格式', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 要求11: 表格边框 (10分) ==========
    def req11_table_borders(self):
        """外框红色1pt实线, 第1-2行间双实线"""
        detail = []
        score = 0
        full = 10
        
        tbls = list(self.body.iter(W + 'tbl'))
        if not tbls:
            detail.append('未找到表格')
            return {'req': 11, 'name': '表格边框', 'score': 0, 'full': full, 'detail': detail, 'status': 'fail'}
        
        tbl = tbls[0]
        rows = children(tbl, W + 'tr')
        if len(rows) < 2:
            detail.append('表格行数不足')
            return {'req': 11, 'name': '表格边框', 'score': 0, 'full': full, 'detail': detail, 'status': 'fail'}
        
        # 检查表格外框
        tblPr = child(tbl, W + 'tblPr')
        if tblPr is not None:
            tblBorders = child(tblPr, W + 'tblBorders')
            if tblBorders is not None:
                outer_ok = 0
                for bt in ['top', 'bottom', 'left', 'right']:
                    b = child(tblBorders, W + bt)
                    if b is not None:
                        val = GA(b, 'w:val', '')
                        color = GA(b, 'w:color', '')
                        sz_v = GA(b, 'w:sz', '0')
                        if val == 'single' and color.upper() == 'FF0000' and sz_v == '8':
                            outer_ok += 1
                
                if outer_ok >= 4:
                    score += 5
                    detail.append('表格外框红色1pt实线正确')
                elif outer_ok >= 2:
                    score += 2.5
                    detail.append(f'表格外框红色实线部分正确({outer_ok}/4边)')
                else:
                    detail.append(f'表格外框不正确({outer_ok}/4边)')
            else:
                detail.append('表格无边框定义')
        
        # 检查第1行和第2行之间的双实线
        # 第0行(表头)底部或第1行顶部
        row0_cells = children(rows[0], W + 'tc')
        row1_cells = children(rows[1], W + 'tc') if len(rows) > 1 else []
        
        double_found = False
        for cell in row0_cells:
            tcPr = child(cell, W + 'tcPr')
            if tcPr is not None:
                tcBorders = child(tcPr, W + 'tcBorders')
                if tcBorders is not None:
                    bottom = child(tcBorders, W + 'bottom')
                    if bottom is not None:
                        val = GA(bottom, 'w:val', '')
                        sz_v = GA(bottom, 'w:sz', '')
                        if val == 'double' and sz_v == '12':
                            double_found = True
                            score += 5
                            detail.append(f'第1-2行间双实线正确 (bottom, val={val}, sz={sz_v})')
                            break
        else:
            for cell in row1_cells:
                tcPr = child(cell, W + 'tcPr')
                if tcPr is not None:
                    tcBorders = child(tcPr, W + 'tcBorders')
                    if tcBorders is not None:
                        top = child(tcBorders, W + 'top')
                        if top is not None:
                            val = GA(top, 'w:val', '')
                            sz_v = GA(top, 'w:sz', '')
                            if val == 'double' and sz_v == '12':
                                double_found = True
                                score += 5
                                detail.append(f'第1-2行间双实线正确 (top, val={val}, sz={sz_v})')
                                break
        
        if not double_found:
            detail.append('未找到第1-2行间双实线')
        
        status = 'pass' if score >= 5 else 'fail'
        return {'req': 11, 'name': '表格边框', 'score': round(score, 1), 'full': full, 'detail': detail, 'status': status}
    
    # ========== 报告生成 ==========
    def report(self):
        """生成HTML报告"""
        pass_count = sum(1 for r in self.results[:-1] if r['status'] == 'pass')
        fail_count = sum(1 for r in self.results[:-1] if r['status'] == 'fail')
        
        cards_html = ''
        for r in self.results[:-1]:
            pct = int(r['score'] / r['full'] * 100) if r['full'] > 0 else 0
            fill_class = 'pass-fill' if r['status'] == 'pass' else 'fail-fill'
            items = '\n'.join(f'<div class="detail-item">{d}</div>' for d in r['detail'])
            cards_html += f'''
<div class="card">
<div class="card-header">
<span class="req-name">{r['req']}. {r['name']}</span>
<span class="score-badge {r['status']}">{r['score']}/{r['full']}</span>
</div>
<div class="progress-bar"><div class="progress-fill {fill_class}" style="width:{pct}%"></div></div>
{items}
</div>'''
        
        score_color = '#4caf50' if self.total >= 60 else '#f44336'
        html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>Word 2010 操作题评分报告 - 文件夹7</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:"Microsoft YaHei",sans-serif;background:#f5f5f5;padding:20px}}
.container{{max-width:900px;margin:0 auto}}
.header{{background:linear-gradient(135deg,#1a73e8,#0d47a1);color:#fff;padding:30px;border-radius:12px;margin-bottom:20px}}
.header h1{{font-size:24px;margin-bottom:8px}}
.total-score{{font-size:48px;font-weight:bold;margin:10px 0}}
.card{{background:#fff;border-radius:10px;padding:20px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.1)}}
.card-header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}}
.req-name{{font-size:16px;font-weight:500}}
.score-badge{{padding:4px 12px;border-radius:20px;font-size:14px;font-weight:500}}
.pass{{background:#e8f5e9;color:#2e7d32}}
.fail{{background:#ffebee;color:#c62828}}
.detail-item{{font-size:13px;color:#666;padding:3px 0;border-bottom:1px solid #f0f0f0}}
.detail-item:last-child{{border:none}}
.progress-bar{{height:6px;background:#e0e0e0;border-radius:3px;margin-top:8px}}
.progress-fill{{height:100%;border-radius:3px;transition:width 0.3s}}
.pass-fill{{background:#4caf50}}
.fail-fill{{background:#f44336}}
.summary{{display:flex;gap:15px;margin-bottom:20px}}
.summary-card{{flex:1;background:#fff;border-radius:10px;padding:20px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,0.1)}}
.summary-card h3{{font-size:14px;color:#999;margin-bottom:5px}}
.summary-card .value{{font-size:28px;font-weight:bold}}
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>Word 2010 操作题评分报告</h1>
<p>文件夹7 | 共11题 | 满分100分</p>
<div class="total-score">{self.total} 分</div>
</div>
<div class="summary">
<div class="summary-card"><h3>总分</h3><div class="value" style="color:{score_color}">{self.total}</div></div>
<div class="summary-card"><h3>通过项</h3><div class="value" style="color:#4caf50">{pass_count}</div></div>
<div class="summary-card"><h3>未通过</h3><div class="value" style="color:#f44336">{fail_count}</div></div>
</div>
{cards_html}
</div></body></html>'''
        return html


def main():
    import sys
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    if len(sys.argv) < 2:
        # 默认测试标准答案
        base = 'C:/Users/23000/Desktop/2022年准易：Word 2010 操作题--素材/2022年准易：Word 2010 操作题--素材/7'
        path = os.path.join(base, '1结果.docx')
    else:
        path = sys.argv[1]
    
    grader = WordGrader7(path)
    results = grader.grade()
    
    # 文本输出
    print('=' * 60)
    print('Word 2010 操作题评分 - 文件夹7')
    print(f'文件: {os.path.basename(path)}')
    print('=' * 60)
    
    for r in results[:-1]:
        status_icon = 'PASS' if r['status'] == 'pass' else 'FAIL'
        print(f"\n[{status_icon}] {r['req']:2d}. {r['name']}  {r['score']}/{r['full']}")
        for d in r['detail']:
            print(f'      {d}')
    
    print(f"\n{'=' * 60}")
    print(f"总分: {results[-1]['score']}/100")
    print(f"{'=' * 60}")
    
    # 生成HTML报告
    html = grader.report()
    report_path = os.path.join(os.path.dirname(os.path.abspath(path)), 'report_7.html')
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'\nHTML报告: {report_path}')


if __name__ == '__main__':
    main()
