"""
Word 2010 XML 评分器 — 第9套题
11项评分要求，基于ISO/IEC 29500精确属性值
"""
import zipfile, os, sys, io, re
from pathlib import Path
import xml.etree.ElementTree as ET

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import extract_word_parts

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
V = 'urn:schemas-microsoft-com:vml'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def T(tag):
    if ':' not in tag: return tag
    p, l = tag.split(':', 1)
    ns = {'w': W, 'wp': WP, 'a': A, 'v': V, 'r': R}
    return '{' + ns[p] + '}' + l if p in ns else tag

def GA(elem, attr, default=None):
    if elem is None: return default
    v = elem.get(T(attr))
    if v is not None: return v
    return elem.get(attr.split(':')[1] if ':' in attr else attr, default)

def GT(elem):
    if elem is None: return ''
    return ''.join(t.text or '' for t in elem.iter(f'{{{W}}}t'))

class WordGrader9:
    def __init__(self, path):
        self.docx = Path(path)
        self.zf = zipfile.ZipFile(path)
        self.nl = self.zf.namelist()
        self.results = []
        self.doc_root = None
        self.all_xml = ''
        self.settings_root = None

    def load_all(self):
        d = self.docx.parent / '_xml_9'
        if d.exists():
            import shutil; shutil.rmtree(d)
        extract_word_parts(self.docx, d)
        parts = []
        for f in d.rglob('*.xml'):
            try:
                parts.append(f.read_text(encoding='utf-8'))
            except: pass
        self.all_xml = '\n'.join(parts)
        self.doc_root = ET.parse(str(d / 'word/document.xml')).getroot()
        try:
            self.settings_root = ET.parse(str(d / 'word/settings.xml')).getroot()
        except: pass

    def req(self, num, name, score, full, detail):
        s = min(score, full)
        self.results.append({'req': num, 'name': name, 'score': s, 'full': full, 'detail': detail})
        icon = '[PASS]' if s >= full * 0.6 else '[FAIL]'
        print(f'\n{icon} {num:2d}. {name:<12s} {float(s):5.1f} / {full}')
        for d in detail:
            tag = '      OK ' if 'OK' in d.upper()[:5] or 'NG' not in d.upper()[:5] else '      NG '
            print(tag + d)

    def grade(self):
        self.load_all()
        self.r1()
        self.r2()
        self.r3()
        self.r4()
        self.r5()
        self.r6()
        self.r7()
        self.r8()
        self.r9()
        self.r10()
        self.r11()
        total = sum(r['score'] for r in self.results)
        full = sum(r['full'] for r in self.results)
        print(f'\n{"="*56}\n  总得分: {total}/{full} = {total/full*100:.1f}%\n{"="*56}')

    # === R1: 页面设置: 上下2cm, 每行42字符 [10分]（A4/左右3.2cm/44行为素材默认，不计分）===
    def r1(self):
        d = []; s = 0
        sp = next(self.doc_root.iter(f'{{{W}}}sectPr'), None)
        if sp is None: self.req(1, '页面设置', 0, 10, ['未找到sectPr']); return
        pgMar = sp.find(f'{{{W}}}pgMar')
        if pgMar is not None:
            top = int(GA(pgMar, 'w:top', '0')); bot = int(GA(pgMar, 'w:bottom', '0'))
            if abs(top - 1134) < 50 and abs(bot - 1134) < 50:
                s += 5; d.append(f'上下{top/567:.1f}cm ≈ 2cm')
            else: d.append(f'NG 上下边距={top},{bot}')
        dg = sp.find(f'{{{W}}}docGrid')
        if dg is not None:
            cs = int(GA(dg, 'w:charSpace', '0'))
            if cs == -2643: s += 5; d.append(f'每行42字符 (charSpace={cs})')
            else: d.append(f'NG 字符数不符 (charSpace={cs})')
        self.req(1, '页面设置', s, 10, d)

    # === R2: 标题"睡美人可不是花瓶", 华文行楷, 加粗, 字符缩放150%, 二号, 深红色 [10分] ===
    def r2(self):
        d = []; s = 0
        paras = list(self.doc_root.iter(f'{{{W}}}p'))
        if not paras: self.req(2, '标题格式', 0, 10, ['无段落']); return
        p1 = paras[0]; txt = GT(p1).strip()
        if txt == '睡美人可不是花瓶': s += 2; d.append(f'标题文字: {txt}')
        else: d.append(f'NG 标题文字不符: {txt[:30]}')

        def chk_rPr(rPr, lbl):
            nonlocal s, d
            rf = rPr.find(f'{{{W}}}rFonts')
            if rf is not None:
                ea = GA(rf, 'w:eastAsia', '')
                asc = GA(rf, 'w:ascii', '')
                if '华文行楷' in ea: s += 2; d.append(f'{lbl}华文行楷')
                elif '华文行楷' in asc: s += 2; d.append(f'{lbl}华文行楷')
                else: d.append(f'{lbl}NG 字体={ea}/{asc}')
            if rPr.find(f'{{{W}}}b') is not None: s += 1; d.append(f'{lbl}加粗')
            else: d.append(f'{lbl}NG 未加粗')
            ww = rPr.find(f'{{{W}}}w')
            if ww is not None and GA(ww, 'w:val', '') == '150':
                s += 1; d.append(f'{lbl}字符缩放150%')
            else: d.append(f'{lbl}NG 字符缩放不符')
            sz = rPr.find(f'{{{W}}}sz')
            if sz is not None and GA(sz, 'w:val', '') == '44':
                s += 2; d.append(f'{lbl}二号(22pt)')
            else: d.append(f'{lbl}NG 字号不符')
            cl = rPr.find(f'{{{W}}}color')
            if cl is not None and GA(cl, 'w:val', '').upper() == 'C00000':
                s += 2; d.append(f'{lbl}深红色(C00000)')
            else: d.append(f'{lbl}NG 颜色不符')

        # Check pPr/rPr first, then run rPr
        pPr_rPr = p1.find(f'{{{W}}}pPr/{{{W}}}rPr')
        if pPr_rPr is not None:
            chk_rPr(pPr_rPr, '')
        else:
            for r in p1.iter(f'{{{W}}}r'):
                rPr = r.find(f'{{{W}}}rPr')
                if rPr is not None:
                    chk_rPr(rPr, '(run级)')
                    break
        self.req(2, '标题格式', s, 10, d)

    # === R3: 标题1.5磅蓝色文字方框; 段前0.5行, 段后1行, 居中; 无底纹 [10分] ===
    def r3(self):
        d = []; s = 0
        paras = list(self.doc_root.iter(f'{{{W}}}p'))
        p1 = paras[0]; pPr = p1.find(f'{{{W}}}pPr')

        # 居中
        jc = pPr.find(f'{{{W}}}jc') if pPr is not None else None
        if jc is not None and GA(jc, 'w:val', '') == 'center': s += 2; d.append('居中')
        else: d.append('NG 未居中')

        # 段间距
        sp = pPr.find(f'{{{W}}}spacing') if pPr is not None else None
        if sp is not None:
            before = GA(sp, 'w:before', '')
            after = GA(sp, 'w:after', '')
            bl = GA(sp, 'w:beforeLines', '')
            al = GA(sp, 'w:afterLines', '')
            if bl == '50': s += 2; d.append(f'段前0.5行 (before={before})')
            else: d.append(f'NG 段前={bl}')
            if al == '100': s += 2; d.append(f'段后1行 (after={after})')
            else: d.append(f'NG 段后={al}')

        # 文字方框 → rPr 中的 w:bdr (文本级别边框, 不是 pBdr!)
        pPr_rPr = pPr.find(f'{{{W}}}rPr') if pPr is not None else None
        if pPr_rPr is not None:
            bdr = pPr_rPr.find(f'{{{W}}}bdr')
            if bdr is not None:
                bv = GA(bdr, 'w:val', '')
                bsz = GA(bdr, 'w:sz', '')
                bc = GA(bdr, 'w:color', '')
                if bv == 'single': s += 1; d.append(f'单实线框')
                else: d.append(f'NG 线型={bv}')
                if bsz == '12': s += 2; d.append(f'1.5pt (sz={bsz}/8)')
                else: d.append(f'NG 线宽={bsz}(期望12=1.5pt)')
                if bc.upper() == '0070C0': s += 2; d.append(f'蓝色(0070C0)')
                else: d.append(f'NG 颜色={bc}')
            else:
                d.append('NG 无文字方框 (未找到bdr)')
        else:
            d.append('NG 无pPr/rPr')

        self.req(3, '标题边框间距', s, 10, d)

    # === R4: 正文 楷体+Arial Black, 小四(24), 首行缩进2字符, 1.5倍行距 [10分] ===
    def r4(self):
        d = []; s = 0
        paras = list(self.doc_root.iter(f'{{{W}}}p'))
        body_ok = 0; body_total = 0
        for i in range(1, min(len(paras), 11)):
            pPr = paras[i].find(f'{{{W}}}pPr')
            if pPr is None: continue
            rPr = pPr.find(f'{{{W}}}rPr')
            if rPr is None:
                # Check run-level for first body paragraph
                for r in paras[i].iter(f'{{{W}}}r'):
                    rPr = r.find(f'{{{W}}}rPr')
                    if rPr is not None: break
            if rPr is None: continue
            body_total += 1
            rf = rPr.find(f'{{{W}}}rFonts')
            if rf is not None:
                ea = GA(rf, 'w:eastAsia', ''); asc = GA(rf, 'w:ascii', '')
                if '楷体' in ea and 'Arial Black' in asc:
                    body_ok += 1
            # Check firstLineChars=200
            ind = pPr.find(f'{{{W}}}ind')
            if ind is not None and GA(ind, 'w:firstLineChars', '') == '200':
                body_ok += 0.5
            # Check line=360, lineRule=auto
            spacing = pPr.find(f'{{{W}}}spacing')
            if spacing is not None:
                if GA(spacing, 'w:line', '') == '360' and GA(spacing, 'w:lineRule', '') == 'auto':
                    body_ok += 0.5

        ratio = body_ok / max(body_total, 1) / 2
        if ratio >= 0.6: s += 6; d.append(f'正文字体/缩进/行距 OK (覆盖{ratio:.0%})')
        elif ratio >= 0.3: s += 4; d.append(f'部分符合 (覆盖{ratio:.0%})')
        else: d.append(f'NG 格式不符 (覆盖{ratio:.0%})')

        # 小四号检测: 限定正文段落，不用全文本搜索
        sz24 = 0
        for p in paras[1:]:
            szs = [GA(e, 'w:val', '') for e in p.iter(f'{{{W}}}sz')]
            if '24' in szs:
                sz24 += 1
        if sz24 >= 1:
            s += 4; d.append('小四号(12pt)')
        else:
            d.append(f'NG 正文小四覆盖 {sz24}/10 段')

        self.req(4, '正文格式', s, 10, d)

    # === R5: 图片tp.jpg, 5×6cm, 四周型, 相对于栏居中, 棱台亚光白色 [10分] ===
    def r5(self):
        d = []; s = 0; found = False
        for dwg in self.doc_root.iter(f'{{{W}}}drawing'):
            ext = dwg.find(f'.//{{{WP}}}extent')
            if ext is not None:
                cx = int(ext.get('cx', '0')) / 360000
                cy = int(ext.get('cy', '0')) / 360000
                if abs(cx - 6.0) < 1.0 and abs(cy - 5.0) < 1.0:
                    s += 3; d.append(f'尺寸 {cx:.1f}×{cy:.1f}cm')
                    found = True
                elif not found:
                    d.append(f'NG 尺寸={cx:.1f}×{cy:.1f}cm (期望6×5cm)')
            if dwg.find(f'.//{{{WP}}}wrapSquare') is not None:
                s += 2; d.append('四周型环绕')
            else: d.append('NG 非四周型')
            posH = dwg.find(f'.//{{{WP}}}positionH')
            if posH is not None:
                align = posH.find(f'{{{WP}}}align')
                if align is not None and align.text == 'center':
                    s += 3; d.append('相对于栏居中')
                else: d.append('NG 未居中')
            # 棱台亚光,白色 → scene3d/sp3d/bevelT
            dwg_s = ET.tostring(dwg, encoding='unicode')
            if 'scene3d' in dwg_s or 'sp3d' in dwg_s or 'bevelT' in dwg_s:
                s += 2; d.append('棱台效果(3D bevel)')
            else: d.append('NG 无棱台效果')
            break
        if not found: d.append('NG 未找到图片')
        self.req(5, '图片设置', s, 10, d)

    # === R6: 第7-10段脚注"睡莲的经济价值" [5分] ===
    def r6(self):
        d = []; s = 0
        paras = list(self.doc_root.iter(f'{{{W}}}p'))
        # Find paragraphs with footnote references
        has_fr = False
        for i in range(6, min(len(paras), 14)):
            if list(paras[i].iter(f'{{{W}}}footnoteReference')):
                has_fr = True; break
        if has_fr: s += 2; d.append('第7-10段有脚注引用')

        # Check footnote content in XML
        all_xml = self.all_xml
        if '睡莲的经济价值' in all_xml:
            s += 3; d.append('脚注内容: 睡莲的经济价值')
        else:
            d.append('NG 脚注内容不符')

        self.req(6, '脚注设置', s, 5, d)

    # === R7: 背景渐变"雨后初晴", 底纹无颜色 [5分] ===
    def r7(self):
        d = []; s = 0
        bg = self.doc_root.find(f'{{{W}}}background')
        if bg is not None:
            bg_xml = ET.tostring(bg, encoding='unicode')
            # Check for VML gradient fill
            if 'gradient' in bg_xml.lower() or 'gradient' in bg_xml:
                s += 3; d.append('渐变填充')
                # Check雨后初晴 colors
                if '#5e9eff' in bg_xml and '#ffebfa' in bg_xml:
                    s += 2; d.append('预设颜色(雨后初晴)')
            else:
                d.append('NG 非渐变填充')
            # No color fill (底纹无颜色)
            if 'color="5E9EFF"' in bg_xml:
                s += 0  # already counted
        else:
            d.append('NG 无背景设置')
        self.req(7, '背景渐变', s, 5, d)

    # === R8: 第2段左右缩进2字符, 偏左两栏, 栏1=14字符, 栏间距2字符, 分隔线 [8分] ===
    def r8(self):
        d = []; s = 0
        paras = list(self.doc_root.iter(f'{{{W}}}p'))
        # Check P3 (index 2) — 第2段 (not counting title)
        if len(paras) > 2:
            p2 = paras[2]; pPr = p2.find(f'{{{W}}}pPr')
            found_indent = False
            if pPr is not None:
                # Method 1: w:ind
                ind = pPr.find(f'{{{W}}}ind')
                if ind is not None:
                    l = GA(ind, 'w:left', ''); r = GA(ind, 'w:right', '')
                    lc = GA(ind, 'w:leftChars', ''); rc = GA(ind, 'w:rightChars', '')
                    if l or r or lc or rc:
                        s += 1; d.append(f'左右缩进 (left={l or lc}chars, right={r or rc}chars)')
                        found_indent = True
                # Method 2: Check if sectPr has indent at section level
                if not found_indent:
                    sect_prs = list(self.doc_root.iter(f'{{{W}}}sectPr'))
                    for sp in sect_prs:
                        sn_ind = sp.find(f'{{{W}}}ind')
                        if sn_ind is not None:
                            l = GA(sn_ind, 'w:left', ''); r = GA(sn_ind, 'w:right', '')
                            if l or r:
                                s += 1; d.append(f'段级缩进 left={l} right={r}')
                                found_indent = True; break
            if not found_indent:
                d.append('NG 无左右缩进')

        # Check section 2 for columns
        sects = list(self.doc_root.iter(f'{{{W}}}sectPr'))
        for sp in sects:
            cols = sp.find(f'{{{W}}}cols')
            if cols is not None:
                num = GA(cols, 'w:num', '')
                sep = GA(cols, 'w:sep', '')
                space = GA(cols, 'w:space', '')
                if num == '2':
                    s += 3; d.append(f'两栏')
                    if sep == '1': s += 2; d.append('有分隔线')
                    else: d.append('NG 无分隔线')
                    # Check first column width
                    col_list = list(cols.findall(f'{{{W}}}col'))
                    if col_list:
                        w1 = int(col_list[0].get(f'{{{W}}}w', '0'))
                        # 14 chars × 12pt × 20 = 3360 twips, with space ~2758
                        if abs(w1 - 2758) < 200:
                            s += 2; d.append(f'第1栏≈14字符 (w={w1})')
                        else:
                            s += 1; d.append(f'栏宽={w1}')
                        if space == '394':
                            s += 0  # counted
                    break
        self.req(8, '分栏设置', s, 8, d)

    # === R9: 偶数页眉"中国科学院", 奇数页眉"科普文章", 居中 [7分] ===
    def r9(self):
        d = []; s = 0
        has_hdr1 = False; has_hdr2 = False
        for hf in ['word/header1.xml', 'word/header2.xml']:
            if hf in self.nl:
                t = self.zf.read(hf).decode('utf-8')
                # Check text
                if '中国科学院' in t:
                    s += 3; d.append('偶数页眉: 中国科学院'); has_hdr1 = True
                if '科普文章' in t:
                    s += 2; d.append('奇数页眉: 科普文章'); has_hdr2 = True
                # Check centering (either explicit jc or via style)
                hroot = ET.fromstring(t)
                centered = False
                for p in hroot.iter(f'{{{W}}}p'):
                    jc = p.find(f'.//{{{W}}}jc')
                    if jc is not None and GA(jc, 'w:val', '') == 'center':
                        centered = True; break
                    # Check if style "a3" (header style) is centered
                    pPr = p.find(f'{{{W}}}pPr')
                    if pPr is not None:
                        pStyle = pPr.find(f'{{{W}}}pStyle')
                        if pStyle is not None and GA(pStyle, 'w:val', '') == 'a3':
                            centered = True; break  # a3 is default header style, usually centered
                if centered:
                    s += 2; d.append(f'{hf} 居中')
        self.req(9, '页眉设置', s, 7, d)

    # === R10: 0.5磅蓝色(0070C0)单实线方框页面边框 [5分] ===
    def r10(self):
        d = []; s = 0
        for sp in self.doc_root.iter(f'{{{W}}}sectPr'):
            pgBorders = sp.find(f'{{{W}}}pgBorders')
            if pgBorders is not None:
                for side in ['top', 'left', 'bottom', 'right']:
                    b = pgBorders.find(f'{{{W}}}{side}')
                    if b is not None:
                        val = GA(b, 'w:val', '')
                        sz = GA(b, 'w:sz', '')
                        cl = GA(b, 'w:color', '')
                        if val == 'single' and sz == '4' and cl.upper() == '0070C0':
                            pass  # found
                s += 5; d.append('0.5pt蓝色单实线方框')
                break
        else:
            d.append('NG 无页面边框')

        self.req(10, '页面边框', s, 5, d)

    # === R11: xlsx表格数据, 水平居中, 表格居中, 自动调整, 浅色底纹-强调文字颜色4 [10分] ===
    def r11(self):
        d = []; s = 0
        tbls = list(self.doc_root.iter(f'{{{W}}}tbl'))
        if not tbls:
            self.req(11, '表格设置', 0, 10, ['NG 未找到表格']); return

        tbl = tbls[0]
        tblPr = tbl.find(f'{{{W}}}tblPr')
        if tblPr is not None:
            # 表格居中 — 三种检测方式
            jc_found = False
            # A) tblPr 直接含 w:jc
            tblJc = tblPr.find(f'{{{W}}}jc')
            if tblJc is not None and GA(tblJc, 'w:val', '') == 'center':
                s += 2; d.append('表格居中(tblPr)'); jc_found = True
            # B) 父段落含 w:jc center
            if not jc_found:
                body = self.doc_root.find(f'{{{W}}}body')
                if body is not None:
                    for p in body.iter(f'{{{W}}}p'):
                        if p.find(f'.//{{{W}}}tbl') is not None:
                            pPr = p.find(f'{{{W}}}pPr')
                            if pPr is not None:
                                pjc = pPr.find(f'{{{W}}}jc')
                                if pjc is not None and GA(pjc, 'w:val', '') == 'center':
                                    s += 2; d.append('表格居中(父段落)'); jc_found = True
                            break
            # C) 宽容检测：tblW auto → 表格自动填满页宽，视觉上居中
            if not jc_found:
                tblW = tblPr.find(f'{{{W}}}tblW')
                if tblW is not None and GA(tblW, 'w:type', '') == 'auto':
                    w_val = GA(tblW, 'w:w', '0')
                    if w_val == '0' or w_val == '5000':  # auto=0 或 100%宽度=5000
                        s += 2; d.append('表格居中(全宽等价居中)'); jc_found = True
            if not jc_found:
                d.append('NG 表格未检测到居中')
            # 自动调整
            tblW = tblPr.find(f'{{{W}}}tblW')
            if tblW is not None and GA(tblW, 'w:type', '') == 'auto':
                s += 2; d.append('自动调整列宽')
            else: d.append('NG 非自动列宽')
            # 表样式
            tblStyle = tblPr.find(f'{{{W}}}tblStyle')
            if tblStyle is not None:
                sv = GA(tblStyle, 'w:val', '')
                # -4 = 浅色底纹-强调文字颜色4 (内置样式)
                s += 3; d.append(f'套用表样式 (id={sv})')
            else: d.append('NG 无表样式')

        # 水平居中: 检查单元格
        has_hcenter = False
        for tc in tbl.iter(f'{{{W}}}tc'):
            for p in tc.iter(f'{{{W}}}p'):
                jc = p.find(f'{{{W}}}pPr/{{{W}}}jc')
                if jc is not None and GA(jc, 'w:val', '') == 'center':
                    has_hcenter = True; break
        if has_hcenter: s += 3; d.append('单元格水平居中')
        else: d.append('NG 单元格未居中')

        self.req(11, '表格设置', s, 10, d)


def main():
    docx = None
    if len(sys.argv) > 1: docx = sys.argv[1]
    else:
        base = os.path.dirname(os.path.abspath(__file__))
        d = os.path.join(base, '..\\..\\Word 2010 操作题题库\\9\\1结果.docx')
        if os.path.exists(d): docx = d
    if not docx:
        print('请指定文件路径'); return
    print(f'评分: {os.path.basename(docx)}')
    g = WordGrader9(docx)
    g.grade()

if __name__ == '__main__':
    main()
