"""
通用修订分析器 — 适用于任意文件夹
用法: python analyze.py N
分析结果写入 N/analysis.txt
"""
import zipfile, os, sys, io, re, xml.etree.ElementTree as ET

# === 通用工具函数 ===
W  = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
V  = '{urn:schemas-microsoft-com:vml}'
O  = '{urn:schemas-microsoft-com:office:office}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
R  = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

NS = {
    'w':  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'v':  'urn:schemas-microsoft-com:vml',
    'o':  'urn:schemas-microsoft-com:office:office',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':  'http://schemas.openxmlformats.org/drawingml/2006/main',
}

def child(e, t):
    for c in e:
        if c.tag == t: return c
    return None

def children(e, t):
    return [c for c in e if c.tag == t]

def GA(elem, attr, default=None):
    if elem is None: return default
    if ':' in attr:
        ns, l = attr.split(':', 1)
        return elem.get('{' + NS.get(ns, '') + '}' + l, default)
    return elem.get(attr, default)

def gt(elem):
    return ''.join(t.text or '' for t in elem.iter(W + 't'))

def parse_rpr(rPr):
    """解析运行属性"""
    if rPr is None: return {}
    p = {}
    for c in rPr:
        tn = c.tag.replace(W, '')
        if tn == 'rFonts':
            f = GA(c, 'w:eastAsia') or GA(c, 'w:ascii')
            if f: p['font'] = f
        elif tn == 'b': p['bold'] = True
        elif tn == 'i': p['italic'] = True
        elif tn == 'sz': p['sz'] = int(GA(c, 'w:val', '0'))
        elif tn == 'color': p['color'] = GA(c, 'w:val', '')
        elif tn == 'spacing': p['charSp'] = int(GA(c, 'w:val', '0'))
        elif tn == 'w': p['charScale'] = GA(c, 'w:val', '')
    return p

def parse_ppr(pPr):
    """解析段落属性"""
    if pPr is None: return {}
    p = {}
    for c in pPr:
        tn = c.tag.replace(W, '')
        if tn == 'jc': p['align'] = GA(c, 'w:val', '')
        elif tn == 'pStyle': p['styleId'] = GA(c, 'w:val', '')
        elif tn == 'spacing':
            if GA(c, 'w:line'): p['line'] = int(GA(c, 'w:line', '0'))
            if GA(c, 'w:lineRule'): p['lineRule'] = GA(c, 'w:lineRule', '')
            if GA(c, 'w:beforeLines'): p['beforeLines'] = int(GA(c, 'w:beforeLines', '0'))
            if GA(c, 'w:afterLines'): p['afterLines'] = int(GA(c, 'w:afterLines', '0'))
        elif tn == 'ind':
            if GA(c, 'w:firstLine'): p['indentFL'] = int(GA(c, 'w:firstLine', '0'))
            if GA(c, 'w:firstLineChars'): p['indentFC'] = int(GA(c, 'w:firstLineChars', '0'))
            if GA(c, 'w:leftChars'): p['leftChars'] = int(GA(c, 'w:leftChars', '0'))
            if GA(c, 'w:rightChars'): p['rightChars'] = int(GA(c, 'w:rightChars', '0'))
        elif tn == 'pBdr':
            for bt in ['top', 'bottom', 'left', 'right']:
                b = child(c, W + bt)
                if b is not None:
                    p[f'border_{bt}'] = f"{GA(b,'w:val','')}|{GA(b,'w:sz','')}|{GA(b,'w:color','')}"
        elif tn == 'shd':
            p['shd_fill'] = GA(c, 'w:fill', '')
        elif tn == 'outlineLvl': p['outline'] = int(GA(c, 'w:val', '0'))
        elif tn == 'numPr':
            ni = child(c, W + 'numId')
            if ni: p['numId'] = int(GA(ni, 'w:val', '0'))
    return p

def parse_sectPr(sp):
    """解析节属性"""
    if sp is None: return {}
    ps = {}
    pgSz = child(sp, W + 'pgSz')
    if pgSz is not None:
        ps['pg_w'] = int(GA(pgSz, 'w:w', '0'))
        ps['pg_h'] = int(GA(pgSz, 'w:h', '0'))
    pgMar = child(sp, W + 'pgMar')
    if pgMar is not None:
        for k in ['top', 'bottom', 'left', 'right', 'header', 'footer']:
            v = GA(pgMar, 'w:' + k, '')
            if v: ps[f'mar_{k}'] = int(v)
    cols = child(sp, W + 'cols')
    if cols is not None:
        ps['cols_num'] = GA(cols, 'w:num', '1')
        ps['cols_sep'] = GA(cols, 'w:sep', '')
        ps['cols_space'] = GA(cols, 'w:space', '')
        cl = children(cols, W + 'col')
        if cl: ps['col1_w'] = GA(cl[0], 'w:w', '')
    dg = child(sp, W + 'docGrid')
    if dg is not None:
        ps['grid_type'] = GA(dg, 'w:type', '')
        ps['grid_linePitch'] = GA(dg, 'w:linePitch', '')
    return ps


# === 分析函数 ===
def analyze_answer(base_folder):
    """分析标准答案的 XML 结构"""
    ans_path = os.path.join(base_folder, '1结果.docx')
    L = []

    with zipfile.ZipFile(ans_path) as z:
        doc = ET.fromstring(z.read('word/document.xml'))
        for c in doc:
            if c.tag == W + 'body': body = c; break

        # 段落分析
        L.append('=' * 60 + '\n[ANSWER PARAGRAPHS]')
        paras = children(body, W + 'p')
        for pi, p in enumerate(paras):
            txt = gt(p)
            # 图片/形状段落
            if not txt.strip():
                has_pict = child(p, W + 'pict') is not None
                has_dwg = any(1 for _ in p.iter(WP + 'anchor'))
                if has_pict:
                    for sh in p.iter(V + 'shape'):
                        L.append(f'P{pi:02d}: [VML] type={sh.get("type","")} id={sh.get("id","")}')
                        tp = sh.find(V + 'textpath')
                        if tp is not None: L.append(f'  textpath: "{tp.get("string","")}"')
                        for tb in sh.iter(V + 'textbox'):
                            L.append(f'  textbox: "{gt(tb).strip()}"')
                        for fl in sh.iter(V + 'fill'):
                            L.append(f'  fill: type={fl.get("type","")}')
                        L.append(f'  fillcolor={sh.get("fillcolor","")}')
                        L.append(f'  strokecolor={sh.get("strokecolor","")}')
                if has_dwg:
                    for anc in p.iter(WP + 'anchor'):
                        ext = child(anc, WP + 'extent')
                        if ext is not None:
                            cx = int(ext.get('cx', '0')) / 914400 * 2.54
                            cy = int(ext.get('cy', '0')) / 914400 * 2.54
                            L.append(f'P{pi:02d}: [IMAGE] {cx:.1f}x{cy:.1f}cm')
                continue

            pPr = child(p, W + 'pPr')
            pp = parse_ppr(pPr)
            rp = parse_rpr(child(pPr, W + 'rPr') if pPr is not None else None)
            L.append(f'\nP{pi:02d}: "{txt[:80]}"')
            if pp: L.append(f'  pPr: {pp}')
            if rp: L.append(f'  pPr/rPr: {rp}')

            # 超链接
            for hl in children(p, W + 'hyperlink'):
                rid = GA(hl, 'r:id', '')
                L.append(f'  HYPERLINK: rid={rid} text="{gt(hl)[:40]}"')

        # 节属性
        L.append('\n\n[ANSWER SECTPR]')
        for i, sp in enumerate(body.iter(W + 'sectPr')):
            ps = parse_sectPr(sp)
            if ps: L.append(f'  SectPr{i}: {ps}')

        # 表格
        L.append('\n[ANSWER TABLES]')
        for ti, tbl in enumerate(body.iter(W + 'tbl')):
            rows = children(tbl, W + 'tr')
            L.append(f'  Table{ti}: {len(rows)} rows')
            tblPr = child(tbl, W + 'tblPr')
            if tblPr is not None:
                for c in tblPr:
                    tn = c.tag.replace(W, '')
                    if tn == 'jc': L.append(f'    jc: {GA(c,"w:val","")}')
                    elif tn == 'tblW': L.append(f'    tblW: {GA(c,"w:type","")}')
                    elif tn == 'tblBorders':
                        for bt in ['top', 'bottom', 'left', 'right', 'insideH', 'insideV']:
                            b = child(c, W + bt)
                            if b is not None:
                                L.append(f'    {bt}: {GA(b,"w:val","")}|{GA(b,"w:sz","")}|{GA(b,"w:color","")}')
            for ri, row in enumerate(rows[:4]):
                for ci, cell in enumerate(children(row, W + 'tc')[:3]):
                    ct = gt(cell)
                    tcPr = child(cell, W + 'tcPr')
                    info = []
                    if tcPr is not None:
                        h = child(tcPr, W + 'trHeight')
                        if h is not None: info.append(f'h={GA(h,"w:val","")}')
                        shd = child(tcPr, W + 'shd')
                        if shd is not None: info.append(f'shd={GA(shd,"w:fill","")}')
                    for r in cell.iter(W + 'r'):
                        rp = parse_rpr(child(r, W + 'rPr'))
                        if rp: info.append(str(rp)); break
                    L.append(f'    R{ri}C{ci}: "{ct[:20]}" {" ".join(info)}')

        # 页眉页脚
        L.append('\n[ANSWER HEADERS/FOOTERS]')
        for name in sorted(z.namelist()):
            if 'header' in name or 'footer' in name:
                try:
                    x = z.read(name).decode('utf-8')
                    root = ET.fromstring(x)
                    ht = gt(root).strip()[:60]
                    has_wm = any('Watermark' in s.get('id', '') for s in root.iter(V + 'shape'))
                    L.append(f'  {name}: text="{ht}" watermark={has_wm}')
                except: pass

        # 关系(超链接)
        L.append('\n[ANSWER RELS]')
        try:
            rels = ET.fromstring(z.read('word/_rels/document.xml.rels'))
            for rel in rels:
                if 'hyperlink' in rel.get('Type', ''):
                    L.append(f'  {rel.get("Id","")}: {rel.get("Target","")}')
        except: pass

        # 编号
        L.append('\n[ANSWER NUMBERING]')
        try:
            num = ET.fromstring(z.read('word/numbering.xml'))
            for absNum in num:
                aid = absNum.get(W + 'abstractNumId', '')
                for lvl in absNum.iter(W + 'lvl'):
                    nf = child(lvl, W + 'numFmt')
                    lt = child(lvl, W + 'lvlText')
                    if nf is not None:
                        L.append(f'  abs{aid}/lvl{lvl.get(W+"ilvl","")}: fmt={GA(nf,"w:val","")} text={GA(lt,"w:val","") if lt else ""}')
        except: pass

    return '\n'.join(L)


def analyze_revisions(base_folder):
    """分析修订标记 — 素材对比答案的差异"""
    rev_path = os.path.join(base_folder, '2素材.docx')
    L = []

    with zipfile.ZipFile(rev_path) as z:
        doc = ET.fromstring(z.read('word/document.xml'))
        for c in doc:
            if c.tag == W + 'body': body = c; break

        # 节属性变更
        L.append('\n[SECTPR CHANGES]')
        for sp in body.iter(W + 'sectPr'):
            chg = child(sp, W + 'sectPrChange')
            if chg is not None:
                old_sp = child(chg, W + 'sectPr')
                cur = parse_sectPr(sp)
                old = parse_sectPr(old_sp)
                for k, v in cur.items():
                    ov = old.get(k)
                    if v != ov: L.append(f'  {k}: {ov} → {v}')

        # 段落级修订
        L.append('\n[PARAGRAPH REVISIONS]')
        paras = children(body, W + 'p')
        for pi, p in enumerate(paras):
            txt = gt(p).strip()
            has_ins = any(1 for _ in p.iter(W + 'ins'))
            has_del = any(1 for _ in p.iter(W + 'del'))

            pPr = child(p, W + 'pPr')
            ppr_chg = child(pPr, W + 'pPrChange') if pPr is not None else None

            # rPr 变更
            rpr_changes = []
            for r in children(p, W + 'r'):
                rPr = child(r, W + 'rPr')
                if rPr is not None:
                    rpc = child(rPr, W + 'rPrChange')
                    if rpc is not None:
                        old_rp = parse_rpr(child(rpc, W + 'rPr'))
                        new_rp = {}
                        for c in rPr:
                            if c.tag == W + 'rPrChange': continue
                            tn = c.tag.replace(W, '')
                            if tn == 'rFonts':
                                f = GA(c, 'w:eastAsia') or GA(c, 'w:ascii')
                                if f: new_rp['font'] = f
                            elif tn == 'sz': new_rp['sz'] = int(GA(c, 'w:val', '0'))
                            elif tn == 'b': new_rp['bold'] = True
                            elif tn == 'color': new_rp['color'] = GA(c, 'w:val', '')
                            elif tn == 'spacing': new_rp['charSp'] = int(GA(c, 'w:val', '0'))
                            elif tn == 'w': new_rp['charScale'] = GA(c, 'w:val', '')
                        rt = ''.join(t.text or '' for t in children(r, W + 't'))
                        rpr_changes.append((rt, old_rp, new_rp))

            # 跳过空段落（无修订）
            if not txt and not has_ins and not has_del and not rpr_changes and ppr_chg is None:
                continue

            L.append(f'\nP{pi:02d}: "{txt[:60]}"')

            # 段落属性变更
            if ppr_chg is not None:
                old_ppr = parse_ppr(child(ppr_chg, W + 'pPr'))
                new_ppr = {}
                for c in pPr:
                    if c.tag == W + 'pPrChange': continue
                    tn = c.tag.replace(W, '')
                    if tn == 'jc': new_ppr['align'] = GA(c, 'w:val', '')
                    elif tn == 'spacing':
                        if GA(c, 'w:line'): new_ppr['line'] = int(GA(c, 'w:line', '0'))
                        if GA(c, 'w:beforeLines'): new_ppr['beforeLines'] = int(GA(c, 'w:beforeLines', '0'))
                        if GA(c, 'w:afterLines'): new_ppr['afterLines'] = int(GA(c, 'w:afterLines', '0'))
                    elif tn == 'ind':
                        if GA(c, 'w:firstLine'): new_ppr['indentFL'] = int(GA(c, 'w:firstLine', '0'))
                        if GA(c, 'w:firstLineChars'): new_ppr['indentFC'] = int(GA(c, 'w:firstLineChars', '0'))
                    elif tn == 'pBdr':
                        for bt in ['top', 'bottom', 'left', 'right']:
                            b = child(c, W + bt)
                            if b is not None:
                                new_ppr[f'border_{bt}'] = f"{GA(b,'w:val','')}|{GA(b,'w:sz','')}|{GA(b,'w:color','')}"
                    elif tn == 'shd': new_ppr['shd'] = GA(c, 'w:fill', '')
                    elif tn == 'pStyle': new_ppr['styleId'] = GA(c, 'w:val', '')
                    elif tn == 'numPr':
                        ni = child(c, W + 'numId')
                        if ni: new_ppr['numId'] = int(GA(ni, 'w:val', '0'))
                for k, v in new_ppr.items():
                    ov = old_ppr.get(k)
                    if ov != v:
                        L.append(f'  pPr CHG: {k}: {ov} → {v}')

            # run属性变更
            for txt_s, old, new in rpr_changes:
                short = f'"{txt_s[:40]}"'
                L.append(f'  rPr CHG {short}:')
                for k, v in new.items():
                    ov = old.get(k)
                    if ov != v:
                        L.append(f'    {k}: {ov} → {v}')

            # 插入
            for ins in p.iter(W + 'ins'):
                it = gt(ins).strip()
                # VML
                for sh in ins.iter(V + 'shape'):
                    L.append(f'  INSERTED VML: type={sh.get("type","")} id={sh.get("id","")}')
                    for tb in sh.iter(V + 'textbox'):
                        L.append(f'    textbox: "{gt(tb).strip()}"')
                # 图片
                for anc in ins.iter(WP + 'anchor'):
                    ext = child(anc, WP + 'extent')
                    if ext is not None:
                        cx = int(ext.get('cx', '0')) / 914400 * 2.54
                        cy = int(ext.get('cy', '0')) / 914400 * 2.54
                        L.append(f'  INSERTED IMAGE: {cx:.1f}x{cy:.1f}cm')
                # 格式
                for r in ins.iter(W + 'r'):
                    rp = parse_rpr(child(r, W + 'rPr'))
                    if rp: L.append(f'  INSERTED fmt: {rp}'); break
                if it: L.append(f'  INSERTED: "{it[:80]}"')

            # 删除
            for d in p.iter(W + 'del'):
                dt = gt(d).strip()
                if dt: L.append(f'  DELETED: "{dt[:80]}"')

    return '\n'.join(L)


# === 主流程 ===
def main():
    if len(sys.argv) < 2:
        print('用法: python analyze.py <文件夹路径>')
        print('  analyze.py 8           — 当前目录下的 8/')
        print('  analyze.py /path/to/8   — 绝对路径')
        sys.exit(1)

    folder = sys.argv[1]
    if not os.path.isabs(folder):
        base = os.path.join(os.getcwd(), folder)
    else:
        base = folder
    if not os.path.isdir(base):
        print(f'错误: 文件夹 {folder} 不存在')
        sys.exit(1)

    print(f'分析文件夹 {folder} ...')

    # 答案分析
    ans_report = analyze_answer(base)

    # 修订分析（素材可能已覆盖为对比文档）
    rev_report = analyze_revisions(base)

    full_report = f'''================================================================================
FOLDER {folder} — 综合分析报告
================================================================================

{ans_report}

================================================================================
REVISION ANALYSIS（修订标记）
================================================================================
{rev_report}
'''

    out_path = os.path.join(base, 'analysis.md')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(full_report)

    print(f'完成 → {out_path}')
    print(f'  答案分析: {len(ans_report)} 字符')
    print(f'  修订分析: {len(rev_report)} 字符')


if __name__ == '__main__':
    main()
