#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 2010 操作题 通用分析器
不依赖修订标记，直接从答案文件读取完整结构

用法: python word_analyze.py N
  N = 文件夹编号
"""
import zipfile, xml.etree.ElementTree as ET, sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
V = 'urn:schemas-microsoft-com:vml'

def T(tag):
    if ':' not in tag: return tag
    p, l = tag.split(':', 1)
    ns = {'w': W, 'wp': WP, 'a': A_NS, 'pic': PIC, 'v': V}
    return '{' + ns[p] + '}' + l if p in ns else tag

def GA(elem, attr, default=None):
    if elem is None: return default
    v = elem.get(T(attr))
    if v is not None: return v
    return elem.get(attr.split(':')[1] if ':' in attr else attr, default)

def GT(elem):
    if elem is None: return ''
    return ''.join(t.text or '' for t in elem.iter(T('w:t')))

def analyze(folder_num):
    base = Path(os.path.dirname(__file__))
    mat_dir = base.parent.parent.parent / 'Word 2010 操作题题库' / str(folder_num)
    ans_path = mat_dir / '1结果.docx'

    if not ans_path.exists():
        ans_path = mat_dir / '结果.docx'   # folder 1 uses this name

    if not ans_path.exists():
        print(f'错误: 找不到答案文件 in {mat_dir}'); return

    zf = zipfile.ZipFile(str(ans_path))
    nl = zf.namelist()

    print(f'文件: {ans_path.name}')
    print(f'ZIP条目: {len(nl)}')
    print(f'图片: {[f for f in nl if "media" in f]}')

    # 1. 段落文本 (前15段)
    doc = zf.read('word/document.xml').decode('utf-8')
    root = ET.fromstring(doc)
    print(f'\n=== 段落 ({len(list(root.iter(T("w:p"))))} 段) ===')
    paras = list(root.iter(T('w:p')))
    for i, p in enumerate(paras[:15]):
        text = GT(p).strip()
        if text:
            pPr = p.find(T('w:pPr'))
            pStyle = pPr.find(T('w:pStyle')) if pPr is not None else None
            style = GA(pStyle, 'w:val', '?') if pStyle is not None else '?'
            print(f'  [{i+1}] style={style} | {text[:80]}')

    # 2. 段落格式详情 (前5段)
    print(f'\n=== 段落格式 (前5段) ===')
    for i, p in enumerate(paras[:5]):
        text = GT(p).strip()
        if not text: continue
        pPr = p.find(T('w:pPr'))
        if pPr is None: continue
        rPr = pPr.find(T('w:rPr')) if pPr is not None else None
        ind = pPr.find(T('w:ind'))
        spacing = pPr.find(T('w:spacing'))
        jc = pPr.find(T('w:jc'))

        print(f'  [{i+1}] {text[:50]}')
        if rPr is not None:
            rf = rPr.find(T('w:rFonts'))
            if rf is not None:
                print(f'    font: eastAsia={GA(rf, "w:eastAsia")} ascii={GA(rf, "w:ascii")}')
            sz = rPr.find(T('w:sz'))
            if sz is not None: print(f'    sz={GA(sz, "w:val")} (={int(GA(sz,"w:val","0"))/2:.0f}pt)')
            b = rPr.find(T('w:b'))
            if b is not None: print(f'    bold=YES')
            cl = rPr.find(T('w:color'))
            if cl is not None: print(f'    color={GA(cl, "w:val")}')
            ww = rPr.find(T('w:w'))
            if ww is not None: print(f'    charScale={GA(ww, "w:val")}%')
            bdr = rPr.find(T('w:bdr'))
            if bdr is not None:
                print(f'    bdr: val={GA(bdr,"w:val")} sz={GA(bdr,"w:sz")}(={int(GA(bdr,"w:sz","0"))/8:.1f}pt) color={GA(bdr,"w:color")}')
        if ind is not None:
            print(f'    indent: left={GA(ind,"w:left")} right={GA(ind,"w:right")} firstLine={GA(ind,"w:firstLine")} leftChars={GA(ind,"w:leftChars")} rightChars={GA(ind,"w:rightChars")} firstLineChars={GA(ind,"w:firstLineChars")}')
        if spacing is not None:
            print(f'    spacing: before={GA(spacing,"w:before")} after={GA(spacing,"w:after")} line={GA(spacing,"w:line")} lineRule={GA(spacing,"w:lineRule")} beforeLines={GA(spacing,"w:beforeLines")} afterLines={GA(spacing,"w:afterLines")}')
        if jc is not None:
            print(f'    align={GA(jc, "w:val")}')

    # 3. 节属性
    print(f'\n=== 节属性 ===')
    for ei, sp in enumerate(root.iter(T('w:sectPr'))):
        pgSz = sp.find(T('w:pgSz'))
        pgMar = sp.find(T('w:pgMar'))
        cols = sp.find(T('w:cols'))
        dg = sp.find(T('w:docGrid'))
        pgBorders = sp.find(T('w:pgBorders'))
        print(f'  SectPr {ei+1}:')
        if pgSz is not None:
            print(f'    pgSz: {int(GA(pgSz,"w:w","0"))/1440:.1f}×{int(GA(pgSz,"w:h","0"))/1440:.1f}in code={GA(pgSz,"w:code")}')
        if pgMar is not None:
            print(f'    pgMar: top={int(GA(pgMar,"w:top","0"))/567:.1f}cm bottom={int(GA(pgMar,"w:bottom","0"))/567:.1f}cm left={int(GA(pgMar,"w:left","0"))/567:.1f}cm right={int(GA(pgMar,"w:right","0"))/567:.1f}cm')
        if cols is not None:
            col_list = list(cols.findall(T('w:col')))
            print(f'    cols: num={GA(cols,"w:num")} sep={GA(cols,"w:sep")} space={GA(cols,"w:space")}')
            for ci, c in enumerate(col_list):
                print(f'      col{ci+1}: w={c.get(T("w:w"))} space={c.get(T("w:space"))}')
        if dg is not None:
            print(f'    docGrid: linePitch={GA(dg,"w:linePitch")} charSpace={GA(dg,"w:charSpace")}')
        if pgBorders is not None:
            for side in ['top', 'left', 'bottom', 'right']:
                b = pgBorders.find(T('w:' + side))
                if b is not None:
                    print(f'    pgBorder.{side}: val={GA(b,"w:val")} sz={GA(b,"w:sz")} color={GA(b,"w:color")}')

    # 4. 页眉/页脚
    print(f'\n=== 页眉页脚 ===')
    for hf in sorted(nl):
        if 'header' in hf.lower() or 'footer' in hf.lower():
            t = zf.read(hf).decode('utf-8')
            hroot = ET.fromstring(t)
            texts = []
            for p in hroot.iter(T('w:p')):
                jc = p.find(T('w:pPr/' + T('w:jc'))) if p.find(T('w:pPr')) is not None else None
                txt = GT(p).strip()
                if txt:
                    al = GA(jc, 'w:val', '?') if jc is not None else '?'
                    texts.append(f'{txt}(align={al})')
            print(f'  {hf}: {", ".join(texts)}' if texts else f'  {hf}: (empty)')

    # 5. 脚注/尾注
    for fn in ['word/footnotes.xml', 'word/endnotes.xml']:
        if fn in nl:
            ft = zf.read(fn).decode('utf-8')
            for note in ET.fromstring(ft).iter(T('w:footnote')):
                fid = note.get(T('w:id'), '?')
                txt = GT(note).strip()
                if txt:
                    # Check numbering format
                    nref = note.find(f'.//{{{W}}}footnoteRef')
                    nr = note.find(f'.//{{{W}}}numPr')
                    print(f'  {fn}[{fid}]: {txt[:80]}')

    # 6. 图片/绘图
    print(f'\n=== 绘图/图片 ===')
    for dwg in root.iter(T('w:drawing')):
        ext = dwg.find(f'.//{{{WP}}}extent')
        if ext is not None:
            cx = int(ext.get('cx', '0')) / 360000
            cy = int(ext.get('cy', '0')) / 360000
            print(f'  图片: {cx:.1f}×{cy:.1f}cm')
            # 环绕
            for wtag in ['wrapSquare', 'wrapTight', 'wrapTopBottom', 'wrapThrough', 'wrapNone']:
                if dwg.find(f'.//{{{WP}}}{wtag}') is not None:
                    print(f'    环绕: {wtag}')
            # 对齐
            posH = dwg.find(f'.//{{{WP}}}positionH')
            if posH is not None:
                align = posH.find(f'{{{WP}}}align')
                if align is not None:
                    print(f'    水平对齐: {align.text}')
            # 旋转
            rot = dwg.find(f'.//{{{A_NS}}}xfrm')
            if rot is not None and rot.get('rot'):
                print(f'    旋转: {int(rot.get("rot"))/60000:.0f}°')
            # 效果
            dwg_s = ET.tostring(dwg, encoding='unicode')
            for eff in ['scene3d', 'sp3d', 'bevelT', 'shadow', 'glow', 'softEdge', 'reflection']:
                if eff in dwg_s: print(f'    效果: {eff}')

    # 7. 表格
    print(f'\n=== 表格 ===')
    tbls = list(root.iter(T('w:tbl')))
    if tbls:
        for ti, tbl in enumerate(tbls):
            rows = list(tbl.iter(T('w:tr')))
            print(f'  Table{ti+1}: {len(rows)}行')
            tblPr = tbl.find(T('w:tblPr'))
            if tblPr is not None:
                tblStyle = tblPr.find(T('w:tblStyle'))
                if tblStyle is not None:
                    print(f'    style: {GA(tblStyle, "w:val")}')
                tblW = tblPr.find(T('w:tblW'))
                if tblW is not None:
                    print(f'    width: type={GA(tblW, "w:type")} w={GA(tblW, "w:w")}')
                tblJc = tblPr.find(T('w:jc'))
                if tblJc is not None:
                    print(f'    align: {GA(tblJc, "w:val")}')
                tblBorders = tblPr.find(T('w:tblBorders'))
                if tblBorders is not None:
                    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                        b = tblBorders.find(T('w:' + side))
                        if b is not None:
                            print(f'    border.{side}: val={GA(b,"w:val")} sz={GA(b,"w:sz")} color={GA(b,"w:color")}')
            # 首行文本
            if rows:
                first_text = ' | '.join(GT(c).strip()[:15] for c in rows[0].iter(T('w:tc')) if GT(c).strip())
                print(f'    首行: {first_text[:120]}')

    # 8. 超链接
    print(f'\n=== 超链接 ===')
    for hl in root.iter(T('w:hyperlink')):
        anchor = hl.get(T('w:anchor'), '')
        rid = hl.get(T('r:id'), '')
        url = ''
        if 'word/_rels/document.xml.rels' in nl:
            rels_doc = zf.read('word/_rels/document.xml.rels').decode('utf-8')
            for rel in ET.fromstring(rels_doc).iter('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                if rel.get('Id', '') == rid:
                    url = rel.get('Target', '')
        text = GT(hl).strip()
        print(f'  "{text}" → {url}' if text else f'  anchor={anchor} rid={rid}')

    # 9. 样式
    if 'word/styles.xml' in nl:
        st = zf.read('word/styles.xml').decode('utf-8')
        print(f'\n=== 样式: {st.count("<w:style")} definitions ===')

    # 10. 水印
    print(f'\n=== 水印 ===')
    for hf in [f for f in nl if 'header' in f.lower()]:
        t = zf.read(hf).decode('utf-8')
        for shape in ET.fromstring(t).iter(T('v:shape')):
            sid = shape.get('id', '')
            if 'watermark' in sid.lower() or 'WaterMark' in sid:
                print(f'  {hf}: {sid}')
                stype = shape.get('type', '')
                txt = ET.tostring(shape, encoding='unicode')
                if 'imagedata' in txt:
                    m = re.search(r'blacklevel="([^"]+)"', txt)
                    mm = re.search(r'gain="([^"]+)"', txt)
                    print(f'    type={stype} blacklevel={m.group(1) if m else "?"} gain={mm.group(1) if mm else "?"}')
                textpath = shape.find(T('v:textpath'))
                if textpath is not None:
                    print(f'    文字水印: {textpath.get("string", "")}')


if __name__ == '__main__':
    from pathlib import Path
    n = sys.argv[1] if len(sys.argv) > 1 else '1'
    analyze(int(n))
