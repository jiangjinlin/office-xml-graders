#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT 2010 操作题 分析器
对比 素材.pptx 与 答案.pptx，快速找出差异

用法: python ppt_analyze.py N
  N = 文件夹编号
"""
import zipfile, xml.etree.ElementTree as ET, sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

def analyze(folder_num):
    base = Path(os.path.dirname(__file__))
    mat_dir = base.parent.parent.parent / 'PPT 2010 操作题题库' / str(folder_num)
    ans_path = mat_dir / '1结果.pptx'
    src_path = mat_dir / '2素材.pptx'

    if not ans_path.exists():
        print(f'错误: 找不到 {ans_path}'); return

    ans_zf = zipfile.ZipFile(str(ans_path))
    ans_nl = ans_zf.namelist()
    src_zf = zipfile.ZipFile(str(src_path)) if src_path.exists() else None
    src_nl = set(src_zf.namelist()) if src_zf else set()

    # 1. 文件差异
    new_files = sorted(set(ans_nl) - src_nl)
    print(f'新增文件: {new_files}')
    ans_slides = sorted(f for f in ans_nl if f.startswith('ppt/slides/slide') and f.endswith('.xml') and '_rels' not in f)
    src_slides = sorted(f for f in src_nl if f.startswith('ppt/slides/slide') and f.endswith('.xml') and '_rels' not in f)
    print(f'幻灯片: 答案{len(ans_slides)} 素材{len(src_slides)}')

    # 2. 主题
    if 'ppt/theme/theme1.xml' in ans_nl:
        t = ans_zf.read('ppt/theme/theme1.xml').decode('utf-8')
        tn = ET.fromstring(t).get('name', '')
        print(f'\n主题: {tn}')

    # 3. 每张幻灯片详情
    for sld_name in ans_slides:
        sld = ans_zf.read(sld_name).decode('utf-8')
        root = ET.fromstring(sld)
        texts = re.findall(r'<a:t[^>]*>([^<]+)</a:t>', sld)
        sizes = re.findall(r'sz="(\d+)"', sld)
        fonts = re.findall(r'typeface="([^"]+)"', sld)

        # 版式
        name = sld_name.split('/')[-1]
        rf = f'ppt/slides/_rels/{name}.rels'
        layout = ''
        if rf in ans_nl:
            for rel in ET.fromstring(ans_zf.read(rf).decode('utf-8')).iter(f'{{{PKG}}}Relationship'):
                if 'slideLayout' in rel.get('Type', ''):
                    layout = rel.get('Target', '').split('/')[-1].replace('.xml', '')
        print(f'\n=== {sld_name} layout={layout} ===')
        if texts: print(f'  texts: {" | ".join(t[:40] for t in texts[:5])}')
        if sizes: print(f'  sizes: {sizes}')
        if fonts: print(f'  fonts: {fonts[:5]}')

        # 切换
        trans = root.find(f'{{{P}}}transition')
        if trans is not None:
            for child in list(trans):
                tag = child.tag.split('}')[-1]
                if tag != 'sndAc':
                    spd = trans.get('spd', '')
                    print(f'  transition: {tag} speed={spd}')

        # 动画
        if 'animEffect' in sld:
            print(f'  anim: animEffect')
        if 'afterEffect' in sld:
            print(f'  timing: afterEffect')

        # 图片
        if rf in ans_nl:
            for rel in ET.fromstring(ans_zf.read(rf).decode('utf-8')).iter(f'{{{PKG}}}Relationship'):
                if 'image' in rel.get('Type', ''):
                    print(f'  image: {rel.get("Target","")}')

        # 页脚
        ftr_m = re.search(r'ph type="ftr".*?<a:t[^>]*>([^<]+)</a:t>', sld, re.DOTALL)
        if ftr_m: print(f'  footer: {ftr_m.group(1)}')

    # 4. 放映设置
    pp = ans_zf.read('ppt/presProps.xml').decode('utf-8') if 'ppt/presProps.xml' in ans_nl else ''
    for tag in ['browse', 'kiosk', 'present']:
        if f'<p:{tag}' in pp: print(f'\n放映: {tag}')

    # 5. 备注
    notes_list = sorted(f for f in ans_nl if 'notesSlides' in f)
    if notes_list:
        print(f'\n备注: {len(notes_list)} 个')


if __name__ == '__main__':
    from pathlib import Path
    n = sys.argv[1] if len(sys.argv) > 1 else '1'
    analyze(int(n))
