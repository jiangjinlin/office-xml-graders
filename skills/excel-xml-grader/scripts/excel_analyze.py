#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 2010 操作题 分析器
对比 素材.xlsx 与 答案.xlsx，找出所有差异，辅助生成评分器代码

用法: python excel_analyze.py N
  N = 文件夹编号 (1-10)
  会自动读取 Mater/N/2素材.xlsx 和 Mater/N/1结果.xlsx
"""
import zipfile, xml.etree.ElementTree as ET, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def analyze(folder_num):
    base = Path(os.path.dirname(__file__))
    mat_dir = base.parent.parent.parent / '2022年准易：Excel 2010 操作题--素材' / str(folder_num)
    src_path = mat_dir / '2素材.xlsx'
    ans_path = mat_dir / '1结果.xlsx'

    if not ans_path.exists():
        print(f'错误: 找不到 {ans_path}'); return

    src_zf = zipfile.ZipFile(str(src_path)) if src_path.exists() else None
    ans_zf = zipfile.ZipFile(str(ans_path))
    src_files = set(src_zf.namelist()) if src_zf else set()
    ans_files = set(ans_zf.namelist())

    # 1. 文件差异
    new_files = sorted(ans_files - src_files)
    removed_files = sorted(src_files - ans_files) if src_zf else []
    print(f'新增文件: {new_files}')
    print(f'移除文件: {removed_files}')

    # 2. SharedStrings 差异
    print('\n=== SharedStrings 差异 ===')
    ans_ss = []; src_ss = []
    if 'xl/sharedStrings.xml' in ans_files:
        for si in ET.fromstring(ans_zf.read('xl/sharedStrings.xml').decode('utf-8')).iter(f'{{{S}}}si'):
            ans_ss.append(''.join(t.text or '' for t in si.iter(f'{{{S}}}t')))
    if src_zf and 'xl/sharedStrings.xml' in src_files:
        for si in ET.fromstring(src_zf.read('xl/sharedStrings.xml').decode('utf-8')).iter(f'{{{S}}}si'):
            src_ss.append(''.join(t.text or '' for t in si.iter(f'{{{S}}}t')))
    new_texts = [t for t in ans_ss if t not in src_ss]
    if new_texts: print(f'新增文本 ({len(new_texts)}): {new_texts[:30]}')

    # 3. Styles 差异
    if 'xl/styles.xml' in ans_files:
        print('\n=== Styles 概览 ===')
        d = ans_zf.read('xl/styles.xml').decode('utf-8')
        root = ET.fromstring(d)
        print(f'  numFmts: {len(list(root.iter(f"{{{S}}}numFmt")))}')
        print(f'  fonts: {len(list(root.iter(f"{{{S}}}font")))}')
        print(f'  fills: {len(list(root.iter(f"{{{S}}}fill")))}')
        print(f'  borders: {len(list(root.iter(f"{{{S}}}border")))}')
        print(f'  cellXfs: {len(list(root.iter(f"{{{S}}}xf")))}')
        print(f'  dxfs: {len(list(root.iter(f"{{{S}}}dxf")))}')

    # 4. 工作表概况
    print('\n=== 工作表 ===')
    d = ans_zf.read('xl/workbook.xml').decode('utf-8')
    for sh in ET.fromstring(d).iter(f'{{{S}}}sheet'):
        name = sh.get('name', '')
        sid = sh.get('sheetId', '')
        print(f'  [{sid}] {name}')

    # 5. 特殊元素检测
    ans_all = b'\n'.join(ans_zf.read(f) for f in ans_files if f.endswith('.xml')).decode('utf-8', errors='replace')
    for kw in ['conditionalFormatting', 'dataBar', 'mergeCell', 'tabColor',
                'sortState', 'sparkline', 'pivotTable', 'tableStyleInfo',
                'chart', 'drawing', 'pageSetup', 'headerFooter',
                'SUBTOTAL', 'VLOOKUP', 'SUMIF', 'RANK', 'IF']:
        count = ans_all.count(kw)
        if count > 0: print(f'  {kw}: {count} occurrences')


if __name__ == '__main__':
    from pathlib import Path
    n = sys.argv[1] if len(sys.argv) > 1 else '1'
    analyze(int(n))
