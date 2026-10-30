#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 2010 操作题 XML 评分器 — 通用模板
使用此模板创建新的 Excel 评分器: cp this.py excel_grader_N.py → 填充 req 函数 → 测试

核心流程:
  1. diff: 对比 素材.xlsx vs 答案.xlsx → 单元格差异 + 新增文件
  2. load: sharedStrings → styles → all_xml
  3. grade: 逐项检测 → score → print
"""
import zipfile, os, sys, io, re, shutil
from pathlib import Path
import xml.etree.ElementTree as ET

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

# ═══════════════════════════════════════════════════════════════
class ExcelGrader:
    """Excel XML 底层评分器"""
    def __init__(self, path):
        self.path = Path(path)
        self.zf = zipfile.ZipFile(path)
        self.namelist = self.zf.namelist()
        self.ss = []          # sharedStrings
        self.sheets = {}      # {name: {'root': Element, 'xml': str}}
        self.all_xml = ''     # 全量拼接 — 格式检测首选
        self.styles_xml = ''
        self.xf_list = []     # cellXfs 格式列表 [{numFmtId,fontId,fillId,borderId,...}]
        self.results = []

    # ──── 加载 ────
    def load(self):
        all_parts = []

        # 1. SharedStrings
        if 'xl/sharedStrings.xml' in self.namelist:
            d = self.zf.read('xl/sharedStrings.xml').decode('utf-8')
            all_parts.append(d)
            for si in ET.fromstring(d).iter(f'{{{S}}}si'):
                t = si.find(f'{{{S}}}t')
                if t is not None:
                    self.ss.append(t.text if t.text else '')
                else:
                    self.ss.append(''.join(tt.text or '' for tt in si.iter(f'{{{S}}}t')))

        # 2. Styles
        if 'xl/styles.xml' in self.namelist:
            d = self.zf.read('xl/styles.xml').decode('utf-8')
            all_parts.append(d)
            self.styles_xml = d
            root = ET.fromstring(d)
            self.xf_list = []
            for xf in root.iter(f'{{{S}}}xf'):
                self.xf_list.append({
                    'numFmtId': xf.get('numFmtId', '0'),
                    'fontId': xf.get('fontId', '0'),
                    'fillId': xf.get('fillId', '0'),
                    'borderId': xf.get('borderId', '0'),
                    'xfId': xf.get('xfId', '0'),
                })

        # 3. Workbook → sheet names + rels 映射
        d = self.zf.read('xl/workbook.xml').decode('utf-8')
        all_parts.append(d)
        wb = ET.fromstring(d)

        # rels 映射
        rel_map = {}
        if 'xl/_rels/workbook.xml.rels' in self.namelist:
            dr = self.zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
            for rel in ET.fromstring(dr).iter(f'{{{PKG}}}Relationship'):
                if 'worksheet' in rel.get('Type', ''):
                    rel_map[rel.get('Id', '')] = rel.get('Target', '')

        for sh in wb.iter(f'{{{S}}}sheet'):
            name = sh.get('name', '')
            r_id = sh.get(f'{{{R_NS}}}id', '')
            target = rel_map.get(r_id, f'worksheets/sheet{name}.xml')
            fname = 'xl/' + target
            if fname in self.namelist:
                d = self.zf.read(fname).decode('utf-8')
                all_parts.append(d)
                self.sheets[name] = {'root': ET.fromstring(d), 'xml': d, 'fname': fname}

        # 4. 兜底: 加载所有 sheet XML
        for f in self.namelist:
            if f.startswith('xl/worksheets/sheet') and f.endswith('.xml'):
                try:
                    d = self.zf.read(f).decode('utf-8')
                    all_parts.append(d)
                except: pass

        self.all_xml = '\n'.join(all_parts)

    # ──── 辅助函数 ────
    def sv(self, sheet_root, ref):
        """Get cell value via shared strings"""
        for row in sheet_root.iter(f'{{{S}}}row'):
            for c in row.iter(f'{{{S}}}c'):
                if c.get('r', '') == ref:
                    v = c.find(f'{{{S}}}v')
                    if v is None or v.text is None: return ''
                    if c.get('t', '') == 's':
                        idx = int(v.text)
                        return self.ss[idx] if idx < len(self.ss) else ''
                    return v.text
        return ''

    def get_cell(self, sheet_root, ref):
        """Get cell element"""
        for row in sheet_root.iter(f'{{{S}}}row'):
            for c in row.iter(f'{{{S}}}c'):
                if c.get('r', '') == ref: return c
        return None

    def get_formula(self, sheet_root, ref):
        """Get cell formula string"""
        c = self.get_cell(sheet_root, ref)
        if c is not None:
            f = c.find(f'{{{S}}}f')
            if f is not None and f.text: return f.text
        return ''

    # ──── 结果记录 ────
    def R(self, req, name, score, full, detail):
        self.results.append({'req': req, 'name': name, 'score': min(score, full),
                              'full': full, 'detail': detail})

    # ──── 评分 ────
    def grade(self):
        self.load()
        # ===== 填充你的 req 函数 =====
        self.req1()
        # self.req2()
        # ... 添加你的检测项

        total = sum(r['score'] for r in self.results)
        full = sum(r['full'] for r in self.results)
        print(f'\n{"="*56}\n  总得分: {total}/{full} = {total/full*100:.1f}%\n{"="*56}')

    # ===== 示例: 工作表+标题 =====
    def req1(self):
        d = []; s = 0
        # 改名检查
        if '员工工资表' in self.sheets:
            s += 2; d.append('改名 OK')
        else:
            d.append('NG 未改名')
        # 标签颜色
        if 'tabColor' in self.all_xml:
            s += 2; d.append('标签颜色 OK')
        # 标题文本
        s1 = self.sheets.get('员工工资表', {}).get('root')
        if s1 is not None:
            a1 = self.sv(s1, 'A1')
            if '工资' in a1:
                s += 2; d.append(f'A1标题: {a1}')
        self.R(1, '工作表+标题', s, 10, d)


# ═══════════════════════════════════════════════════════════════
if __name__ == '__main__':
    docx = sys.argv[1] if len(sys.argv) > 1 else None
    if docx is None:
        base = os.path.dirname(os.path.abspath(__file__))
        for f in Path(base).glob('1结果.xlsx'): docx = str(f); break
    if not docx:
        print('用法: python excel_grader_template.py <文件.xlsx>'); sys.exit(1)
    print(f'评分: {os.path.basename(docx)}')
    g = ExcelGrader(docx)
    g.grade()
