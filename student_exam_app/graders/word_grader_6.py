#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Word 2010 XML评分器 - 第6套题 v2 (总分100)"""
import sys, os, io, zipfile, datetime, shutil, re
from pathlib import Path
import xml.etree.ElementTree as ET
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import has_field, extract_word_parts

W  = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
V  = '{urn:schemas-microsoft-com:vml}'
R  = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
RELNS = '{http://schemas.openxmlformats.org/package/2006/relationships}'

def get_text(elem):
    if elem is None: return ""
    return ''.join(t.text or '' for t in elem.iter(W + 't'))

def load_rels(d):
    rd = {}
    for rel in ET.parse(str(d)).getroot().findall(RELNS + 'Relationship'):
        rd[rel.get('Id', '')] = {'target': rel.get('Target', ''), 'type': rel.get('Type', '')}
    return rd

class XMLGrader:
    def __init__(self, docx_path):
        self.docx = Path(docx_path); self.xml_dir = None; self.results = []
        self.doc_root = None; self.styles_root = None
        self.headers = {}; self.footers = {}; self.rels = {}; self.main_rels = {}
        self.doc_xml = ""

    def unzip(self):
        self.xml_dir = self.docx.parent / ('_xml_' + self.docx.stem)
        if self.xml_dir.exists(): shutil.rmtree(self.xml_dir)
        extract_word_parts(self.docx, self.xml_dir)

    def load_all(self):
        d = self.xml_dir / 'word'
        if (d / 'document.xml').exists():
            self.doc_root = ET.parse(str(d / 'document.xml')).getroot()
            self.doc_xml = ET.tostring(self.doc_root, encoding='unicode')
        if (d / 'styles.xml').exists(): self.styles_root = ET.parse(str(d / 'styles.xml')).getroot()
        for p in d.glob('header*.xml'):
            self.headers[p.name] = ET.parse(str(p)).getroot()
            self.doc_xml += ET.tostring(self.headers[p.name], encoding='unicode')
        for p in d.glob('footer*.xml'):
            self.footers[p.name] = ET.parse(str(p)).getroot()
            self.doc_xml += ET.tostring(self.footers[p.name], encoding='unicode')
        rels_dir = d / '_rels'
        if rels_dir.exists():
            for p in rels_dir.glob('*.xml.rels'):
                self.rels[p.stem.replace('.xml', '')] = load_rels(p)
        self.main_rels = self.rels.get('document', {})

    def _result(self, req, name, score, full, log):
        st = '通过' if score >= full * 0.75 else ('部分通过' if score >= full * 0.3 else '未通过')
        return {'req': req, 'name': name, 'score': min(score, full), 'full': full, 'detail': log, 'status': st}

    # req1: 标题10分
    def req1(self):
        log, score = [], 0
        paras = list(self.doc_root.iter(W + 'p'))
        if not paras: return self._result(1, '标题格式', 0, 10, log)
        title = paras[0]; txt = get_text(title); log.append('  标题: ' + txt[:40])
        if '灵魂砍价' in txt and '底价' in txt: score += 1; log.append('  OK 标题文字 (+1)')
        pPr = title.find(W + 'pPr')
        if pPr is not None:
            jc = pPr.find(W + 'jc')
            if jc is not None and jc.get(W + 'val', '') == 'center': score += 2; log.append('  OK 居中 (+2)')
            sp = pPr.find(W + 'spacing')
            if sp is not None and sp.get(W + 'beforeLines', '') == '50' and sp.get(W + 'afterLines', '') == '100':
                score += 1; log.append('  OK 段前0.5行段后1行 (+1)')
            rPr = pPr.find(W + 'rPr')
            if rPr is not None:
                rf = rPr.find(W + 'rFonts')
                if rf is not None and '行楷' in rf.get(W + 'eastAsia', ''): score += 2; log.append('  OK 华文行楷 (+2)')
                if rPr.find(W + 'b') is not None: score += 1; log.append('  OK 加粗 (+1)')
                sz = rPr.find(W + 'sz')
                if sz is not None and sz.get(W + 'val', '') == '44': score += 1; log.append('  OK 二号 (+1)')
                color = rPr.find(W + 'color')
                if color is not None and 'FF0000' in color.get(W + 'val', '').upper(): score += 1; log.append('  OK 红色 (+1)')
                bdr = rPr.find(W + 'bdr')
                if bdr is not None and bdr.get(W + 'sz', '') == '12' and '0070C0' in bdr.get(W + 'color', '').upper():
                    score += 1; log.append('  OK 蓝色字符边框 (+1)')
        if score >= 8: score = 10
        return self._result(1, '标题格式', score, 10, log)

    # req2: 字符边框+段间距10分
    def req2(self):
        log, score = [], 0
        paras = list(self.doc_root.iter(W + 'p'))
        if not paras: return self._result(2, '字符边框+段间距', 0, 10, log)
        pPr = paras[0].find(W + 'pPr')
        if pPr is not None:
            rPr = pPr.find(W + 'rPr')
            if rPr is not None:
                bdr = rPr.find(W + 'bdr')
                if bdr is not None:
                    if bdr.get(W + 'sz', '') == '12': score += 4; log.append('  OK 1.5磅边框 (+4)')
                    if '0070C0' in bdr.get(W + 'color', '').upper(): score += 3; log.append('  OK 蓝色 (+3)')
            sp = pPr.find(W + 'spacing')
            if sp is not None:
                if sp.get(W + 'beforeLines', '') == '50': score += 2; log.append('  OK 段前0.5行 (+2)')
                if sp.get(W + 'afterLines', '') == '100': score += 1; log.append('  OK 段后1行 (+1)')
        if score >= 8: score = 10
        return self._result(2, '字符边框+段间距', score, 10, log)

    # req3: 正文格式10分
    def req3(self):
        log, score = [], 0
        paras = list(self.doc_root.iter(W + 'p'))
        if len(paras) < 2: return self._result(3, '正文格式', 0, 10, log)
        pPr = paras[1].find(W + 'pPr')
        if pPr is not None:
            ind = pPr.find(W + 'ind')
            if ind is not None:
                flc = ind.get(W + 'firstLineChars', ind.get(W + 'firstLine', '0'))
                if flc and int(flc) >= 200: score += 4; log.append('  OK 首行缩进2字符 (+4)')
                elif flc and int(flc) > 0: score += 2; log.append('  OK 有缩进 (+2)')
            sp = pPr.find(W + 'spacing')
            if sp is not None:
                line, rule = sp.get(W + 'line', '0'), sp.get(W + 'lineRule', '')
                if rule == 'exact' and line and int(line) == 280: score += 6; log.append('  OK 固定值14磅 (+6)')
                elif rule == 'exact': score += 4; log.append('  OK 固定行距 (+4)')
        return self._result(3, '正文格式', score, 10, log)

    # req4: 图片+题注10分
    def req4(self):
        log, score = [], 0
        img_found = False
        for dwg in self.doc_root.iter(W + 'drawing'):
            ext = dwg.find('.//' + WP + 'extent')
            if ext is not None:
                cx = int(ext.get('cx', 0)) / 914400 * 2.54
                cy = int(ext.get('cy', 0)) / 914400 * 2.54
                log.append('  图片: %.1fcm x %.1fcm' % (cx, cy))
                if abs(cx - 11) < 0.2 and abs(cy - 7) < 0.2: score += 4; img_found = True; log.append('  OK 尺寸11×7cm (+4)')
                elif abs(cx - 7) < 0.2 and abs(cy - 11) < 0.2: score += 3; img_found = True; log.append('  OK 尺寸7×11cm (+3)')
                if dwg.find('.//' + WP + 'wrapTopAndBottom') is not None: score += 3; log.append('  OK 上下型环绕 (+3)')
                elif dwg.find('.//' + WP + 'wrapTight') is not None: score += 3; log.append('  OK 紧密型环绕 (+3)')
                elif dwg.find('.//' + WP + 'wrapSquare') is not None: score += 2; log.append('  OK 四周型环绕 (+2)')
                posH = dwg.find('.//' + WP + 'positionH/' + WP + 'align')
                if posH is not None and posH.text == 'center': score += 1; log.append('  OK 水平居中 (+1)')
            break
        if not img_found: log.append('  NG 未找到图片')
        # 题注
        for p in self.doc_root.iter(W + 'p'):
            if '2021' in get_text(p) and '医保' in get_text(p):
                jc = p.find(W + 'pPr/' + W + 'jc')
                if jc is not None and jc.get(W + 'val', '') == 'center':
                    score += 2; log.append('  OK 图片题注居中 (+2)')
                else:
                    log.append('  NG 题注未居中')
                break
        if score >= 8: score = 10
        return self._result(4, '图片设置', score, 10, log)

    # req5: 脚注5分
    def req5(self):
        log, score = [], 0
        paras = list(self.doc_root.iter(W + 'p'))
        if len(paras) > 1:
            p2 = paras[1]
            if p2.find('.//' + W + 'footnoteReference') is not None: score += 2; log.append('  OK 脚注引用 (+2)')
            if '70万' in get_text(p2): score += 1; log.append('  OK 双引号内容 (+1)')
        fn_path = self.xml_dir / 'word' / 'footnotes.xml'
        if fn_path.exists():
            for fn in ET.parse(str(fn_path)).getroot().findall(W + 'footnote'):
                if get_text(fn).strip() == 'https://www.sina.com.cn/':
                    score += 2; log.append('  OK 脚注URL=sina.com.cn (+2)'); break
        return self._result(5, '双引号脚注', score, 5, log)

    # req6: 水印10分
    def req6(self):
        log, score = [], 0
        found = False
        for nm, root in self.headers.items():
            xs = ET.tostring(root, encoding='unicode')
            if '人民日报' in xs: found = True; break
        if not found and '人民日报' in self.doc_xml: found = True
        if found: score += 6; log.append('  OK 水印"人民日报" (+6)')
        else: return self._result(6, '文字水印', 0, 10, log)
        if '华文新魏' in self.doc_xml: score += 2; log.append('  OK 字体 (+2)')
        rot_found = any('rotation' in ET.tostring(r, encoding='unicode').lower() for r in self.headers.values())
        if rot_found or 'rotation' in self.doc_xml.lower(): score += 2; log.append('  OK 斜式 (+2)')
        if score >= 8: score = 10
        return self._result(6, '文字水印', score, 10, log)

    # req7: 页面边框10分
    def req7(self):
        log, score = [], 0
        for sp in self.doc_root.iter(W + 'sectPr'):
            pgBorders = sp.find(W + 'pgBorders')
            if pgBorders is not None:
                bx = ET.tostring(pgBorders, encoding='unicode')
                if 'double' in bx: score += 4; log.append('  OK 双实线 (+4)')
                if any(s == '6' for s in re.findall(r'sz="(\d+)"', bx)): score += 3; log.append('  OK 0.75磅 (+3)')
                if '0070C0' in bx: score += 3; log.append('  OK 蓝色 (+3)')
            break
        if score >= 8: score = 10
        return self._result(7, '页面颜色边框', score, 10, log)

    # req8: 页眉页脚10分
    def req8(self):
        log, score = [], 0
        for rid, info in self.main_rels.items():
            if 'header' in info.get('type', ''):
                root = self.headers.get(info['target'])
                if root is None: continue
                txt = get_text(root)
                if '北京日报' in txt: score += 4; log.append('  OK 页眉"北京日报" (+4)')
                elif '国家力量' in txt: score += 3; log.append('  OK 页眉"国家力量" (+3)')
        for nm, root in self.footers.items():
            if has_field(ET.tostring(root, encoding='unicode'), 'PAGE'): score += 3; log.append('  OK 页脚含页码域 (+3)'); break
        if score >= 8: score = 10
        return self._result(8, '页眉页脚', score, 10, log)

    # req9: 表格格式10分
    def req9(self):
        log, score = [], 0
        tables = self.doc_root.findall('.//' + W + 'tbl')
        if not tables: return self._result(9, '表格格式', 0, 10, log)
        tbl = tables[0]; tblPr = tbl.find(W + 'tblPr')
        rows = tbl.findall(W + 'tr'); cols = len(rows[0].findall(W + 'tc')) if rows else 0
        log.append('  表格: %dx%d' % (len(rows), cols))
        if len(rows) >= 8: score += 4; log.append('  OK 8行 (+4)')
        if cols >= 2: score += 2; log.append('  OK 2列 (+2)')
        if tblPr is not None:
            tblW = tblPr.find(W + 'tblW')
            if tblW is not None and tblW.get(W + 'type', '') in ('auto', 'pct'): score += 2; log.append('  OK 自动调整 (+2)')
            jc = tblPr.find(W + 'jc')
            if jc is not None and jc.get(W + 'val', '') == 'center': score += 2; log.append('  OK 表格居中 (+2)')
        if score >= 8: score = 10
        return self._result(9, '表格格式', score, 10, log)

    # req10: 表格边框10分
    def req10(self):
        log, score = [], 0
        tables = self.doc_root.findall('.//' + W + 'tbl')
        if not tables: return self._result(10, '表格边框', 0, 10, log)
        tblPr = tables[0].find(W + 'tblPr')
        if tblPr is None: return self._result(10, '表格边框', 0, 10, log)
        tblBorders = tblPr.find(W + 'tblBorders')
        if tblBorders is not None:
            bx = ET.tostring(tblBorders, encoding='unicode')
            if 'double' in bx: score += 3; log.append('  OK 外双实线 (+3)')
            if 'FF0000' in bx: score += 2; log.append('  OK 红色外框 (+2)')
            if any(s == '12' for s in re.findall(r'sz="(\d+)"', bx)): score += 2; log.append('  OK 1.5磅外框 (+2)')
            if 'single' in bx: score += 2; log.append('  OK 内单实线 (+2)')
            if '002060' in bx: score += 1; log.append('  OK 蓝色内框 (+1)')
        if score >= 8: score = 10
        return self._result(10, '表格边框', score, 10, log)

    # req11: 超链接5分
    def req11(self):
        log, score = [], 0
        for hl in self.doc_root.findall('.//' + W + 'hyperlink'):
            rid = hl.get(R + 'id', '')
            if rid in self.main_rels:
                t = self.main_rels[rid].get('target', '')
                if t == 'http://www.gov.cn/' or t == 'http://www.gov.cn':
                    score = 5; log.append('  OK 超链接=http://www.gov.cn/ (+5)'); break
                elif 'gov.cn' in t:
                    score = 5; log.append('  OK 超链接gov.cn (+5)'); break
        if score == 0:
            if 'gov.cn' in self.doc_xml: score = 3; log.append('  OK 找到gov.cn (+3)')
        if score == 0: log.append('  NG 未找到超链接')
        return self._result(11, '超链接', score, 5, log)

    # 评分主流程
    def grade(self):
        print('\n' + '=' * 60)
        print('  Word 2010 XML评分器 - 第6套题 v2 (总分100)')
        print('=' * 60)
        self.unzip(); self.load_all()
        print('\n' + '-' * 60 + '\n  逐项检测...\n' + '-' * 60)
        checks = [self.req1, self.req2, self.req3, self.req4, self.req5,
                  self.req6, self.req7, self.req8, self.req9, self.req10, self.req11]
        names = ['标题格式', '字符边框+段间距', '正文格式', '图片设置',
                 '双引号脚注', '文字水印', '页面颜色边框', '页眉页脚',
                 '表格格式', '表格边框', '超链接']
        for i, chk in enumerate(checks):
            r = chk(); self.results.append(r)
            icon = {'通过': '[OK]', '部分通过': '[--]', '未通过': '[NG]'}.get(r['status'], '?')
            print('\n[%02d] %s' % (i+1, names[i]))
            print('  %s %d/%d (%s)' % (icon, r['score'], r['full'], r['status']))
            for d in r['detail']: print('    ' + d)
        total = sum(r['score'] for r in self.results); full = sum(r['full'] for r in self.results)
        rate = total / full * 100 if full > 0 else 0
        print('\n' + '=' * 60); print('  总得分: %d/%d = %.1f%%' % (total, full, rate)); print('=' * 60)
        return self.results

    def report(self, path=None):
        if not self.results: return
        path = path or (self.docx.parent / 'report_6.html')
        total = sum(r['score'] for r in self.results); full = sum(r['full'] for r in self.results)
        rate = total / full * 100 if full > 0 else 0
        gc = '#28a745' if rate >= 90 else ('#17a2b8' if rate >= 75 else ('#ffc107' if rate >= 60 else '#dc3545'))
        gr = 'A 优秀' if rate >= 90 else ('B 良好' if rate >= 75 else ('C 及格' if rate >= 60 else 'D 不及格'))
        h = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Word评分报告 - 第6套题</title>'
        h += '<style>*{{margin:0;padding:0;box-sizing:border-box}}body{{font-family:Microsoft YaHei,sans-serif;background:linear-gradient(135deg,#667eea,#764ba2);min-height:100vh;padding:20px}}.box{{max-width:1000px;margin:0 auto;background:white;border-radius:20px;box-shadow:0 20px 60px rgba(0,0,0,0.3);overflow:hidden}}.hd{{background:linear-gradient(135deg,#667eea,#764ba2);color:white;padding:30px;text-align:center}}h1{{font-size:26px;margin-bottom:10px}}.sc{{display:flex;align-items:center;justify-content:center;padding:30px;gap:50px;background:#f8f9fa}}.num{{font-size:56px;font-weight:bold;color:' + gc + '}}.gb{{background:' + gc + ';color:white;padding:8px 20px;border-radius:20px;font-weight:bold;display:inline-block}}.c{{background:#f8f9fa;border-radius:12px;padding:16px;margin:12px 30px;border-left:5px solid #dee2e6}}.c.p1{{border-left-color:#28a745}}.c.p2{{border-left-color:#ffc107}}.c.p3{{border-left-color:#dc3545}}.ch{{display:flex;justify-content:space-between;margin-bottom:8px}}.ct{{font-size:15px;font-weight:bold}}.bd{{padding:3px 10px;border-radius:15px;font-size:12px;font-weight:bold}}.b1{{background:#d4edda;color:#155724}}.b2{{background:#fff3cd;color:#856404}}.b3{{background:#f8d7da;color:#721c24}}.d{{background:white;padding:10px;border-radius:8px;margin-top:8px;font-size:12px;line-height:1.8;max-height:200px;overflow-y:auto}}.ft{{text-align:center;padding:15px;background:#f8f9fa;color:#6c757d;font-size:12px}}</style></head><body>'
        h += '<div class="box"><div class="hd"><h1>Word 2010 操作题评分报告（第6套）</h1><p>基于XML底层解析 | 共11项评分要求 | v2</p></div>'
        h += '<div class="sc"><div style="text-align:center;"><div class="num">' + str(total) + '</div><div style="color:#6c757d;font-size:16px;">/ ' + str(full) + ' 分</div><div style="font-size:40px;color:' + gc + ';">' + ('%.1f' % rate) + '%</div><div class="gb">' + gr + '</div></div></div>'
        for r in self.results:
            sc = {'通过': 'p1', '部分通过': 'p2', '未通过': 'p3'}.get(r['status'], 'p3')
            bd = {'通过': 'b1', '部分通过': 'b2', '未通过': 'b3'}.get(r['status'], 'b3')
            h += '<div class="c ' + sc + '"><div class="ch">'
            h += '<span class="ct">要求%02d %s (%d/%d)</span>' % (r['req'], r['name'], r['score'], r['full'])
            h += '<span class="bd ' + bd + '">' + r['status'] + '</span></div><div class="d">'
            for d in r['detail']: h += '<div>' + d + '</div>'
            h += '</div></div>'
        h += '<div class="ft"><p>基于XML底层解析 | 精准对照11项要求</p><p>生成: ' + datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + '</p></div></div></body></html>'
        with open(path, 'w', encoding='utf-8') as f: f.write(h)
        print('\n报告已生成: ' + str(path))

def main():
    docx = None
    if len(sys.argv) > 1: docx = sys.argv[1]
    else:
        for f in Path('.').glob('*结果*.docx'): docx = str(f); break
        if not docx:
            for f in Path('.').glob('*.docx'): docx = str(f); break
    if not docx: print('未找到.docx文件！'); return
    print('目标文件: ' + docx)
    g = XMLGrader(docx); g.grade(); g.report()

if __name__ == '__main__': main()
