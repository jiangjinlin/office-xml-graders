"""
Word 2010 操作题评分器 - 文件夹8 (v2 — 精确值修正)
"""
import zipfile, os, re, io, sys
import xml.etree.ElementTree as ET
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import has_field

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
V = '{urn:schemas-microsoft-com:vml}'
O = '{urn:schemas-microsoft-com:office:office}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'

NS_URI = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'v': 'urn:schemas-microsoft-com:vml',
    'o': 'urn:schemas-microsoft-com:office:office',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
}

def child(e,t):
    for c in e:
        if c.tag==t: return c
    return None

def children(e,t):
    return [c for c in e if c.tag==t]

def ga(prefix, local):
    return '{' + NS_URI.get(prefix,'') + '}' + local

def GA(elem, ns_local, default=None):
    if elem is None: return default
    if ':' in ns_local:
        ns,l = ns_local.split(':',1)
        return elem.get(ga(ns,l), default)
    return elem.get(ns_local, default)

def get_text(elem):
    return ''.join(t.text or '' for t in elem.iter(W+'t'))

class Grader8:
    def __init__(self,path):
        self.path=path
        self.body=None
        self.doc_root=None
        self.doc_xml=''
        self.all_xml_str=''
        self.headers={}
        self.rel={}
        self.footnotes_root=None
        self.styles_root=None
    
    def load(self):
        with zipfile.ZipFile(self.path) as z:
            self.doc_xml=z.read('word/document.xml').decode('utf-8')
            self.doc_root=ET.fromstring(self.doc_xml.encode('utf-8'))
            for c in self.doc_root:
                if c.tag==W+'body': self.body=c; break
            all_x=[self.doc_xml]
            self.all_parts={}
            for n in z.namelist():
                if n.endswith('.xml'):
                    try:
                        c=z.read(n).decode('utf-8',errors='replace')
                        all_x.append(c)
                        self.all_parts[n]=c
                        if 'header' in n: self.headers[n]=c
                    except: pass
            self.all_xml_str='\n'.join(all_x)
            try:
                rels=ET.fromstring(z.read('word/_rels/document.xml.rels'))
                for rel in rels:
                    self.rel[rel.get('Id','')]={'t':rel.get('Target',''),'y':rel.get('Type','')}
            except: pass
            try: self.styles_root=ET.fromstring(z.read('word/styles.xml'))
            except: pass
            try: self.footnotes_root=ET.fromstring(z.read('word/footnotes.xml'))
            except: pass
    
    def grade(self):
        self.load()
        rr=[]
        rr.append(self.r1()); rr.append(self.r2()); rr.append(self.r3())
        rr.append(self.r4()); rr.append(self.r5()); rr.append(self.r6())
        rr.append(self.r7()); rr.append(self.r8()); rr.append(self.r9())
        rr.append(self.r10()); rr.append(self.r11())
        t=round(sum(x['s'] for x in rr),1)
        rr.append({'req':0,'name':'总分','s':t,'full':100,'d':[],'st':'pass' if t>=60 else 'fail'})
        return rr
    
    def R(self,req,name,score,full,detail):
        score=min(score,full)
        st='pass' if score>=full*0.6 else 'fail'
        return {'req':req,'name':name,'s':round(score,1),'full':full,'d':detail,'st':st}
    
    # ====== R1: 标题 (9) ======
    def r1(self):
        d=[]; s=0; f=9
        paras=children(self.body,W+'p')
        if not paras: return self.R(1,'标题',s,f,d)
        pPr=child(paras[0],W+'pPr')
        rPr=child(pPr,W+'rPr') if pPr is not None else None
        if rPr is None: return self.R(1,'标题',s,f,d)
        
        rf=child(rPr,W+'rFonts')
        szE=child(rPr,W+'sz')
        bE=child(rPr,W+'b')
        iE=child(rPr,W+'i')
        cE=child(rPr,W+'color')
        wE=child(rPr,W+'w')
        
        if rf is not None and GA(rf,'w:eastAsia','')=='华文行楷': s+=2; d.append('华文行楷 OK')
        if bE is not None: s+=1.5; d.append('加粗 OK')
        if iE is not None: s+=1.5; d.append('倾斜 OK')
        if szE is not None and GA(szE,'w:val','')=='44': s+=1.5; d.append('二号(44) OK')
        if cE is not None and GA(cE,'w:val','')=='7030A0': s+=1.5; d.append('紫色(7030A0) OK')
        if wE is not None and GA(wE,'w:val','')=='150': s+=1; d.append('字符缩放150% OK')
        return self.R(1,'标题',s,f,d)
    
    # ====== R2: 标题边框间距 (9) ======
    def r2(self):
        d=[]; s=0; f=9
        paras=children(self.body,W+'p')
        if not paras: return self.R(2,'标题边框间距',s,f,d)
        pPr=child(paras[0],W+'pPr')
        if pPr is None: return self.R(2,'标题边框间距',s,f,d)
        
        # 边框(w:rPr/w:bdr)
        rPr=child(pPr,W+'rPr')
        if rPr is not None:
            bdr=child(rPr,W+'bdr')
            if bdr is not None and GA(bdr,'w:val','')=='single' and GA(bdr,'w:sz','')=='12' and GA(bdr,'w:color','')=='0070C0':
                s+=3; d.append('1.5磅蓝色单实线方框 OK')
        
        # 间距
        spc=child(pPr,W+'spacing')
        if spc is not None:
            bl=GA(spc,'w:beforeLines','')
            al=GA(spc,'w:afterLines','')
            if bl=='50': s+=2; d.append('段前0.5行 OK')
            if al=='100': s+=2; d.append('段后1行 OK')
        
        # 居中
        jc=child(pPr,W+'jc')
        if jc is not None and GA(jc,'w:val','')=='center': s+=2; d.append('居中 OK')
        return self.R(2,'标题边框间距',s,f,d)
    
    # ====== R3: 正文格式 (10) ======
    def r3(self):
        d=[]; s=0; f=10
        paras=children(self.body,W+'p')
        body_count=0; font_ok=0; sz_ok=0; ind_ok=0; line_ok=0
        for p in paras:
            txt=get_text(p)
            if not txt or '遇上碎片化' in txt or '表格' in txt: continue
            body_count+=1
            pPr=child(p,W+'pPr')
            if pPr is None: continue
            
            # 字体字号—先查pPr/rPr, 再查r/rPr, 再查样式
            rPr=child(pPr,W+'rPr')
            if rPr is not None and child(rPr,W+'rFonts') is not None:
                rf=child(rPr,W+'rFonts')
                szE=child(rPr,W+'sz')
                if GA(rf,'w:eastAsia','')=='华文新魏': font_ok+=1
                if szE is not None and GA(szE,'w:val','')=='28': sz_ok+=1
            else:
                # Check run level
                for r in children(p,W+'r'):
                    rrPr=child(r,W+'rPr')
                    if rrPr is not None:
                        rf=child(rrPr,W+'rFonts')
                        if rf is not None:
                            if GA(rf,'w:eastAsia','')=='华文新魏': font_ok+=1; break
            
            # 缩进
            ind=child(pPr,W+'ind')
            if ind is not None:
                flc=GA(ind,'w:firstLineChars','0')
                if int(flc)>=150: ind_ok+=1
            
            # 行距
            spc=child(pPr,W+'spacing')
            if spc is not None and GA(spc,'w:line','0')=='360': line_ok+=1
        
        if body_count==0: return self.R(3,'正文格式',s,f,d)
        if font_ok/body_count>=0.8: s+=3; d.append(f'华文新魏 ({font_ok}/{body_count}) OK')
        elif font_ok>0: s+=1.5; d.append(f'华文新魏 部分({font_ok}/{body_count})')
        if sz_ok/body_count>=0.5: s+=2; d.append(f'四号 ({sz_ok}/{body_count}) OK')
        if ind_ok/body_count>=0.7: s+=2.5; d.append(f'首行缩进 ({ind_ok}/{body_count}) OK')
        if line_ok/body_count>=0.7: s+=2.5; d.append(f'1.5倍行距 ({line_ok}/{body_count}) OK')
        return self.R(3,'正文格式',s,f,d)
    
    # ====== R4: 图片 (10) ======
    def r4(self):
        d=[]; s=0; f=10
        for anchor in self.doc_root.iter(WP+'anchor'):
            ext=child(anchor,WP+'extent')
            if ext is None: continue
            cx=int(ext.get('cx','0'))/914400*2.54
            cy=int(ext.get('cy','0'))/914400*2.54
            # 图片总分调整
            if abs(cx-5.0)<0.5 and abs(cy-4.0)<0.5: s+=2.5; d.append(f'尺寸{cx:.1f}x{cy:.1f}cm OK')
            if len(children(anchor,WP+'wrapTight'))>0: s+=2; d.append('紧密型 OK')
            for ph in anchor.iter(WP+'positionH'):
                rf=ph.get('relativeFrom','')
                al=child(ph,WP+'align')
                if rf=='column' and al is not None: s+=1.5; d.append('栏居中 OK')
            if 'bevel' in self.doc_xml.lower(): s+=1.5; d.append('棱台 OK')
            if 'pencil' in self.all_xml_str.lower(): s+=1.5; d.append('铅笔灰度 OK')
            found_tp = any('tp.jpg' in rel['t'] for rel in self.rel.values())
            if found_tp: s+=1.5; d.append('tp.jpg引用 OK')
            elif any('image' in rel['y'] for rel in self.rel.values()): s+=1; d.append('有图片引用')
            return self.R(4,'插入图片',s,f,d)
        d.append('未找到图片')
        return self.R(4,'插入图片',s,f,d)
    
    # ====== R5: 脚注 (8) ======
    def r5(self):
        d=[]; s=0; f=8
        fn_refs=list(self.doc_root.iter(W+'footnoteReference'))
        if fn_refs: s+=2; d.append(f'{len(fn_refs)}个脚注引用')
        else: d.append('无脚注引用')
        target='碎片化的形式有很多，但是它有利有弊'
        if self.footnotes_root:
            found=False
            for fn in self.footnotes_root.iter(W+'footnote'):
                if target in get_text(fn):
                    s+=3; d.append('脚注内容 OK'); found=True; break
            if not found: d.append('脚注内容未找到')
        elif target in self.all_xml_str: s+=3; d.append('字符串搜索找到脚注')
        
        # 碎片化加粗
        paras=children(self.body,W+'p')
        if len(paras)>1:
            for r in children(paras[1],W+'r'):
                rt=''.join(t.text or '' for t in children(r,W+'t'))
                if '碎片化' in rt:
                    rPr=child(r,W+'rPr')
                    if rPr is not None and child(rPr,W+'b') is not None: s+=1; d.append('碎片化加粗 OK')
                    # Check red color
                    cE=child(rPr,W+'color') if rPr is not None else None
                    if cE is not None and GA(cE,'w:val','')=='FF0000': s+=2; d.append('碎片化红色 OK')
                    break
        return self.R(5,'脚注',s,f,d)
    
    # ====== R6: 云形标注 (9) ======
    def r6(self):
        d=[]; s=0; f=9
        for pict in self.doc_root.iter(W+'pict'):
            for sh in pict.iter(V+'shape'):
                if sh.get('type','')=='#_x0000_t106':
                    s+=3; d.append('云形标注 OK')
                    for tb in sh.iter(V+'textbox'):
                        if get_text(tb).strip()=='合理利用电子产品': s+=3; d.append('文字 OK')
                    for fl in sh.iter(V+'fill'):
                        if fl.get('type','')=='gradient': s+=2; d.append('渐变填充 OK'); break
                    if 'tight' in sh.get('style','').lower() or 'square' in sh.get('style','').lower(): s+=1; d.append('环绕 OK')
                    return self.R(6,'云形标注',s,f,d)
        return self.R(6,'云形标注',s,f,d)
    
    # ====== R7: 分栏 (9) ======
    def r7(self):
        d=[]; s=0; f=9
        paras=children(self.body,W+'p')
        last_p=None
        for p in reversed(paras):
            txt=get_text(p).strip()
            if txt and '表格' not in txt: last_p=p; break
        if last_p is not None:
            pPr=child(last_p,W+'pPr')
            if pPr is not None:
                ind=child(pPr,W+'ind')
                if ind is not None:
                    if int(GA(ind,'w:leftChars','0'))>=150: s+=1.5; d.append('左缩进2字符 OK')
                    if int(GA(ind,'w:rightChars','0'))>=150: s+=1.5; d.append('右缩进2字符 OK')
        
        found=False
        for sp in self.doc_root.iter(W+'sectPr'):
            cols=child(sp,W+'cols')
            if cols is not None and GA(cols,'w:num','1')=='2':
                found=True; s+=2; d.append('两栏 OK')
                if GA(cols,'w:sep','')=='1': s+=1.5; d.append('分隔线 OK')
                clist=children(cols,W+'col')
                if clist and abs(int(GA(clist[0],'w:w','0'))-2940)<200: s+=1.5; d.append('第1栏宽14字符 OK')
                if GA(cols,'w:space','0')=='210': s+=1; d.append('栏间距1字符 OK')
                break
        if not found: d.append('未找到2栏设置(sectPr存在但cols.num!=2?)')
        return self.R(7,'分栏',s,f,d)
    
    # ====== R8: 页眉页脚 (9) ======
    def r8(self):
        d=[]; s=0; f=9
        seen_center=set()
        for name,hxml in self.headers.items():
            hroot=ET.fromstring(hxml)
            htxt=''.join(t.text or '' for t in hroot.iter(W+'t')).strip()
            if htxt=='时间碎片化': s+=2; d.append(f'偶数页眉"时间碎片化" OK')
            elif htxt=='信息时代': s+=2; d.append(f'奇数页眉"信息时代" OK')
            # 居中：显式 jc=center，或使用内置页眉样式（默认居中）
            for p in hroot.iter(W+'p'):
                pPr=child(p,W+'pPr')
                if pPr is not None:
                    jc=child(pPr,W+'jc')
                    if jc is not None and GA(jc,'w:val','')=='center':
                        seen_center.add(name); break
                    pStyle=child(pPr,W+'pStyle')
                    if pStyle is not None and GA(pStyle,'w:val','') in ('a3','a4','a5','1','2'):
                        seen_center.add(name); break
        s += min(len(seen_center), 2); d.append(f'页眉居中 {len(seen_center)} 个')
        # 页码：只看页脚部件里的 PAGE 域和加粗
        footers = {n: x for n, x in self.all_parts.items() if 'footer' in n}
        for n, x in footers.items():
            if has_field(x, 'PAGE') and '<w:b' in x:
                s += 3; d.append(f'页码域加粗 OK ({n})')
                break
        return self.R(8,'页眉页脚',s,f,d)
    
    # ====== R9: 表格 (9) ======
    def r9(self):
        d=[]; s=0; f=9
        tbls=list(self.body.iter(W+'tbl'))
        if not tbls: return self.R(9,'表格',s,f,d)
        tbl=tbls[0]
        rows=children(tbl,W+'tr')
        if len(rows)>=5: s+=1.5; d.append(f'{len(rows)}行 OK')
        tblPr=child(tbl,W+'tblPr')
        if tblPr is not None:
            if GA(child(tblPr,W+'tblW'),'w:type','')=='auto': s+=1; d.append('自动列宽 OK')
            if GA(child(tblPr,W+'jc'),'w:val','')=='center': s+=1.5; d.append('居中 OK')
        # 黑体小四
        font_ok=False
        for row in rows:
            for cell in children(row,W+'tc'):
                for r in cell.iter(W+'r'):
                    rPr=child(r,W+'rPr')
                    if rPr is not None:
                        rf=child(rPr,W+'rFonts')
                        szE=child(rPr,W+'sz')
                        if rf is not None and GA(rf,'w:eastAsia','')=='黑体' and szE is not None and GA(szE,'w:val','')=='24':
                            font_ok=True; break
                if font_ok: break
            if font_ok: break
        if font_ok: s+=2.5; d.append('黑体小四 OK')
        # 中部两端对齐 — 垂直居中+段落默认两端对齐
        va_ok=any(GA(child(child(c,W+'tcPr'),W+'vAlign'),'w:val','')=='center' for row in rows[:3] for c in children(row,W+'tc') if child(c,W+'tcPr') is not None)
        if va_ok: s+=2.5; d.append('中部(垂直居中) OK')
        return self.R(9,'表格',s,f,d)
    
    # ====== R10: 表格边框 (9) ======
    def r10(self):
        d=[]; s=0; f=9
        tbls=list(self.body.iter(W+'tbl'))
        if not tbls: return self.R(10,'表格边框',s,f,d)
        tblPr=child(tbls[0],W+'tblPr')
        if tblPr is None: return self.R(10,'表格边框',s,f,d)
        tblBdr=child(tblPr,W+'tblBorders')
        if tblBdr is None: return self.R(10,'表格边框',s,f,d)
        outer_ok=sum(1 for bt in ['top','bottom','left','right']
            if (b:=child(tblBdr,W+bt)) is not None and GA(b,'w:val','')=='double' and GA(b,'w:sz','')=='12' and GA(b,'w:color','')=='0070C0')
        if outer_ok>=4: s+=5; d.append('外框蓝色双实线 OK')
        elif outer_ok>0: s+=2; d.append(f'外框部分({outer_ok}/4)')
        inner_ok=sum(1 for bt in ['insideH','insideV']
            if (b:=child(tblBdr,W+bt)) is not None and GA(b,'w:val','')=='single' and GA(b,'w:sz','')=='4' and GA(b,'w:color','')=='FFC000')
        if inner_ok>=2: s+=4; d.append('内框橙色单实线 OK')
        elif inner_ok>0: s+=2; d.append(f'内框部分({inner_ok}/2)')
        return self.R(10,'表格边框',s,f,d)
    
    # ====== R11: 题注 (9) ======
    def r11(self):
        d=[]; s=0; f=9
        for p in children(self.body,W+'p'):
            txt=get_text(p).strip()
            if '阅读媒介' in txt:
                s+=3; d.append(f'题注: "{txt}"')
                pPr=child(p,W+'pPr')
                if pPr is not None:
                    ps=child(pPr,W+'pStyle')
                    if ps is not None and GA(ps,'w:val','')=='a9': s+=3; d.append('caption样式(a9) OK')
                    jc=child(pPr,W+'jc')
                    if jc is not None and GA(jc,'w:val','')=='center': s+=3; d.append('居中 OK')
                return self.R(11,'题注',s,f,d)
        return self.R(11,'题注',s,f,d)
    
    def report(self):
        results=[] if not hasattr(self,'_results') else self._results
        pc=sum(1 for r in results[:-1] if r['st']=='pass')
        fc=sum(1 for r in results[:-1] if r['st']=='fail')
        cards=''
        for r in results[:-1]:
            pct=int(r['s']/r['full']*100) if r['full']>0 else 0
            cc='pass-fill' if r['st']=='pass' else 'fail-fill'
            its='\n'.join(f'<div class="di">{d}</div>' for d in r['d'])
            cards+=f'<div class="card"><div class="ch"><span class="rn">{r["req"]}. {r["name"]}</span><span class="sb {r["st"]}">{r["s"]}/{r["full"]}</span></div><div class="pb"><div class="pf {cc}" style="width:{pct}%"></div></div>{its}</div>'
        sc='#4caf50' if self.total>=60 else '#f44336'
        return f'<!DOCTYPE html><html lang="zh"><head><meta charset="UTF-8"><title>评分-文件夹8</title><style>*{{margin:0;padding:0;box-sizing:border-box}}body{{font-family:"Microsoft YaHei",sans-serif;background:#f5f5f5;padding:20px}}.c{{max-width:900px;margin:0 auto}}.h{{background:linear-gradient(135deg,#1a73e8,#0d47a1);color:#fff;padding:30px;border-radius:12px;margin-bottom:20px}}.ts{{font-size:48px;font-weight:bold;margin:10px 0}}.card{{background:#fff;border-radius:10px;padding:20px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.1)}}.ch{{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}}.rn{{font-size:16px;font-weight:500}}.sb{{padding:4px 12px;border-radius:20px;font-size:14px;font-weight:500}}.pass{{background:#e8f5e9;color:#2e7d32}}.fail{{background:#ffebee;color:#c62828}}.di{{font-size:13px;color:#666;padding:3px 0;border-bottom:1px solid #f0f0f0}}.di:last-child{{border:none}}.pb{{height:6px;background:#e0e0e0;border-radius:3px;margin-top:8px}}.pf{{height:100%;border-radius:3px;transition:width 0.3s}}.pass-fill{{background:#4caf50}}.fail-fill{{background:#f44336}}.s{{display:flex;gap:15px;margin-bottom:20px}}.sc{{flex:1;background:#fff;border-radius:10px;padding:20px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,0.1)}}.sc h3{{font-size:14px;color:#999;margin-bottom:5px}}.sc .v{{font-size:28px;font-weight:bold}}</style></head><body><div class="c"><div class="h"><h1>Word评分报告-文件夹8</h1><p>11题 | 100分</p><div class="ts">{self.total}分</div></div><div class="s"><div class="sc"><h3>总分</h3><div class="v" style="color:{sc}">{self.total}</div></div><div class="sc"><h3>通过</h3><div class="v" style="color:#4caf50">{pc}</div></div><div class="sc"><h3>未通过</h3><div class="v" style="color:#f44336">{fc}</div></div></div>{cards}</div></body></html>'

def main():
    import sys
    sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')
    base=os.path.dirname(os.path.abspath(sys.argv[1])) if len(sys.argv)>1 else os.path.dirname(os.path.abspath(__file__))
    path=sys.argv[1] if len(sys.argv)>1 else base+'/1结果.docx'
    g=Grader8(path)
    g._results=g.grade()
    g.total=g._results[-1]['s']
    print('='*60)
    print(f'Word评分-文件夹8 | {os.path.basename(path)}')
    print('='*60)
    for r in g._results[:-1]:
        print(f"\n[{r['st'].upper()}] {r['req']}. {r['name']} - {r['s']}/{r['full']}")
        for d in r['d']: print(f'  {d}')
    print(f"\n{'='*60}\n总分: {g._results[-1]['s']}/100")
    html=g.report()
    with open(base+'/report_8.html','w',encoding='utf-8') as f: f.write(html)
    print(f'\nHTML: {base}/report_8.html')

if __name__=='__main__':
    main()
