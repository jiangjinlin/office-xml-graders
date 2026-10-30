"""Word 2010 评分器 - 文件夹10 (v2 修订驱动 + 章节调整)"""
import zipfile, os, re, io, sys
import xml.etree.ElementTree as ET
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import mail_merge_info, merge_fields

W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
RNS='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
V='{urn:schemas-microsoft-com:vml}'
O='{urn:schemas-microsoft-com:office:office}'
WP='{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'

NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'v':'urn:schemas-microsoft-com:vml',
    'o':'urn:schemas-microsoft-com:office:office'}

def child(e,t):
    for c in e:
        if c.tag==t: return c
    return None

def children(e,t):
    return [c for c in e if c.tag==t]

def GA(elem,attr,default=None):
    if elem is None: return default
    if ':' in attr:
        ns,l=attr.split(':',1)
        return elem.get('{'+NS.get(ns,'')+'}'+l, default)
    return elem.get(attr, default)

def gt(elem):
    return ''.join(t.text or '' for t in elem.iter(W+'t'))

class Grader10:
    def __init__(self,path):
        self.path=path
        self.body=None
        self.all_xml=''
        self.headers={}
        self.rels={}
    
    def load(self):
        with zipfile.ZipFile(self.path) as z:
            d=z.read('word/document.xml').decode('utf-8')
            self.doc_root=ET.fromstring(d.encode('utf-8'))
            for c in self.doc_root:
                if c.tag==W+'body': self.body=c; break
            all_x=[d]
            for n in z.namelist():
                if n.endswith('.xml'):
                    try:
                        c=z.read(n).decode('utf-8',errors='replace')
                        all_x.append(c)
                        if 'header' in n: self.headers[n]=c
                    except: pass
            self.all_xml='\n'.join(all_x)
            try:
                rels=ET.fromstring(z.read('word/_rels/document.xml.rels'))
                for rel in rels:
                    self.rels[rel.get('Id','')]={'t':rel.get('Target',''),'y':rel.get('Type','')}
            except: pass
    
    def grade(self):
        self.load()
        rr=[]
        rr.append(self.r1()); rr.append(self.r2()); rr.append(self.r3())
        rr.append(self.r4()); rr.append(self.r5()); rr.append(self.r6())
        rr.append(self.r7()); rr.append(self.r8()); rr.append(self.r9())
        rr.append(self.r10()); rr.append(self.r11()); rr.append(self.r12())
        t=round(sum(x['s'] for x in rr),1)
        rr.append({'req':0,'name':'总分','s':t,'full':100,'d':[],'st':'pass' if t>=60 else 'fail'})
        return rr
    
    def R(self,req,name,score,full,detail):
        score=float(score); full=float(full); score=min(score,full)
        st='pass' if score>=full*0.6 else 'fail'
        return {'req':req,'name':name,'s':round(score,1),'full':int(full),'d':detail,'st':st}
    
    def r1(self):
        d=[]; s=0; f=9
        for sp in self.doc_root.iter(W+'sectPr'):
            pgMar=child(sp,W+'pgMar')
            dg=child(sp,W+'docGrid')
            if pgMar is not None:
                if GA(pgMar,'w:top','')=='1588': s+=2; d.append('上边距2.8cm OK')
                if GA(pgMar,'w:bottom','')=='1588': s+=2; d.append('下边距2.8cm OK')
                if GA(pgMar,'w:left','')=='1701': s+=2; d.append('左边距3cm OK')
                if GA(pgMar,'w:right','')=='1701': s+=2; d.append('右边距3cm OK')
            if dg is not None and GA(dg,'w:linePitch','')=='325': s+=2; d.append('每页42行 OK')
            break
        return self.R(1,'页面设置',s,f,d)
    
    def r2(self):
        d=[]; s=0; f=6
        paras=children(self.body,W+'p')
        if not paras: return self.R(2,'标题',s,f,d)
        pPr=child(paras[0],W+'pPr')
        rPr=child(pPr,W+'rPr') if pPr is not None else None
        if rPr is not None:
            if GA(child(rPr,W+'rFonts'),'w:eastAsia','')=='黑体': s+=2; d.append('黑体 OK')
            if GA(child(rPr,W+'sz'),'w:val','')=='44': s+=2; d.append('二号(44) OK')
            if GA(child(rPr,W+'spacing'),'w:val','')=='40': s+=1.5; d.append('加宽2磅 OK')
        if pPr is not None:
            if GA(child(pPr,W+'jc'),'w:val','')=='center': s+=1.5; d.append('居中 OK')
            if GA(child(pPr,W+'spacing'),'w:afterLines','')=='100': s+=1; d.append('段后1行 OK')
        return self.R(2,'标题',s,f,d)
    
    def r3(self):
        d=[]; s=0; f=9
        paras=children(self.body,W+'p')
        ind_ok=0; line_ok=0; body_n=0
        for p in paras:
            txt=gt(p)
            if not txt or '成绩通知单' in txt or '期末成绩单' in txt: continue
            if txt.startswith('尊敬的各位家长'):
                pPr=child(p,W+'pPr')
                if pPr is not None and GA(child(pPr,W+'jc'),'w:val','')=='left': s+=2; d.append('左对齐 OK')
                continue
            body_n+=1
            pPr=child(p,W+'pPr')
            if pPr is not None:
                ind=child(pPr,W+'ind')
                if ind is not None:
                    if GA(ind,'w:firstLineChars','')=='200': ind_ok+=1
                spc=child(pPr,W+'spacing')
                if spc is not None and GA(spc,'w:line','')=='360' and GA(spc,'w:lineRule','')=='exact': line_ok+=1
        if body_n>0:
            if ind_ok/body_n>=0.2: s+=4; d.append(f'首行缩进({ind_ok}/{body_n}) OK')
            if line_ok/body_n>=0.7: s+=4; d.append(f'固定18磅({line_ok}/{body_n}) OK')
        return self.R(3,'正文缩进行距',s,f,d)
    
    def r4(self):
        d=[]; s=0; f=9
        paras=children(self.body,W+'p')
        font_ok=0; sz_ok=0; body_n=0
        for p in paras:
            txt=gt(p)
            if not txt or '成绩通知单' in txt or '期末成绩单' in txt: continue
            body_n+=1
            for r in children(p,W+'r'):
                rPr=child(r,W+'rPr')
                if rPr is not None:
                    if GA(child(rPr,W+'rFonts'),'w:eastAsia','')=='黑体': font_ok+=1
                    if GA(child(rPr,W+'sz'),'w:val','')=='24': sz_ok+=1
                break
        if body_n>0:
            if font_ok/body_n>=0.7: s+=5; d.append(f'黑体({font_ok}/{body_n}) OK')
            if sz_ok/body_n>=0.7: s+=4; d.append(f'小四({sz_ok}/{body_n}) OK')
        return self.R(4,'正文字体',s,f,d)
    
    def r5(self):
        d=[]; s=0; f=9
        # 解析 numbering.xml：numId → abstractNumId → lvl0 的 numFmt
        num2abs={}
        numroot=None
        with zipfile.ZipFile(self.path) as z:
            if 'word/numbering.xml' in z.namelist():
                numroot=ET.fromstring(z.read('word/numbering.xml'))
        if numroot is not None:
            for num in numroot.iter(W+'num'):
                a=child(num,W+'abstractNumId')
                num2abs[num.get(W+'numId','')]=a.get(W+'val') if a is not None else None

        def fmt_of(numid):
            if numroot is None or numid not in num2abs: return ''
            aid=num2abs[numid]
            for an in numroot.iter(W+'abstractNum'):
                if an.get(W+'abstractNumId','')==aid:
                    lvl=child(an,W+'lvl')
                    if lvl is not None:
                        nf=child(lvl,W+'numFmt')
                        return nf.get(W+'val','') if nf is not None else ''
            return ''

        used=0
        for p in children(self.body,W+'p'):
            pPr=child(p,W+'pPr')
            npr=child(pPr,W+'numPr') if pPr is not None else None
            if npr is not None:
                nid=child(npr,W+'numId')
                if nid is not None and fmt_of(nid.get(W+'val',''))=='chineseCountingThousand':
                    used+=1
        if used>=10: s+=10; d.append(f'编号(一)(二)(三) 应用 {used} 段 OK')
        elif used>=5: s+=5; d.append(f'编号应用 {used} 段（偏少）')
        else: d.append(f'NG 编号未应用（{used} 段）')
        return self.R(5,'编号样式',s,f,d)
    
    def r6(self):
        d=[]; s=0; f=7
        for p in children(self.body,W+'p'):
            if '期末成绩单' in gt(p):
                pPr=child(p,W+'pPr')
                rPr=child(pPr,W+'rPr') if pPr is not None else None
                if rPr is not None:
                    if GA(child(rPr,W+'rFonts'),'w:eastAsia','')=='华文新魏': s+=2; d.append('华文新魏 OK')
                    if child(rPr,W+'b') is not None: s+=1.5; d.append('加粗 OK')
                    if GA(child(rPr,W+'sz'),'w:val','')=='36': s+=2; d.append('小二(36) OK')
                if pPr is not None and GA(child(pPr,W+'jc'),'w:val','')=='center': s+=1.5; d.append('居中 OK')
                return self.R(6,'第二标题',s,f,d)
        return self.R(6,'第二标题',s,f,d)
    
    def r7(self):
        d=[]; s=0; f=9
        tbls=list(self.body.iter(W+'tbl'))
        if not tbls: return self.R(7,'表格格式',s,f,d)
        tbl=tbls[0]; rows=children(tbl,W+'tr')
        if len(rows)==2: s+=1; d.append('2行 OK')
        cells0=children(rows[0],W+'tc') if rows else []
        if len(cells0)==7: s+=1; d.append('7列 OK')
        tblPr=child(tbl,W+'tblPr')
        if tblPr is not None:
            if GA(child(tblPr,W+'jc'),'w:val','')=='center': s+=1; d.append('居中 OK')
            tblBdr=child(tblPr,W+'tblBorders')
            if tblBdr is not None:
                outer=sum(1 for bt in['top','bottom','left','right'] if (b:=child(tblBdr,W+bt)) is not None and GA(b,'w:val','')=='double' and GA(b,'w:sz','')=='12' and GA(b,'w:color','')=='FF0000')
                if outer>=4: s+=2; d.append('外框红双实线 OK')
                inner=sum(1 for bt in['insideH','insideV'] if (b:=child(tblBdr,W+bt)) is not None and GA(b,'w:val','')=='single' and GA(b,'w:sz','')=='6' and GA(b,'w:color','')=='FF0000')
                if inner>=2: s+=1.5; d.append('内框红单实线 OK')
        for ri,row in enumerate(rows):
            trH=child(row,W+'trPr')
            hval=GA(child(trH,W+'trHeight'),'w:val','') if trH is not None else ''
            if hval and abs(int(hval)-567)<50: s+=1; d.append('行高1cm OK')
        for cell in cells0+children(rows[1],W+'tc') if len(rows)>1 else cells0:
            tcPr=child(cell,W+'tcPr')
            if tcPr is not None and (w:=child(tcPr,W+'tcW')) is not None:
                if abs(int(GA(w,'w:w','0'))-850)<50: s+=1; d.append('列宽1.5cm OK'); break
        for cell in cells0:
            tcPr=child(cell,W+'tcPr')
            if tcPr is not None and GA(child(tcPr,W+'shd'),'w:fill','')=='FFFF00': s+=1.5; d.append('黄色底纹 OK'); break
        return self.R(7,'表格格式',s,f,d)
    
    def r8(self):
        d=[]; s=0; f=8
        for name,hxml in self.headers.items():
            hroot=ET.fromstring(hxml)
            htxt=gt(hroot).strip()
            if htxt=='成绩通知单': s+=2; d.append('偶数"成绩通知单" OK')
            elif htxt=='安全通知书': s+=2; d.append('奇数"安全通知书" OK')
            elif htxt=='': continue
            for r in hroot.iter(W+'r'):
                rPr=child(r,W+'rPr')
                if rPr is not None and child(rPr,W+'rFonts') is not None:
                    if GA(child(rPr,W+'rFonts'),'w:eastAsia','')=='隶书': s+=2; d.append('隶书 OK'); break
        return self.R(8,'页眉',min(s,f),f,d)
    
    def r9(self):
        d=[]; s=0; f=5
        url_ok=False
        for rel in self.rels.values():
            if 'hyperlink' in rel['y'] and rel['t']=='https://cn.bing.com/':
                url_ok=True
        if url_ok: s+=5; d.append('URL正确 OK')
        else: d.append('NG 未找到指向 https://cn.bing.com/ 的超链接')
        # 锚文字必须出现在某个 hyperlink 内
        anchor_ok=False
        for link in self.doc_root.iter(W+'hyperlink'):
            if '中职学生日常行为规范' in gt(link):
                anchor_ok=True
        if anchor_ok: s+=2; d.append('超链接文字 OK')
        else: d.append('NG 超链接锚文字不符')
        return self.R(9,'超链接',s,f,d)
    
    def r10(self):
        d=[]; s=0; f=10
        tbls=list(self.body.iter(W+'tbl'))
        if not tbls: return self.R(10,'表格内容',s,f,d)
        cells0=children(children(tbls[0],W+'tr')[0],W+'tc')
        expected=['姓名','语文','数学','英语','物理','化学','总分']
        found=sum(1 for i,e in enumerate(expected) if i<len(cells0) and any(c in gt(cells0[i]) for c in [e[0],e]))
        if found>=6: s+=4; d.append(f'表头({found}/7) OK')
        elif found>=4: s+=3; d.append(f'表头部分({found}/7)')
        font_ok=any(GA(child(rPr,W+'rFonts'),'w:eastAsia','')=='宋体' for cell in cells0 for r in cell.iter(W+'r') if (rPr:=child(r,W+'rPr')) is not None)
        if font_ok: s+=2; d.append('宋体 OK')
        tcPr=child(cells0[0],W+'tcPr') if cells0 else None
        if tcPr is not None and child(tcPr,W+'shd') is not None: s+=2; d.append('加粗/黄色 OK')
        jc_ok=sum(1 for row in children(tbls[0],W+'tr') for cell in children(row,W+'tc') for p in cell.iter(W+'p') if (pp:=child(p,W+'pPr')) is not None and GA(child(pp,W+'jc'),'w:val','')=='center')
        if jc_ok>=2: s+=4; d.append(f'水平居中({jc_ok}) OK')
        return self.R(10,'表格内容',s,f,d)
    
    def r11(self):
        d=[]; s=0; f=9
        for name,hxml in self.headers.items():
            hroot=ET.fromstring(hxml)
            for sh in hroot.iter(V+'shape'):
                if 'Watermark' in sh.get('id',''):
                    tp=sh.find(V+'textpath')
                    if tp is not None and tp.get('string','')=='朝阳中职': s+=10; d.append(f'水印"朝阳中职" OK')
                    if '华文行楷' in hxml: s+=1; d.append('华文行楷 OK')
                    style=sh.get('style','')
                    if 'opacity' in style.lower() or 'semitransparent' in style.lower(): s+=2; d.append('半透明 OK')
                    if 'rotation' not in style.lower() or 'rotation:0' in style: s+=2; d.append('水平 OK')
                    return self.R(11,'水印',s,f,d)
        if '朝阳中职' in self.all_xml: s+=10; d.append('水印文字存在(字符串)')
        return self.R(11,'水印',s,f,d)
    
    def r12(self):
        d=[]; s=0; f=10
        info = mail_merge_info(self.all_xml)
        if info:
            s += 3; d.append('启用邮件合并 OK')
            if info.get('mainDocumentType') == 'formLetters':
                s += 2; d.append('主文档类型=信函 OK')
            conn = (info.get('connectString') or '') + (info.get('dataSource') or '') + (info.get('query') or '')
            if '成绩' in conn:
                s += 2; d.append('数据源指向成绩表 OK')
            else:
                d.append('数据源=%s' % (conn[:60] or '未设置'))
        else:
            d.append('未检测到 w:mailMerge')

        names = merge_fields(self.all_xml)
        if len(names) >= 5:
            s += 1.5; d.append('合并域 %d 个 OK' % len(names))
        else:
            d.append('合并域数量=%d（至少应有姓名、语文、数学等）' % len(names))
        need = {'姓名', '语文', '数学', '英语', '物理', '化学', '总分'}
        hit = len(need & set(names))
        if hit >= 6:
            s += 1.5; d.append('域覆盖 %d/7 OK' % hit)
        elif names:
            d.append('域仅覆盖 %d/7：%s' % (hit, '、'.join(sorted(set(names))[:7])))
        return self.R(12, '邮件合并', s, f, d)

    def report(self):
        results=self._results if hasattr(self,'_results') else []
        pc=sum(1 for r in results[:-1] if r['st']=='pass')
        fc=sum(1 for r in results[:-1] if r['st']=='fail')
        cards=''
        for r in results[:-1]:
            pct=int(r['s']/r['full']*100) if r['full']>0 else 0
            cc='pass-fill' if r['st']=='pass' else 'fail-fill'
            its='\n'.join(f'<div class="di">{d}</div>' for d in r['d'])
            cards+=f'<div class="card"><div class="ch"><span class="rn">{r["req"]}. {r["name"]}</span><span class="sb {r["st"]}">{r["s"]}/{r["full"]}</span></div><div class="pb"><div class="pf {cc}" style="width:{pct}%"></div></div>{its}</div>'
        sc='#4caf50' if self.total>=60 else '#f44336'
        return f'<!DOCTYPE html><html lang="zh"><head><meta charset="UTF-8"><title>评分-文件夹10</title><style>*{{margin:0;padding:0;box-sizing:border-box}}body{{font-family:"Microsoft YaHei",sans-serif;background:#f5f5f5;padding:20px}}.c{{max-width:900px;margin:0 auto}}.h{{background:linear-gradient(135deg,#1a73e8,#0d47a1);color:#fff;padding:30px;border-radius:12px;margin-bottom:20px}}.ts{{font-size:48px;font-weight:bold;margin:10px 0}}.card{{background:#fff;border-radius:10px;padding:20px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,0.1)}}.ch{{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}}.rn{{font-size:16px;font-weight:500}}.sb{{padding:4px 12px;border-radius:20px;font-size:14px;font-weight:500}}.pass{{background:#e8f5e9;color:#2e7d32}}.fail{{background:#ffebee;color:#c62828}}.di{{font-size:13px;color:#666;padding:3px 0;border-bottom:1px solid #f0f0f0}}.di:last-child{{border:none}}.pb{{height:6px;background:#e0e0e0;border-radius:3px;margin-top:8px}}.pf{{height:100%;border-radius:3px}}.pass-fill{{background:#4caf50}}.fail-fill{{background:#f44336}}.s{{display:flex;gap:15px;margin-bottom:20px}}.sc{{flex:1;background:#fff;border-radius:10px;padding:20px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,0.1)}}.sc h3{{font-size:14px;color:#999;margin-bottom:5px}}.sc .v{{font-size:28px;font-weight:bold}}</style></head><body><div class="c"><div class="h"><h1>Word评分报告-文件夹10</h1><p>11题|100分</p><div class="ts">{self.total}分</div></div><div class="s"><div class="sc"><h3>总分</h3><div class="v" style="color:{sc}">{self.total}</div></div><div class="sc"><h3>通过</h3><div class="v" style="color:#4caf50">{pc}</div></div><div class="sc"><h3>未通过</h3><div class="v" style="color:#f44336">{fc}</div></div></div>{cards}</div></body></html>'

def main():
    import sys
    sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')
    base=os.path.dirname(os.path.abspath(sys.argv[1])) if len(sys.argv)>1 else os.path.dirname(os.path.abspath(__file__))
    path=sys.argv[1] if len(sys.argv)>1 else base+'/1结果.docx'
    g=Grader10(path)
    g._results=g.grade(); g.total=g._results[-1]['s']
    print('='*60+f'\nWord评分-文件夹10 | {os.path.basename(path)}\n'+'='*60)
    for r in g._results[:-1]:
        print(f"\n[{r['st'].upper()}] {r['req']}. {r['name']} - {r['s']}/{r['full']}")
        for d in r['d']: print(f'  {d}')
    print(f"\n{'='*60}\n总分: {g._results[-1]['s']}/100")
    html=g.report()
    with open(base+'/report_10.html','w',encoding='utf-8') as f: f.write(html)
    print(f'\nHTML: {base}/report_10.html')

if __name__=='__main__':
    main()
