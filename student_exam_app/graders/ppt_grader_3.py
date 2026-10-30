"""
PPT 2010 评分器 - 文件夹3

题干(7项):
1. 主题"主管人员" + Slide1标题幻灯片 + 副标题"中国杂交水稻之父"30号
2. Slide2后插入空白新幻灯片(slide3) + 图片(袁隆平.jpg) + 高10cm宽15cm
3. 全部页脚"稻田里的守望者"
4. 图片劈裂动画(中央向上下展开)
5. 全部百叶窗切换
6. 放映观众自行浏览
"""
import zipfile, os, sys, io, xml.etree.ElementTree as ET, re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import slide_no

P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'

class PPTGrader3:
    def __init__(self, path):
        self.path=path; self.zf=None
        self.slides={}; self.theme_name=''; self.pres_props_xml=''

    def load(self):
        self.zf=zipfile.ZipFile(self.path); self.namelist=self.zf.namelist()
        for f in self.namelist:
            if f.startswith('ppt/slides/slide') and f.endswith('.xml') and '_rels' not in f:
                self.slides[f]={'root': ET.fromstring(self.zf.read(f).decode('utf-8')), 'xml': self.zf.read(f).decode('utf-8')}
        if 'ppt/theme/theme1.xml' in self.namelist:
            self.theme_name=ET.fromstring(self.zf.read('ppt/theme/theme1.xml').decode('utf-8')).get('name','')
        if 'ppt/presProps.xml' in self.namelist:
            self.pres_props_xml=self.zf.read('ppt/presProps.xml').decode('utf-8')

    def get_slide_count(self): return len(self.slides)

    def get_slide_layout(self, fname):
        name=fname.split('/')[-1]; rf=f'ppt/slides/_rels/{name}.rels'
        if rf in self.namelist:
            for rel in ET.fromstring(self.zf.read(rf).decode('utf-8')).iter(f'{{{PKG}}}Relationship'):
                if 'slideLayout' in rel.get('Type',''): return rel.get('Target','').split('/')[-1].replace('.xml','')
        return ''

    def has_image(self, fname):
        name=fname.split('/')[-1]; rf=f'ppt/slides/_rels/{name}.rels'
        if rf in self.namelist:
            for rel in ET.fromstring(self.zf.read(rf).decode('utf-8')).iter(f'{{{PKG}}}Relationship'):
                if 'image' in rel.get('Type',''): return rel.get('Target','')
        return ''

    def grade(self):
        self.load()
        rr=[self.r1(),self.r2(),self.r3(),self.r4(),self.r5(),self.r6()]
        t=round(sum(x['s'] for x in rr),1)
        rr.append({'req':0,'name':'总分','s':t,'full':100,'d':[],'st':'PASS' if t>=60 else 'FAIL'})
        return rr

    def R(self,req,name,score,full,detail):
        s=min(float(score),float(full))
        return {'req':req,'name':name,'s':round(s,1),'full':int(full),'d':detail,'st':'PASS' if s>=full*.6 else 'FAIL'}

    # R1: 主题+Slide1标题幻灯片+副标题30号 (15)
    def r1(self):
        d=[]; s=0; full=15
        if self.theme_name=='主管人员': s+=5; d.append('主题"主管人员" OK')
        elif '主管人员' in self.theme_name: s+=5; d.append(f'主题={self.theme_name}')

        if 'ppt/slides/slide1.xml' not in self.slides: return self.R(1,'主题+标题',s,full,d)
        layout=self.get_slide_layout('ppt/slides/slide1.xml')
        if layout=='slideLayout1': s+=5; d.append('标题幻灯片版式 OK')

        s1=self.slides['ppt/slides/slide1.xml']['xml']
        if '中国杂交水稻之父' in s1: s+=3; d.append('副标题文本 OK')
        if '3000' in s1: s+=2; d.append('字号30 OK')
        return self.R(1,'主题+标题',s,full,d)

    # R2: Slide2后插入空白幻灯片+图片+尺寸 (20)
    def r2(self):
        d=[]; s=0; full=20
        if self.get_slide_count()>=3: s+=5; d.append(f'新增幻灯片({self.get_slide_count()}张) OK')
        else: d.append(f'只有{self.get_slide_count()}张')

        if 'ppt/slides/slide3.xml' not in self.slides: return self.R(2,'插入新幻灯片',s,full,d)
        layout=self.get_slide_layout('ppt/slides/slide3.xml')
        if layout=='slideLayout7': s+=5; d.append('空白版式 OK')

        img=self.has_image('ppt/slides/slide3.xml')
        if img: s+=3; d.append(f'图片({img}) OK')

        s3=self.slides['ppt/slides/slide3.xml']['xml']
        offs=re.findall(r'<a:off x="(\d+)" y="(\d+)"',s3)
        exts=re.findall(r'<a:ext cx="(\d+)" cy="(\d+)"',s3)
        for i,(ox,oy) in enumerate(offs):
            if ox=='0' and oy=='0': continue
            cx,cy=exts[i] if i<len(exts) else ('0','0')
            w=int(cx)/360000; h=int(cy)/360000
            if abs(w-15)<1: s+=4; d.append(f'宽15cm OK')
            if abs(h-10)<1: s+=3; d.append(f'高10cm OK')
            break
        return self.R(2,'插入新幻灯片',s,full,d)

    # R3: 全部页脚"稻田里的守望者" (15)
    def r4(self, name='页脚'):
        """Called as r3 (R4 in spec)"""
        d=[]; s=0; full=15
        slide_list=sorted(self.slides.keys(), key=slide_no)
        ftr_count=sum(1 for k in slide_list if '稻田里的守望者' in self.slides[k]['xml'])
        if ftr_count>=len(slide_list):
            s+=15; d.append(f'页脚({ftr_count}/{len(slide_list)}) OK')
        elif ftr_count>0:
            s+=8; d.append(f'部分页脚({ftr_count}/{len(slide_list)})')
        return self.R(3,'页脚',s,full,d)

    # R4: 图片劈裂动画(barn) (15)
    def r3(self, name='动画'):
        """Called as r4 (R5 in spec)"""
        d=[]; s=0; full=15
        if 'ppt/slides/slide3.xml' not in self.slides: return self.R(4,'图片动画',s,full,d)
        s3=self.slides['ppt/slides/slide3.xml']['xml']
        if 'animEffect' in s3:
            s+=7; d.append('动画效果 OK')
            if 'barn' in s3: s+=4; d.append('劈裂(barn) OK')
            if 'outHorizontal' in s3: s+=4; d.append('中央向上下展开 OK')
        return self.R(4,'图片动画',s,full,d)

    # R5: 全部百叶窗切换 (15)
    def r5(self, name='切换'):
        """Called as r5 (R6 in spec)"""
        d=[]; s=0; full=20
        slide_list=sorted(self.slides.keys(), key=slide_no)
        blinds_count=sum(1 for k in slide_list if '<p:blinds' in self.slides[k]['xml'])
        if blinds_count>=len(slide_list):
            s+=20; d.append(f'百叶窗({blinds_count}/{len(slide_list)}) OK')
        elif blinds_count>0:
            s+=10; d.append(f'部分百叶窗({blinds_count})')
        return self.R(5,'切换效果',s,full,d)

    # R6: 观众自行浏览 (15)
    def r6(self, name='放映'):
        d=[]; s=0; full=15
        if '<p:browse' in self.pres_props_xml: s+=15; d.append('观众自行浏览 OK')
        else: d.append('未找到')
        return self.R(6,'放映方式',s,full,d)

    def r3_alt(self):
        """R3: 页脚"""
        return self.r4()
    def r4_alt(self):
        """R4: 动画"""
        return self.r3()


def main():
    if sys.platform=='win32':
        try:sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')
        except: pass
    base=os.path.dirname(os.path.abspath(__file__))
    path=sys.argv[1] if len(sys.argv)>1 else os.path.join(base,'1结果.pptx')
    print(f'评分: {os.path.basename(path)}')
    print('='*50)
    g=PPTGrader3(path)
    results=g.grade()
    for r in results[:-1]:
        icon='PASS' if r['st']=='PASS' else 'FAIL'
        print(f"[{icon}] {r['req']:2d}. {r['name']:<18s}  {r['s']:5.1f} / {r['full']}")
        for d in r['d']: print(f'       {d}')
    print(f'{"="*50}\n  总分: {results[-1]["s"]} / 100')

if __name__=='__main__':
    main()
