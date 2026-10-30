"""
PPT 2010 评分器 - 文件夹5

题干(7项):
1. 主题"穿越" + Slide1标题幻灯片 + 副标题"太空植物"华文新魏28号
2. Slide2/3→两栏内容 + 右侧图片 + 标题
3. 全部页脚"太空植物"
4. 图片动画玩具风车 + 上一动画之后
5. 全部溶解切换
6. 放映观众自行浏览
"""
import zipfile, os, sys, io, xml.etree.ElementTree as ET, re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import slide_no
P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'

class PPTGrader5:
    def __init__(self, path):
        self.path=path; self.zf=None; self.slides={}
        self.theme_name=''; self.pres_props_xml=''

    def load(self):
        self.zf=zipfile.ZipFile(self.path); self.namelist=self.zf.namelist()
        for f in self.namelist:
            if f.startswith('ppt/slides/slide') and f.endswith('.xml') and '_rels' not in f:
                self.slides[f]={'root': ET.fromstring(self.zf.read(f).decode('utf-8')),
                                'xml': self.zf.read(f).decode('utf-8')}
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

    # R1: 主题+标题幻灯片+副标题华文新魏28号 (15)
    def r1(self):
        d=[]; s=0; full=15
        if self.theme_name=='穿越': s+=5; d.append('主题"穿越" OK')
        elif '穿越' in self.theme_name: s+=5; d.append(f'主题={self.theme_name}')
        if 'ppt/slides/slide1.xml' not in self.slides: return self.R(1,'主题+标题',s,full,d)
        if self.get_slide_layout('ppt/slides/slide1.xml')=='slideLayout1': s+=3; d.append('标题幻灯片 OK')
        s1=self.slides['ppt/slides/slide1.xml']['xml']
        if '太空植物' in s1: s+=5; d.append('副标题文本 OK')
        if '2800' in s1: s+=2; d.append('字号28 OK')
        return self.R(1,'主题+标题',s,full,d)

    # R2: Slide2/3两栏内容+图片+标题 (20)
    def r2(self):
        d=[]; s=0; full=20
        for sn, title_kw in [('slide2','太空辣椒'),('slide3','太空南瓜')]:
            fname=f'ppt/slides/{sn}.xml'
            if fname not in self.slides: continue
            if self.get_slide_layout(fname)=='slideLayout4': s+=5; d.append(f'{sn}两栏内容 OK')
            if self.has_image(fname): s+=5; d.append(f'{sn}图片 OK')
        return self.R(2,'版式+图片+标题',s,full,d)

    # R3: 全部页脚"太空植物" (15)
    def r3(self):
        d=[]; s=0; full=15
        slide_list=sorted(self.slides.keys(), key=slide_no)
        fc=sum(1 for k in slide_list if '太空植物' in self.slides[k]['xml'])
        if fc>=len(slide_list) and fc>=3: s+=15; d.append(f'页脚({fc}/{len(slide_list)}) OK')
        else: d.append(f'页脚 {fc}/{len(slide_list)}（素材第1张自带，需全部设置才给分）')
        return self.R(3,'页脚',s,full,d)

    # R4: 图片玩具风车动画+上一动画之后 (20)
    def r4(self):
        d=[]; s=0; full=20
        anim_count=0; timing_count=0
        for sn in ['slide2','slide3']:
            fname=f'ppt/slides/{sn}.xml'
            if fname not in self.slides: continue
            x=self.slides[fname]['xml']
            if 'animEffect' in x: anim_count+=1
            if 'afterEffect' in x: timing_count+=1
        if anim_count>=2: s+=12; d.append(f'图片动画({anim_count}/2) OK')
        elif anim_count>=1: s+=6; d.append(f'部分动画({anim_count}/2)')
        if timing_count>=2: s+=8; d.append(f'动画之后开始({timing_count}/2) OK')
        return self.R(4,'动画+时序',s,full,d)

    # R5: 全部溶解切换 (15)
    def r5(self):
        d=[]; s=0; full=15
        slide_list=sorted(self.slides.keys(), key=slide_no)
        dc=sum(1 for k in slide_list if 'dissolve' in self.slides[k]['xml'])
        if dc>=len(slide_list) and dc>=3: s+=15; d.append(f'溶解切换({dc}/{len(slide_list)}) OK')
        else: d.append(f'溶解切换 {dc}/{len(slide_list)}（素材自带个别溶解，需全部设置才给分）')
        return self.R(5,'切换效果',s,full,d)

    # R6: 放映方式 (15)
    def r6(self):
        d=[]; s=0; full=15
        if '<p:browse' in self.pres_props_xml: s+=15; d.append('观众自行浏览 OK')
        else: d.append('未找到')
        return self.R(6,'放映方式',s,full,d)


def main():
    if sys.platform=='win32':
        try:sys.stdout=io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')
        except: pass
    base=os.path.dirname(os.path.abspath(__file__))
    path=sys.argv[1] if len(sys.argv)>1 else os.path.join(base,'1结果.pptx')
    print(f'评分: {os.path.basename(path)}')
    print('='*50)
    g=PPTGrader5(path)
    results=g.grade()
    for r in results[:-1]:
        icon='PASS' if r['st']=='PASS' else 'FAIL'
        print(f"[{icon}] {r['req']:2d}. {r['name']:<18s}  {r['s']:5.1f} / {r['full']}")
        for det in r['d']: print(f'       {det}')
    print(f'{"="*50}\n  总分: {results[-1]["s"]} / 100')

if __name__=='__main__':
    main()
