"""
PPT 2010 评分器 - 文件夹4

题干(7项):
1. 主题"元素" + Slide1标题幻灯片 + 副标题"虚拟3D世界"华文新魏28号
2. Slide2后插入空白幻灯片 + 图片(20×15cm居中)
3. 全部页脚"元宇宙的特征与属性"
4. 图片回旋动画 + 全部碎片切换 + 声音捶打
5. Slide2备注
6. 放映观众自行浏览
"""
import zipfile, os, sys, io, xml.etree.ElementTree as ET, re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import slide_no
P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'

class PPTGrader4:
    def __init__(self, path):
        self.path=path; self.zf=None; self.slides={}
        self.theme_name=''; self.pres_props_xml=''

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
        if self.theme_name=='元素': s+=5; d.append('主题"元素" OK')
        elif '元素' in self.theme_name: s+=5; d.append(f'主题={self.theme_name}')

        if 'ppt/slides/slide1.xml' not in self.slides: return self.R(1,'主题+标题',s,full,d)
        layout=self.get_slide_layout('ppt/slides/slide1.xml')
        if layout=='slideLayout1': s+=3; d.append('标题幻灯片 OK')

        s1=self.slides['ppt/slides/slide1.xml']['xml']
        if '虚拟' in s1 and '3D' in s1 and '世界' in s1: s+=3; d.append('副标题文本 OK')
        if '华文新魏' in s1: s+=2; d.append('华文新魏 OK')
        if '2800' in s1: s+=2; d.append('字号28 OK')
        return self.R(1,'主题+标题',s,full,d)

    # R2: 插入空白幻灯片+图片+尺寸居中 (20)
    def r2(self):
        d=[]; s=0; full=20
        if self.get_slide_count()>=3: s+=4; d.append(f'新增幻灯片({self.get_slide_count()}张) OK')
        if 'ppt/slides/slide3.xml' not in self.slides: return self.R(2,'插入新幻灯片',s,full,d)

        layout=self.get_slide_layout('ppt/slides/slide3.xml')
        if layout=='slideLayout7': s+=4; d.append('空白版式 OK')
        if any('image' in rel.get('Type','') for rel in ET.fromstring(self.zf.read('ppt/slides/_rels/slide3.xml.rels').decode('utf-8')).iter(f'{{{PKG}}}Relationship')) if 'ppt/slides/_rels/slide3.xml.rels' in self.namelist else False:
            s+=3; d.append('图片 OK')

        s3=self.slides['ppt/slides/slide3.xml']['xml']
        offs=re.findall(r'<a:off x="(\d+)" y="(\d+)"',s3)
        exts=re.findall(r'<a:ext cx="(\d+)" cy="(\d+)"',s3)
        for i,(ox,oy) in enumerate(offs):
            if ox=='0' and oy=='0': continue
            cx,cy=exts[i] if i<len(exts) else ('0','0')
            w=int(cx)/360000; h=int(cy)/360000
            x=int(ox)/360000; y=int(oy)/360000
            if abs(w-20)<1: s+=3; d.append(f'宽20cm OK')
            if abs(h-15)<1: s+=3; d.append(f'高15cm OK')
            # Center: (25.4-w)/2≈2.7, (19.05-h)/2≈2.0
            if abs(x-(25.4-w)/2)<1: s+=2; d.append('水平居中 OK')
            if abs(y-(19.05-h)/2)<1: s+=1; d.append('垂直居中 OK')
            break
        return self.R(2,'插入新幻灯片',s,full,d)

    # R3: 全部页脚 (15)
    def r3(self):
        d=[]; s=0; full=15
        slide_list=sorted(self.slides.keys(), key=slide_no)
        fc=sum(1 for k in slide_list if '元宇宙的特征与属性' in self.slides[k]['xml'])
        if fc>=len(slide_list): s+=15; d.append(f'页脚({fc}/{len(slide_list)}) OK')
        elif fc>0: s+=8; d.append(f'部分页脚({fc})')
        return self.R(3,'页脚',s,full,d)

    # R4: 回旋动画+碎片切换+声音 (20)
    def r4(self):
        d=[]; s=0; full=20
        if 'ppt/slides/slide3.xml' in self.slides:
            s3=self.slides['ppt/slides/slide3.xml']['xml']
            if 'animEffect' in s3: s+=6; d.append('图片动画 OK')

        slide_list=sorted(self.slides.keys(), key=slide_no)
        shred_count=sum(1 for k in slide_list if 'shred' in self.slides[k]['xml'])
        if shred_count>=len(slide_list): s+=8; d.append(f'碎片切换({shred_count}/{len(slide_list)}) OK')

        if 'ppt/media/audio1.wav' in self.namelist: s+=6; d.append('声音锤打(audio1.wav) OK')
        return self.R(4,'动画+切换+声音',s,full,d)

    # R5: Slide2备注 (15)
    def r5(self):
        d=[]; s=0; full=15
        if 'ppt/notesSlides/notesSlide1.xml' in self.namelist:
            notes=self.zf.read('ppt/notesSlides/notesSlide1.xml').decode('utf-8')
            note_texts=re.findall(r'<a:t[^>]*>([^<]+)</a:t>',notes)
            full_note=''.join(note_texts)
            if '元宇宙本身不是一种技术' in full_note:
                s+=15; d.append(f'备注文本 OK')
            elif note_texts:
                s+=8; d.append(f'备注存在({len(note_texts)} chars)')
        else:
            d.append('未找到备注')
        return self.R(5,'备注',s,full,d)

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
    g=PPTGrader4(path)
    results=g.grade()
    for r in results[:-1]:
        icon='PASS' if r['st']=='PASS' else 'FAIL'
        print(f"[{icon}] {r['req']:2d}. {r['name']:<18s}  {r['s']:5.1f} / {r['full']}")
        for d in r['d']: print(f'       {d}')
    print(f'{"="*50}\n  总分: {results[-1]["s"]} / 100')

if __name__=='__main__':
    main()
