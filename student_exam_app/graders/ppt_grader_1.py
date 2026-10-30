"""
PPT 2010 评分器 - 文件夹1 (v1)

题干(7项, 10评分点):
1. 主题"奥斯汀" + 第1张标题幻灯片版式 + 标题华文行楷40号 + 副标题华文楷体20号
2. 第2张两栏内容 + 右边插入壮美草原.jpg + 标题"万马奔腾"
3. 图片样式柔化边缘矩形 + 动画浮入下浮
4. 所有幻灯片编号 + 页脚"我的草原我的马"
5. 第3张两栏内容 + 移动第4张内容到第3张右侧 + 删除第4张 + 标题"草原生态补奖政策"
6. 切换擦除 + 放映观众自行浏览
"""
import zipfile, os, sys, io, xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import slide_no

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

class PPTGrader1:
    def __init__(self, path):
        self.path = path; self.zf = None
        self.slides = {}; self.all_xml = ''; self.theme_name = ''
        self.pres_props_xml = ''

    def load(self):
        self.zf = zipfile.ZipFile(self.path)
        self.namelist = self.zf.namelist()
        all_x = []

        # Presentation
        d = self.zf.read('ppt/presentation.xml').decode('utf-8')
        all_x.append(d); self.pres_xml = d

        # Each slide
        for fname in self.namelist:
            if fname.startswith('ppt/slides/slide') and fname.endswith('.xml') \
               and '_rels' not in fname:
                d = self.zf.read(fname).decode('utf-8')
                all_x.append(d)
                self.slides[fname] = {'root': ET.fromstring(d), 'xml': d}

        # Theme
        if 'ppt/theme/theme1.xml' in self.namelist:
            d = self.zf.read('ppt/theme/theme1.xml').decode('utf-8')
            all_x.append(d)
            self.theme_name = ET.fromstring(d).get('name', '')

        # PresProps
        if 'ppt/presProps.xml' in self.namelist:
            d = self.zf.read('ppt/presProps.xml').decode('utf-8')
            all_x.append(d); self.pres_props_xml = d

        # Slide master
        if 'ppt/slideMasters/slideMaster1.xml' in self.namelist:
            d = self.zf.read('ppt/slideMasters/slideMaster1.xml').decode('utf-8')
            all_x.append(d)

        self.all_xml = '\n'.join(all_x)

    def get_slide_count(self):
        return len(self.slides)

    def get_slide_files(self):
        return sorted(self.slides.keys(), key=slide_no)

    def get_slide_layout(self, slide_fname):
        name = slide_fname.split('/')[-1]
        rel_fname = f'ppt/slides/_rels/{name}.rels'
        if rel_fname in self.namelist:
            rels_xml = ET.fromstring(self.zf.read(rel_fname).decode('utf-8'))
            for rel in rels_xml.iter(f'{{{PKG}}}Relationship'):
                t = rel.get('Type', '')
                if 'slideLayout' in t:
                    target = rel.get('Target', '')
                    return target.split('/')[-1].replace('.xml', '')
        return ''

    def has_image(self, slide_fname):
        name = slide_fname.split('/')[-1]
        rel_fname = f'ppt/slides/_rels/{name}.rels'
        if rel_fname in self.namelist:
            rels_xml = ET.fromstring(self.zf.read(rel_fname).decode('utf-8'))
            for rel in rels_xml.iter(f'{{{PKG}}}Relationship'):
                if 'image' in rel.get('Type', ''):
                    return rel.get('Target', '')
        return ''

    def grade(self):
        self.load()
        rr = [self.r1(), self.r2(), self.r3(), self.r4(),
              self.r5(), self.r6()]
        t = round(sum(x['s'] for x in rr), 1)
        rr.append({'req': 0, 'name': '总分', 's': t, 'full': 100,
                   'd': [], 'st': 'PASS' if t >= 60 else 'FAIL'})
        return rr

    def R(self, req, name, score, full, detail):
        s = min(float(score), float(full))
        st = 'PASS' if s >= full * 0.6 else 'FAIL'
        return {'req': req, 'name': name, 's': round(s, 1),
                'full': int(full), 'd': detail, 'st': st}

    # R1: 主题+幻灯片1版式+标题/副标题字体 (15)
    def r1(self):
        d = []; s = 0; full = 15
        # 主题
        if self.theme_name == '奥斯汀':
            s += 4; d.append('主题"奥斯汀" OK')
        elif '奥斯汀' in self.theme_name:
            s += 4; d.append(f'主题={self.theme_name}')

        # Slide1（素材本就是标题幻灯片版式，不给分）
        if 'ppt/slides/slide1.xml' not in self.slides: return self.R(1, '主题+标题字体', s, full, d)

        # 字体/字号
        s1_xml = self.slides['ppt/slides/slide1.xml']['xml']
        if '华文行楷' in s1_xml: s += 4; d.append('标题华文行楷 OK')
        if 'sz="4000"' in s1_xml: s += 3; d.append('标题40号 OK')
        if '华文楷体' in s1_xml: s += 2; d.append('副标题华文楷体 OK')
        if 'sz="2000"' in s1_xml: s += 2; d.append('副标题20号 OK')
        return self.R(1, '主题+标题字体', s, full, d)

    # R2: 幻灯片2版式+图片+标题 (15)
    def r2(self):
        d = []; s = 0; full = 15
        if 'ppt/slides/slide2.xml' not in self.slides: return self.R(2, '两栏+图片+标题', s, full, d)
        s2_xml = self.slides['ppt/slides/slide2.xml']['xml']

        layout = self.get_slide_layout('ppt/slides/slide2.xml')
        if layout == 'slideLayout4': s += 4; d.append('两栏内容版式 OK')

        img = self.has_image('ppt/slides/slide2.xml')
        if img: s += 6; d.append(f'插入图片({img}) OK')
        else: d.append('未找到图片')

        if '万马奔腾' in s2_xml: s += 5; d.append('标题"万马奔腾" OK')
        return self.R(2, '两栏+图片+标题', s, full, d)

    # R3: 图片样式+动画 (15)
    def r3(self):
        d = []; s = 0; full = 15
        if 'ppt/slides/slide2.xml' not in self.slides: return self.R(3, '图片样式+动画', s, full, d)
        s2_xml = self.slides['ppt/slides/slide2.xml']['xml']

        if 'softEdge' in s2_xml: s += 5; d.append('柔化边缘(softEdge) OK')
        else: d.append('未找到柔化边缘')

        # 动画: animEffect + anim (浮入)
        if 'animEffect' in s2_xml: s += 5; d.append('动画效果(animEffect) OK')
        if '<p:anim' in s2_xml: s += 5; d.append('动画(浮入) OK')
        return self.R(3, '图片样式+动画', s, full, d)

    # R4: 幻灯片编号+页脚 (15)
    def r4(self):
        d = []; s = 0; full = 15
        all_slides = [self.slides[k]['xml'] for k in sorted(self.slides.keys(), key=slide_no)]

        num_count = sum(1 for sl in all_slides if 'sldNum' in sl and 'slidenum' in sl)
        ftr_count = sum(1 for sl in all_slides if '我的草原我的马' in sl)

        if num_count >= self.get_slide_count():
            s += 7; d.append(f'幻灯片编号({num_count}/{self.get_slide_count()}) OK')
        elif num_count > 0:
            s += 4; d.append(f'部分编号({num_count}/{self.get_slide_count()})')

        if ftr_count >= self.get_slide_count():
            s += 8; d.append(f'页脚"我的草原我的马"({ftr_count}/{self.get_slide_count()}) OK')
        elif ftr_count > 0:
            s += 4; d.append(f'部分页脚({ftr_count}/{self.get_slide_count()})')
        return self.R(4, '编号+页脚', s, full, d)

    # R5: 幻灯片3版式+移动内容+删除第4张+标题 (20)
    def r5(self):
        d = []; s = 0; full = 20
        if 'ppt/slides/slide3.xml' not in self.slides: return self.R(5, '移动+删除+标题', s, full, d)
        s3_xml = self.slides['ppt/slides/slide3.xml']['xml']
        s3_root = self.slides['ppt/slides/slide3.xml']['root']

        layout = self.get_slide_layout('ppt/slides/slide3.xml')
        if layout == 'slideLayout4': s += 4; d.append('两栏内容版式 OK')

        if '草原生态补奖政策' in s3_xml: s += 6; d.append('标题"草原生态补奖政策" OK')

        # 有右侧内容(ph idx=14)
        has_right = 'idx="14"' in s3_xml
        if has_right: s += 5; d.append('右侧内容(从第4张移动) OK')

        # 第4张已删除
        if self.get_slide_count() == 3 and 'ppt/slides/slide4.xml' not in self.slides:
            s += 5; d.append('第4张删除 OK')
        return self.R(5, '移动+删除+标题', s, full, d)

    # R6: 切换效果+放映方式 (20)
    def r6(self):
        d = []; s = 0; full = 20
        all_slides = [self.slides[k]['xml'] for k in sorted(self.slides.keys(), key=slide_no)]

        wipe_count = sum(1 for sl in all_slides if '<p:wipe' in sl)
        if wipe_count >= self.get_slide_count():
            s += 10; d.append(f'切换擦除({wipe_count}/{self.get_slide_count()}) OK')
        elif wipe_count > 0:
            s += 5; d.append(f'部分擦除({wipe_count}/{self.get_slide_count()})')

        if '<p:browse' in self.pres_props_xml:
            s += 10; d.append('放映方式:观众自行浏览 OK')
        else:
            d.append('未找到放映方式设置')
        return self.R(6, '切换+放映方式', s, full, d)


def main():
    if sys.platform == 'win32':
        try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except: pass
    base = os.path.dirname(os.path.abspath(__file__))
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base, '1结果.pptx')
    print(f'评分: {os.path.basename(path)}')
    print('=' * 50)
    g = PPTGrader1(path)
    results = g.grade()
    for r in results[:-1]:
        icon = 'PASS' if r['st'] == 'PASS' else 'FAIL'
        print(f"[{icon}] {r['req']:2d}. {r['name']:<18s}  {r['s']:5.1f} / {r['full']}")
        for det in r['d']: print(f'       {det}')
    print(f'{"=" * 50}\n  总分: {results[-1]["s"]} / 100')

if __name__ == '__main__':
    main()
