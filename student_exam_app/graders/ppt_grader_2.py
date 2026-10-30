"""
PPT 2010 评分器 - 文件夹2 (v1)

题干(7项, 评分点):
1. 主题"活力" + Slide1仅标题+标题文本32号
2. 图片插入+位置+尺寸(12×18cm, 3.5×5cm)
3. 图片旋转动画+百叶窗切换
4. Slide2空白版式+文字移动
5. Slide3标题和内容+箭头项目符号+标题"AI构想"
6. 放映观众自行浏览
"""
import zipfile, os, sys, io, xml.etree.ElementTree as ET, re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import slide_no

P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

class PPTGrader2:
    def __init__(self, path):
        self.path = path; self.zf = None
        self.slides = {}; self.all_xml = ''
        self.theme_name = ''; self.pres_props_xml = ''

    def load(self):
        self.zf = zipfile.ZipFile(self.path); self.namelist = self.zf.namelist()
        all_x = []
        d = self.zf.read('ppt/presentation.xml').decode('utf-8'); all_x.append(d)
        for fname in self.namelist:
            if fname.startswith('ppt/slides/slide') and fname.endswith('.xml') and '_rels' not in fname:
                d = self.zf.read(fname).decode('utf-8'); all_x.append(d)
                self.slides[fname] = {'root': ET.fromstring(d), 'xml': d}
        if 'ppt/theme/theme1.xml' in self.namelist:
            d = self.zf.read('ppt/theme/theme1.xml').decode('utf-8'); all_x.append(d)
            self.theme_name = ET.fromstring(d).get('name', '')
        if 'ppt/slideMasters/slideMaster1.xml' in self.namelist:
            d = self.zf.read('ppt/slideMasters/slideMaster1.xml').decode('utf-8'); all_x.append(d)
        if 'ppt/presProps.xml' in self.namelist:
            d = self.zf.read('ppt/presProps.xml').decode('utf-8'); all_x.append(d)
            self.pres_props_xml = d
        self.all_xml = '\n'.join(all_x)

    def get_slide_count(self): return len(self.slides)

    def get_slide_layout(self, fname):
        name = fname.split('/')[-1]
        rf = f'ppt/slides/_rels/{name}.rels'
        if rf in self.namelist:
            for rel in ET.fromstring(self.zf.read(rf).decode('utf-8')).iter(f'{{{PKG}}}Relationship'):
                if 'slideLayout' in rel.get('Type', ''):
                    return rel.get('Target', '').split('/')[-1].replace('.xml', '')
        return ''

    def has_image(self, fname):
        name = fname.split('/')[-1]; rf = f'ppt/slides/_rels/{name}.rels'
        if rf in self.namelist:
            for rel in ET.fromstring(self.zf.read(rf).decode('utf-8')).iter(f'{{{PKG}}}Relationship'):
                if 'image' in rel.get('Type', ''): return rel.get('Target', '')
        return ''

    def grade(self):
        self.load()
        rr = [self.r1(), self.r2(), self.r3(), self.r4(), self.r5(), self.r6()]
        t = round(sum(x['s'] for x in rr), 1)
        rr.append({'req': 0, 'name': '总分', 's': t, 'full': 100, 'd': [], 'st': 'PASS' if t >= 60 else 'FAIL'})
        return rr

    def R(self, req, name, score, full, detail):
        s = min(float(score), float(full))
        return {'req': req, 'name': name, 's': round(s,1), 'full': int(full), 'd': detail, 'st': 'PASS' if s >= full*.6 else 'FAIL'}

    # R1: 主题+Slide1仅标题+标题文本32号 (15)
    def r1(self):
        d = []; s = 0; full = 15
        if self.theme_name == '活力': s += 5; d.append('主题"活力" OK')
        elif '活力' in self.theme_name: s += 5; d.append(f'主题={self.theme_name}')

        if 'ppt/slides/slide1.xml' not in self.slides: return self.R(1, '主题+标题', s, full, d)
        layout = self.get_slide_layout('ppt/slides/slide1.xml')
        if layout == 'slideLayout6': s += 4; d.append('仅标题版式 OK')
        else: d.append(f'版式={layout}')

        s1 = self.slides['ppt/slides/slide1.xml']['xml']
        # Title text (multi-part)
        if '从预测进化' in s1 and '构想' in s1 and '新蛋白质结构' in s1:
            s += 3; d.append('标题文本 OK')
        if '3200' in s1: s += 3; d.append('字号32 OK')
        return self.R(1, '主题+标题', s, full, d)

    # R2: 图片+位置+尺寸 (20)
    def r2(self):
        d = []; s = 0; full = 20
        if 'ppt/slides/slide1.xml' not in self.slides: return self.R(2, '图片+位置+尺寸', s, full, d)
        img = self.has_image('ppt/slides/slide1.xml')
        if img: s += 6; d.append(f'图片({img}) OK')

        s1 = self.slides['ppt/slides/slide1.xml']['xml']
        # Position: x=3.5cm=1260000EMU, y=5cm=1800000EMU
        off_matches = re.findall(r'<a:off x="(\d+)" y="(\d+)"', s1)
        ext_matches = re.findall(r'<a:ext cx="(\d+)" cy="(\d+)"', s1)
        for i, (ox, oy) in enumerate(off_matches):
            # Skip the group xfrm (0,0)
            if ox == '0' and oy == '0': continue
            cx, cy = (ext_matches[i] if i < len(ext_matches) else ('0','0'))
            x_cm = int(ox)/360000; y_cm = int(oy)/360000
            w_cm = int(cx)/360000; h_cm = int(cy)/360000
            if abs(x_cm - 3.5) < 0.5: s += 4; d.append(f'水平3.5cm OK')
            if abs(y_cm - 5) < 0.5: s += 3; d.append(f'垂直5cm OK')
            if abs(w_cm - 18) < 1: s += 4; d.append(f'宽18cm OK')
            if abs(h_cm - 12) < 1: s += 3; d.append(f'高12cm OK')
            break
        return self.R(2, '图片+位置+尺寸', s, full, d)

    # R3: 图片旋转动画+百叶窗切换 (20)
    def r3(self):
        d = []; s = 0; full = 20
        if 'ppt/slides/slide1.xml' not in self.slides: return self.R(3, '动画+切换', s, full, d)
        s1 = self.slides['ppt/slides/slide1.xml']['xml']

        # 旋转动画
        if 'animEffect' in s1: s += 10; d.append('图片动画(animEffect) OK')
        elif 'anim' in s1: s += 5; d.append('有动画(非旋转)')

        # 百叶窗切换(所有幻灯片)
        slide_list = sorted(self.slides.keys(), key=slide_no)
        blinds_count = sum(1 for k in slide_list if '<p:blinds' in self.slides[k]['xml'])
        if blinds_count >= len(slide_list):
            s += 10; d.append(f'百叶窗切换({blinds_count}/{len(slide_list)}) OK')
        elif blinds_count > 0:
            s += 5; d.append(f'部分百叶窗({blinds_count})')
        return self.R(3, '动画+切换', s, full, d)

    # R4: Slide2空白版式+文字 (15)
    def r4(self):
        d = []; s = 0; full = 15
        if 'ppt/slides/slide2.xml' not in self.slides: return self.R(4, '空白版式+文字', s, full, d)
        layout = self.get_slide_layout('ppt/slides/slide2.xml')
        if layout == 'slideLayout7': s += 7; d.append('空白版式 OK')

        s2 = self.slides['ppt/slides/slide2.xml']['xml']
        texts2 = re.findall(r'<a:t[^>]*>([^<]+)</a:t>', s2)
        if layout == 'slideLayout7' and texts2 and len(''.join(texts2)) > 20:
            s += 8; d.append('文字内容 OK')
        return self.R(4, '空白版式+文字', s, full, d)

    # R5: Slide3标题和内容+箭头项目符号+标题 (15)
    def r5(self):
        d = []; s = 0; full = 15
        if 'ppt/slides/slide3.xml' not in self.slides: return self.R(5, '标题内容+符号', s, full, d)
        layout = self.get_slide_layout('ppt/slides/slide3.xml')
        if layout == 'slideLayout2': s += 5; d.append('标题和内容版式 OK')

        s3 = self.slides['ppt/slides/slide3.xml']['xml']

        # 标题（素材里本就有"构想"字样，必须配合版式改动才给分）
        if layout == 'slideLayout2' and '构想' in s3: s += 5; d.append('标题"AI构想" OK')

        # 项目符号
        has_any_bullet = re.findall(r'<a:buChar char="([^"]+)"', s3)
        has_buAutoNum = 'buAutoNum' in s3
        if has_any_bullet:
            s += 5; d.append(f'项目符号({has_any_bullet[0]}) OK')
        elif has_buAutoNum:
            s += 5; d.append('自动编号 OK')
        return self.R(5, '标题内容+符号', s, full, d)

    # R6: 放映观众自行浏览 (15)
    def r6(self):
        d = []; s = 0; full = 15
        if '<p:browse' in self.pres_props_xml:
            s += 15; d.append('观众自行浏览 OK')
        else:
            d.append('未找到放映设置')
        return self.R(6, '放映方式', s, full, d)


def main():
    if sys.platform == 'win32':
        try: sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        except: pass
    base = os.path.dirname(os.path.abspath(__file__))
    path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(base, '1结果.pptx')
    print(f'评分: {os.path.basename(path)}')
    print('=' * 50)
    g = PPTGrader2(path)
    results = g.grade()
    for r in results[:-1]:
        icon = 'PASS' if r['st'] == 'PASS' else 'FAIL'
        print(f"[{icon}] {r['req']:2d}. {r['name']:<18s}  {r['s']:5.1f} / {r['full']}")
        for det in r['d']: print(f'       {det}')
    print(f'{"=" * 50}\n  总分: {results[-1]["s"]} / 100')

if __name__ == '__main__':
    main()
