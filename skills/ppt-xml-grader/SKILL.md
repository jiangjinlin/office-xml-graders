---
name: PPT-XML-评分器
description: 通过解析.pptx文件的底层Open XML结构，对PowerPoint 2010操作题进行精准评分。支持主题、版式、字体/字号、图片插入、图片样式、动画效果、幻灯片编号/页脚、幻灯片切换、放映方式、幻灯片移动/删除等检测。
agent_created: true
---

# PPT XML 评分器

通过直接解析.pptx文件内部的Open XML结构，精准检测PPT文档格式和内容，实现操作题自动评分。

## 权威参考（查结构先看这里）

- 规范：ECMA-376 Part 1 §19（PresentationML，幻灯片/版式/母版/切换）、
  §21（DrawingML，文字、颜色、填充、形状）。字号 `a:rPr@sz` 在 §21.1.2.3.7。
- MS 文档：https://learn.microsoft.com/zh-cn/office/open-xml/presentation/overview
  SDK 类名即元素名：`SlideIdList`=`p:sldIdLst`、`Transition`=`p:transition`。
- ooxml.org：https://ooxml.org/drawingml/ 有 142 个 DrawingML 元素页（含 a:latin/a:ea/a:solidFill）；
  PresentationML 页面还是 Coming soon，PPT 主结构照 MS 文档和 §19。
- 本地速查表：`student_exam_app/tools/reference/OOXML-速查.md`。

## 核心能力

**PPTX 结构：** ZIP解压 → presentation.xml(幻灯片列表) → slideMasters(母版) → slides(幻灯片) → _rels(关系映射)

**核心数据源：**
- `presentation.xml` → 幻灯片列表、幻灯片尺寸、放映方式
- `slideMaster1.xml` → 母版页脚/编号设置
- `slideN.xml` → 每张幻灯片的内容(文字/图片/效果/动画/切换)
- `slideN.xml.rels` → 图片引用映射(rId → media/imageN.jpeg)
- `theme1.xml` → 主题名称
- `presProps.xml` → 放映属性(browse/kiosk)

## 考点覆盖表（按题库 10 套实测）

| 考点 | XML 证据 | 判定要点 | 已有实现 |
| --- | --- | --- | --- |
| 主题 | `ppt/theme/theme1.xml` 的 `a:clrScheme`/`a:fontScheme` | 主题名在 `theme@name`；改配色看 `a:clrScheme` 的 `a:accent1..6` | ppt1-10 |
| 母版/版式 | `ppt/slideMasters/*.xml`、`ppt/slideLayouts/*.xml`、`p:sldLayoutIdLst` | 版式归属看 `sldLayout` 的 rels；母版文字样式在 `p:txStyles` | ppt1-10 |
| 幻灯片增删/顺序 | `ppt/presentation.xml` 的 `p:sldIdLst` + `p:sldId@r:id` | 顺序以 sldIdLst 为准，文件名不可信；删除=该 sldId 不在列表且 rels 无引用 | ppt4/9 |
| 动画 | `p:timing` → `p:animEffect@transition="in"` + `p:cBhvr` | 滤镜形如 `barn(inVertical)`、`fade(in)`；"进入/退出"看 transition 值 | ppt1-9 |
| 动画顺序/计时 | `p:animEffect` 的 `p:cTn@dur`、`p:seq` | "自动播放"看 `p:cTn@indefinite="1"` 或 `p:cond` 事件类型 | ppt3/5 |
| 切换效果 | `p:transition` 下的 `p:blinds@dir`、`p:dissolve`、`p14:fade` | 百叶窗=`p:blinds`，溶解=`p:dissolve`；2010 有的效果在 `p14:` 命名空间 | ppt1-9 |
| 切换声音/速度 | `p:transition@spd="slow"` + `p:sndAc` | 速度 slow/med/fast；声音在 `p:sndAc/p:stSnd` 引用 rels | ppt1-9 |
| 放映方式 | `p:showPr@showType`（present/kiosk/browse） | 观众自行浏览 = `showType="browse"`，在展台浏览 = `kiosk` | ppt1-9 |
| 备注 | `ppt/notesSlides/notesSlideN.xml` 的 `a:t` | 备注是独立部件，不能从 slide xml 找 | ppt4/6 |
| 动作按钮/超链接 | `a:hlinkClick@r:id` 或 `p:sldJmp`；按钮形状 `a:prstGeom@prst="actionButton..."` | 跳到某页看 `p:sldJmp` 或 hlink 指向 slide rels | ppt6 |
| 音频/视频 | `p:pic` 里的 `a:videoFile@r:link` / `a:audioFile`，或 `p:timing` 的 `p:audio` | 媒体文件在 `ppt/media/`，引用在 `p:nvPr` 的 media rels | ppt1-10 |
| 图片插入/尺寸 | `p:pic/p:spPr/a:xfrm/a:ext@cx/cy` | 单位 EMU，1 厘米 360000；"铺满整页"要 == 幻灯片尺寸（12192000×6858000，16:9） | ppt1-10 |
| 图片样式/柔化 | `a:prstGeom@prst`（如 `roundRect`/`softEdge` 效果在 `a:effectLst/a:softEdge@rad`） | "柔化边缘矩形"= `prstGeom prst="rect"` + `a:softEdge`，"圆角矩形"= `roundRect` | ppt1/6 |
| 图片裁剪 | `a:srcRect@l/t/r/b`（千分比） | 题库未考，新题按此写 | 待补 |
| 艺术字 | `a:prstTxWarp@prst` 或文字用 `a:gradFill` 渐变填充 | 题库判"艺术字"时多用渐变填充/轮廓做证据（见 ppt6） | ppt6/7 |
| 表格 | `a:tbl` 的 `a:tr/a:tc`、`a:tblPr@firstRow` | 表格样式看 `a:tblPr` 的 bandRow/firstRow 开关 | ppt1/10 |
| SmartArt | `ppt/diagrams/data*.xml` + `p:graphicFrame` 的 `dgm:relIds` | 文字在 diagrams 部件，slide xml 只有引用 | ppt4 |
| 图表 | `ppt/charts/chartN.xml`（`c:` 命名空间） | 与 Excel 图表同构，判类型/标题照样用 `c:` 元素 | ppt2/6 |
| 编号/项目符号 | `a:buAutoNum@type`、`a:buChar@char` | "1. 2. 3." = `buAutoNum arabicPeriod`；圆点是 `buChar "•"` | ppt1/2/6 |

## 动画/切换判定的三个坑

1. **动画在 `p:timing` 里，不在 `p:sp`**：同一形状可以挂多个动画（进入+强调+退出），
   要按 `p:cBhvr/p:tgtEl/p:spTgt@spid` 定位到具体形状。
2. **`p:animEffect@filter` 带方向**：`barn(inVertical)` 是"垂直百叶窗"，
   只比 `barn` 会放过方向错的答案。
3. **切换 2010 有 `p14:` 扩展**：新版效果（如 `p14:reveal`、`p14:flythrough`）不在 `p:transition`
   主命名空间，扫字符串时要把 `p14:` 也算进去。

## 母版与继承的判定顺序

`slide → slideLayout → slideMaster → theme`，四级都要看：

- 占位符位置/大小：`p:sp/p:spPr/a:xfrm` 在 slide 里没有就往上找 layout、master。
- 字体：run 的 `a:rPr/a:latin|a:ea` → 段落 `a:defRPr` → 占位符 `p:txBody/a:lstStyle` →
  母版 `p:txStyles` → 主题 `a:fontScheme`。
- 颜色：`a:solidFill/a:srgbClr` → `a:schemeClr@val` 查主题 `clrScheme`。

`_core.pptcheck.resolved_shapes()` 已经做了这层合并，新评分器直接用它，别自己爬 XML。

## 新补的陷阱（2026-10-30，对照 ECMA-376 校验后）

- **中文字体在 `a:ea` 上**：`a:latin` 只管西文。题干说"设置为华文行楷"时，
  答案常常把字体挂在 `a:ea`（或 `a:cs`），只查 latin 会漏判。`_core.pptcheck` 的
  `resolved_shapes()` 已经把 run 级/段落级/继承的字体合并好了，优先用它。
- **字号是百分之一磅**：`sz="4000"` 是 40 磅，`sz="1800"` 是 18 磅；别按 pt 直接比。
- **放映顺序看 `p:sldIdLst` 再走 rels**：文件名 slide10.xml 不代表第 10 张，
  一律用 `slide_no()` 映射；排序用数字不能用字符串（slide10 会排到 slide2 前面）。
- **继承链是四级**：slide → slideLayout → slideMaster → theme。占位符位置、字号、
  颜色都可能来自母版 `p:txStyles` 或主题 `a:fontScheme`，只看 slide xml 会误判"没设置"。
- **动画看 `p:timing`**：`p:animEffect@transition="in"` + `p:cBhvr` 里的 `p:attrNameLst`
  才是"进入+滤镜"，`p:animEffect` 的 filter 值形如 `barn(inVertical)`。
- **切换效果 2010 有扩展**：常见 `p:blinds@dir`、`p:dissolve`、`p14:fade`；
  判"百叶窗"就是 `<p:blinds`，别去 `p14` 里找。

## 统一内核（优先用这套）

`student_exam_app/graders/_core/` 里的 `PptDeck` 已经处理好放映顺序、版式/母版/主题继承、切换和备注：

```python
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import PptDeck, Report

deck = PptDeck(path)
len(deck.slides), deck.slide_size, deck.theme['name'], deck.show_type

for slide in deck.slides:                       # 已是放映顺序，不是文件名顺序
    slide.index, slide.title, slide.text, slide.transition
    deck.layout_part(slide), deck.notes_text(slide), deck.image_targets(slide)
    for item in deck.resolved_shapes(slide):    # 坐标和字体已补全
        item['name'], item['placeholder'], item['left'], item['width']
        item['runs'][0]['font'], item['runs'][0]['size'], item['runs'][0]['bold']
```

`resolved_shapes()` 会依次从版式占位符、母版占位符、母版 `txStyles`（titleStyle/bodyStyle）、主题字体里补齐字号、加粗、字体和坐标。**直接读 `slide.xml` 里的 `a:rPr` 经常是空的**，因为格式继承自版式，这是 PPT 和 Excel/Word 最大的差别。

## PPTX XML 结构

```
文件名.pptx/
├── ppt/
│   ├── presentation.xml       ← 幻灯片列表 + 尺寸 + 放映方式
│   ├── presProps.xml           ← 放映属性 (browse/kiosk)
│   ├── theme/theme1.xml        ← 主题名称
│   ├── slideMasters/
│   │   └── slideMaster1.xml    ← 母版(页脚/编号/日期设置)
│   ├── slideLayouts/
│   │   └── slideLayoutN.xml    ← 版式定义
│   ├── slides/
│   │   ├── slide1.xml          ← 幻灯片内容
│   │   ├── slide2.xml
│   │   └── _rels/              ← 关系文件(图片/版式引用)
│   ├── media/
│   │   ├── image1.jpeg         ← 图片资源
│   │   └── image2.jpeg
│   └── _rels/
│       └── presentation.xml.rels ← 幻灯片/母版/主题映射
```

## 核心代码模板

```python
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

class PPTGrader:
    def __init__(self, path):
        self.path = path
        self.zf = None
        self.slides = {}  # {fname: {'root': Element, 'xml': str}}
        self.theme_name = ''
        self.all_xml = ''

    def load(self):
        self.zf = zipfile.ZipFile(self.path)
        self.namelist = self.zf.namelist()
        all_x = []

        # 1. Presentation XML
        d = self.zf.read('ppt/presentation.xml').decode('utf-8')
        all_x.append(d)
        self.pres_root = ET.fromstring(d)
        self.pres_xml = d

        # 2. Each slide
        for fname in self.namelist:
            if fname.startswith('ppt/slides/slide') and fname.endswith('.xml') \
               and '_rels' not in fname:
                d = self.zf.read(fname).decode('utf-8')
                all_x.append(d)
                self.slides[fname] = {
                    'root': ET.fromstring(d),
                    'xml': d
                }

        # 3. Theme
        if 'ppt/theme/theme1.xml' in self.namelist:
            d = self.zf.read('ppt/theme/theme1.xml').decode('utf-8')
            all_x.append(d)
            self.theme_name = ET.fromstring(d).get('name', '')

        # 4. Slide master
        if 'ppt/slideMasters/slideMaster1.xml' in self.namelist:
            d = self.zf.read('ppt/slideMasters/slideMaster1.xml').decode('utf-8')
            all_x.append(d)
            self.master_root = ET.fromstring(d)
            self.master_xml = d

        # 5. PresProps
        if 'ppt/presProps.xml' in self.namelist:
            d = self.zf.read('ppt/presProps.xml').decode('utf-8')
            all_x.append(d)
            self.pres_props_xml = d

        self.all_xml = '\n'.join(all_x)

    def get_slide_count(self):
        return len(self.slides)

    def get_slide_layout(self, slide_fname):
        """Get layout type from slide's _rels"""
        rel_fname = slide_fname.replace('.xml', '.xml.rels').replace('slides/', 'slides/_rels/')
        if rel_fname in self.namelist:
            rels_xml = ET.fromstring(self.zf.read(rel_fname).decode('utf-8'))
            for rel in rels_xml.iter(f'{{{PKG}}}Relationship'):
                if 'slideLayout' in rel.get('Type', ''):
                    target = rel.get('Target', '')
                    # Extract layout name: '../slideLayouts/slideLayout4.xml'
                    return target.split('/')[-1].replace('.xml', '')
        return ''

    def get_text(self, slide_root):
        """Extract all text runs from a slide with font info"""
        texts = []
        for sp in slide_root.iter(f'{{{A}}}sp'):
            for p in sp.iter(f'{{{A}}}p'):
                for r in p.iter(f'{{{A}}}r'):
                    t = r.find(f'{{{A}}}t')
                    rPr = r.find(f'{{{A}}}rPr')
                    if t is not None and t.text:
                        text = t.text
                        info = {}
                        if rPr is not None:
                            ea = rPr.find(f'{{{A}}}ea')
                            latin = rPr.find(f'{{{A}}}latin')
                            sz = rPr.get('sz', '')
                            b = rPr.find(f'{{{A}}}b')
                            info['font'] = (ea or latin or {}).get('typeface', '')
                            info['size'] = int(sz) // 100 if sz else 0
                            info['bold'] = b is not None
                        texts.append({'text': text, **info})
        return texts

    def has_image(self, slide_fname):
        """Check if slide has an image"""
        sld_xml = self.slides[slide_fname]['xml']
        if 'r:embed' in sld_xml or 'blip' in sld_xml:
            rel_fname = slide_fname.replace('.xml', '.xml.rels').replace('slides/', 'slides/_rels/')
            if rel_fname in self.namelist:
                rels_xml = ET.fromstring(self.zf.read(rel_fname).decode('utf-8'))
                for rel in rels_xml.iter(f'{{{PKG}}}Relationship'):
                    if 'image' in rel.get('Type', ''):
                        return rel.get('Target', '')
        return ''
```

## 常见考题检测模式

### 1. 主题
```python
# 主题名在 theme1.xml 的 name 属性
if self.theme_name == '奥斯汀': score += 3
elif '奥斯汀' in self.theme_name: score += 3
```

### 2. 幻灯片版式
```python
# 版式通过 slide _rels 中的 slideLayout 引用确定
layout = self.get_slide_layout('ppt/slides/slide1.xml')
# 标题幻灯片 → slideLayout1
# 两栏内容 → slideLayout4
if layout == 'slideLayout1': score += 2  # 标题幻灯片
elif layout == 'slideLayout4': score += 2  # 两栏内容
```

### 3. 字体/字号 (⚠️ 字号×100)
```python
# PPT字号存储为实际值×100: 40pt → sz="4000"
texts = self.get_text(slide_root)
for t in texts:
    if t['font'] == '华文行楷': score += 2
    if t['size'] == 40: score += 1  # sz=4000
    if t['size'] == 20: score += 1  # sz=2000
```

### 4. 图片插入
```python
img = self.has_image('ppt/slides/slide2.xml')
if img: score += 3  # 包含壮美草原.jpg等
```

### 5. 图片样式 (柔化边缘矩形)
```python
sld_xml = self.slides['ppt/slides/slide2.xml']['xml']
# 柔化边缘 = softEdge
if 'softEdge' in sld_xml: score += 2
# 矩形 = prstGeom prst="rect"
if 'prstGeom' in sld_xml and 'rect' in sld_xml: score += 1
```

### 6. 动画效果
```python
# 浮入 = animEffect + anim
sld_xml = self.slides['ppt/slides/slide2.xml']['xml']
if 'animEffect' in sld_xml and 'anim' in sld_xml: score += 3
```

### 7. 幻灯片编号 + 页脚
```python
# 编号: slide XML中 ph type="sldNum" + fld type="slidenum"
if 'sldNum' in sld_xml and 'slidenum' in sld_xml: score += 2
# 页脚: ph type="ftr" + 文本
if '我的草原我的马' in sld_xml: score += 2
```

### 8. 幻灯片切换
```python
# 切换在 slide XML的 <p:transition> 中
if '<p:wipe' in sld_xml: score += 2  # 擦除
```

### 9. 放映方式
```python
# 观众自行浏览 = presProps.xml 中有 <p:browse/>
if '<p:browse' in self.pres_props_xml: score += 2
```

### 10. 幻灯片数量/移动/删除
```python
# 原4张删为3张
if self.get_slide_count() == 3: score += 3
```

## 陷阱：幻灯片顺序与继承

### `sorted(slides.keys())` 会排错

字符串排序结果是 slide1、slide10、slide2……超过 9 张就乱。题目里"第 3 张幻灯片"指的是放映顺序，不是文件名。

```python
# ❌ 字符串排序
slide_list = sorted(self.slides.keys())

# ✅ 数字排序（_core.slide_no）
from _core import slide_no
slide_list = sorted(self.slides.keys(), key=slide_no)

# ✅ 更准：直接用放映顺序
for slide in deck.slides:  # 已按 presentation.xml 的 sldIdLst 展开
    ...
```

学生移动过幻灯片后，`slide2.xml` 可能已经排到第 5 位。凡是按"第 N 张"判分的题，必须走放映顺序。

### 占位符的字体字号通常在版式里

标题占位符在 `slide2.xml` 里往往只有 `<p:ph type="title"/>` 和文字，字号继承自 `slideLayout2.xml`，再不行继承 `slideMaster1.xml` 的 `p:txStyles/titleStyle`，再不行才是主题字体。

```python
# ❌ 直接读 slide 里的 rPr，size 全是 0
for sh in slide.shapes:
    sh.runs[0].size

# ✅ 走继承解析
for item in deck.resolved_shapes(slide):
    item['runs'][0]['size'], item['runs'][0]['font'], item['runs'][0]['bold']
```

### 项目符号可能是 `buNone`

段落没写 `a:buChar` 不代表没项目符号——默认继承版式；反过来，删掉项目符号后会有显式 `<a:buNone/>`。判"取消项目符号"要查 `buNone`，"加箭头符号"要查 `buChar char="Ø"`。

## 字号对照表

| PPT字号(px) | XML sz值 | 说明 |
|------------|---------|------|
| 40pt | 4000 | 华文行楷标题 |
| 20pt | 2000 | 华文楷体副标题 |
| 16pt | 1600 | 常见正文 |

## 检查清单

- [ ] 幻灯片数量是否正确
- [ ] 每张幻灯片的版式(layout reference)
- [ ] 标题/副标题文本内容
- [ ] 字体(typeface)和字号(sz/100)
- [ ] 图片是否存在(_rels → image)
- [ ] 图片效果(effectLst → softEdge)
- [ ] 动画(animEffect/anim)
- [ ] 切换效果(transition → wipe等)
- [ ] 页脚文本(ph type=ftr)
- [ ] 幻灯片编号(ph type=sldNum + fld type=slidenum)
- [ ] 主题名称(theme name属性)
- [ ] 放映方式(presProps → browse/kiosk)
- [ ] 幻灯片删除(比较源文件和答案文件数量)
- [ ] 第 N 张 → 按放映顺序，不用 `sorted(slides.keys())` 字符串排序
- [ ] 字体字号 → `deck.resolved_shapes()`，占位符格式继承自版式/母版
- [ ] 取消项目符号 → 查 `a:buNone`

---

## 回归与部署

```bash
python tools/selftest.py     # 全部评分器对标准答案跑一遍，必须全 100
python tools/core_sweep.py   # 核心库遍历 10 套 PPT 答案，确认无解析异常
python deploy.py --dry-run   # 发布前预检 + 打包
```

评分器满分不一定是 100（有的按 15 分/题给），`tools/selftest.py` 和 `app.py` 都按"实际得分 ÷ 满分"折算百分比，不用手动调分。

---

**技能创建：** 2026-07-03  
**版本：** 1.6（统一内核 + 顺序/继承陷阱 + 回归部署）  
**更新：** 新增 `scripts/ppt_analyze.py`（对比素材vs答案）+ `scripts/ppt_grader_template.py`（评分器骨架）
- 1.6：`_core.PptDeck` 用法（放映顺序、`resolved_shapes` 继承解析）；新增顺序/继承/项目符号三个陷阱；回归与部署命令

## 新增检测模式 (v1.1)

### 11. 图片位置和尺寸 (⚠️ EMU单位)

```python
# EMU转厘米: 1cm = 360000 EMU
off_matches = re.findall(r'<a:off x="(\d+)" y="(\d+)"', sld_xml)
ext_matches = re.findall(r'<a:ext cx="(\d+)" cy="(\d+)"', sld_xml)
for i, (ox, oy) in enumerate(off_matches):
    if ox == '0' and oy == '0': continue  # skip group xfrm
    cx, cy = ext_matches[i]
    x_cm = int(ox) / 360000  # → 3.5cm
    y_cm = int(oy) / 360000  # → 5.0cm
    w_cm = int(cx) / 360000  # → 18cm
    h_cm = int(cy) / 360000  # → 12cm
    if abs(x_cm - 3.5) < 0.5: score += 4
```

### 12. 切换效果 (百叶窗 blinds)

```python
# 百叶窗 = <p:blinds/>
if '<p:blinds' in sld_xml: score += 2
# 其他常见: wipe(擦除), fade(淡出), push(推进), cover(覆盖)
```

### 13. 项目符号

```python
# 箭头项目符号: buChar="Ø" 或其他Unicode字符
bullet_chars = re.findall(r'<a:buChar char="([^"]+)"', sld_xml)
if bullet_chars: score += 5  # 有项目符号
# buNone → 无项目符号, buAutoNum → 自动编号
```

### 14. 常用版式映射

```python
LAYOUT_MAP = {
    'slideLayout1': '标题幻灯片',
    'slideLayout2': '标题和内容',
    'slideLayout4': '两栏内容',
    'slideLayout6': '仅标题',
    'slideLayout7': '空白',
}
```

## 检查清单 (v1.1 新增)

- [ ] 图片位置: <a:off> x/y → EMU/360000=cm
- [ ] 图片尺寸: <a:ext> cx/cy → EMU/360000=cm
- [ ] 切换类型: wipe/blinds/fade/push/cover
- [ ] 项目符号: buChar / buAutoNum / buNone
- [ ] 版式: slideLayout2(标题和内容), slideLayout6(仅标题), slideLayout7(空白)
- [ ] 劈裂动画: animEffect filter="barn(outHorizontal)"
- [ ] 页脚全覆盖: 统计每页footer文本出现次数
- [ ] 幻灯片新增: 对比源文件slide数量
- [ ] ⚠️ p14命名空间: shred过渡用`<p14:shred/>`(搜`shred`不限定前缀)
- [ ] 过渡声音: sndAc→stSnd→snd embed
- [ ] 备注: notesSlides/notesSlide1.xml(对应slide2)
- [ ] dissolve(溶解)切换: `<p:dissolve/>`
- [ ] 动画时序: nodeType="afterEffect"(上一动画之后)
- [ ] 字体继承: 可能在theme/master中(不在slide显式设置)

---

## 精确 XML 参考表（ISO/IEC 29500:2016）

### 幻灯片切换 完整类型
| 元素 | 中文 | 方向/参数 |
|------|------|----------|
| `<p:blinds/>` | 百叶窗 | `horz`/`vert` |
| `<p:checker/>` | 棋盘 | `horz`/`vert` |
| `<p:cover/>` | 覆盖 | `down`/`left`/`right`/`up`/`downLeft`/`downRight`/`leftDown`/`leftUp`/`rightDown`/`rightUp`/`upLeft`/`upRight` |
| `<p:cut/>` | 切出 | `thruBlk="1"` |
| `<p:fade/>` | 淡出 | `thruBlk="1"` |
| `<p:push/>` | 推进 | 同cover方向 |
| `<p:wipe/>` | 擦除 | 同cover方向 |
| `<p:wheel/>` | 轮子 | `spokes`=1~8 |
| `<p:zoom/>` | 缩放 | `in`/`out` |
| `<p:dissolve/>` | 溶解 | 无 |
| `<p:split/>` | 分割 | `in`/`out`, `orient=horz/vert` |
| `<p:strips/>` | 条纹 | `downLeft`/`downRight`/`upLeft`/`upRight` |
| `<p:clock/>` | 时钟 | 无 |
| `<p:shred/>` | 碎片 | ⚠️ `p14:` 命名空间 |
| `<p:cube/>` | 立方体 | `left`/`right` |
| `<p:doors/>` | 门 | `horz`/`vert` |

**属性**: `spd`(`slow`/`med`/`fast`), `advClick`, `advTm`(ms), `durm`(ms), `sndAc`

### 动画
| 属性 | 说明 | 示例 |
|------|------|------|
| `presetID` | 预设编号 | `1`=出现, `35`=玩具风车 |
| `presetClass` | 类别 | `entr`/`exit`/`emph` |
| `presetSubtype` | 子类型 | `0` |
| `nodeType` | 时序 | `afterEffect`/`withEffect`/`clickEffect` |
| `filter` | 过渡滤镜 | `fade` |

### 放映类型
| 元素 | 说明 | XML |
|------|------|-----|
| `browse` | 观众自行浏览 | `<p:browse/>` |
| `kiosk` | 展台浏览 | `<p:kiosk/>` |
| `present` | 演讲者放映 | `<p:present/>` |

### 幻灯片尺寸 `sldSz`
| 属性 | 说明 | 值 |
|------|------|-----|
| `cx` | 宽度 EMU | `9144000`(4:3), `12192000`(16:9) |
| `cy` | 高度 EMU | `6858000`(4:3) |
| `type` | 类型 | `screen4x3`, `screen16x9` |

### 页眉页脚
| 属性 | 说明 | 值 |
|------|------|-----|
| `sldNum` | 编号 | `"1"` |
| `ftr` | 页脚 | 占位符 `<a:t>` 文本 |

### DrawingML 文本 `a:rPr`
| 属性 | 说明 | 示例 |
|------|------|------|
| `lang` | 语言 | `zh-CN` |
| `sz` | 字号×100 | `2800`=28pt, `1800`=18pt |
| `b` | 加粗 | `1` |
| `i` | 斜体 | `1` |
| `<a:solidFill/>`→`<a:srgbClr val="FF0000"/>` | 颜色 | |

### 图片效果
| 元素 | 说明 | 检测 |
|------|------|------|
| `<a:softEdge/>` | 柔化边缘 | `softEdge` |
| `<a:glow/>` | 发光 | `glow` |
| `<a:shadow/>` | 阴影 | `shadow` |

### 单位换算
| 单位 | EMU | cm |
|------|-----|-----|
| 1 pt | 12700 | 0.035 |
| 1 cm | 360000 | 1 |
| 字号 | ×100 | sz="3200"=32pt |

### 检查清单（新增）
- [ ] 主题: `theme1.xml→name` 属性
- [ ] 版式: slide rels→slideLayout Target
- [ ] 字体: `a:rPr typeface` 或 theme `a:ea`/`a:latin`
- [ ] 动画声音: `sndAc→stSnd→snd` embed
- [ ] 动画速度: `durm="500"` 毫秒
- [ ] 幻灯片编号: `<p:sldNum/>` 存在 → 编号可见
- [ ] 打印设置: `horizontalCentered`/`verticalCentered` in printOptions

---

## 精确枚举参考（ISO/IEC 29500 深度分析）

### 过渡速度
| spd | 含义 |
|-----|------|
| `slow` | 慢 |
| `med` | 中(默认) |
| `fast` | 快 |

### 占位符类型 PlaceholderValues
| 值 | 含义 |
|----|------|
| `title`, `ctrTitle` | 标题 |
| `subTitle` | 副标题 |
| `body` | 正文 |
| `ftr` | 页脚 |
| `sldNum` | 幻灯片编号 |
| `dt` | 日期 |
| `pic`, `chart`, `tbl`, `clipArt`, `media` | 对象占位符 |

### 颜色方案 ColorSchemeIndexValues
`dk1`, `lt1`, `dk2`, `lt2`, `accent1`~`accent6`, `hlink`, `folHlink`

### 幻灯片尺寸
| type | cx | cy |
|------|-----|-----|
| `screen4x3` | 9144000 | 6858000 |
| `screen16x9` | 12192000 | 6858000 |
| `screen16x10` | 12192000 | 7620000 |

### 动画枚举
| presetClass | nodeType |
|-------------|----------|
| `entr` 进入 | `afterEffect` 上一动画之后 |
| `exit` 退出 | `withEffect` 同时 |
| `emph` 强调 | `clickEffect` 单击时 |
| `path` 路径 | `tmRoot`, `mainSeq`, `interactiveSeq` |

---

## 架构级参考（ISO/IEC 29500:2016 字符级）

### CT_GroupShape / CT_Shape 属性
| 属性 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `name` | ST_String | — | 形状名称 |
| `id` | ST_DrawingElementId | — | 唯一ID |
| `hidden` | ST_Boolean | `false` | 隐藏 |
| `locked` | ST_Boolean | `false` | 锁定 |

### CT_TextBody 属性
| 属性 | 类型 | 说明 |
|------|------|------|
| `vert` (文字方向) | ST_TextVerticalType | `horz`(横排), `eaVert`(东亚竖排), `vert`(竖排), `vert270`(270度) |
| `wrap` (自动换行) | ST_TextWrappingType | `none`, `square` |
| `lIns` (左内边距) | ST_Coordinate32 | EMU |
| `rIns` (右内边距) | ST_Coordinate32 | EMU |
| `tIns` (上内边距) | ST_Coordinate32 | EMU |
| `bIns` (下内边距) | ST_Coordinate32 | EMU |

### SlideSizeValues 完整枚举
| 枚举值 | CX (EMU) | CY (EMU) | 比例 |
|---------|----------|----------|------|
| `Screen4x3` | 9,144,000 | 6,858,000 | 4:3 |
| `Screen16x9` | 12,192,000 | 6,858,000 | 16:9 |
| `Screen16x10` | 12,192,000 | 7,620,000 | 16:10 |
| `Widescreen` | 12,192,000 | 6,858,000 | 16:9 |
| `OnScreenShow4x3` | 9,144,000 | 6,858,000 | 4:3 |
| `OnScreenShow16x9` | 12,192,000 | 6,858,000 | 16:9 |
| `B4IsoPaper` | 10,287,360 | 14,572,800 | B4 |
| `B5IsoPaper` | 8,505,360 | 12,002,400 | B5 |
| `Letter` | 9,144,000 | 11,001,600 | 8.5×11 |
| `A4` | 8,644,080 | 11,892,240 | A4 |
| `Banner` | 17,640,000 | 2,160,000 | 横幅 |
| `Ledger` | 11,001,600 | 15,379,200 | 11×17 |
| `Custom` | 可变 | 可变 | 自定义 |

### PlaceholderValues 完整枚举
| 枚举值 | XML 值 | 说明 |
|---------|--------|------|
| `Title` | `title` | 标题 |
| `Body` | `body` | 正文 |
| `CenteredTitle` | `ctrTitle` | 居中标题 |
| `SubTitle` | `subTitle` | 副标题 |
| `Date` | `dt` | 日期 |
| `Footer` | `ftr` | 页脚 |
| `SlideNumber` | `sldNum` | 编号 |
| `Header` | `hdr` | 页眉 |
| `Object` | `obj` | 对象 |
| `Chart` | `chart` | 图表 |
| `Table` | `tbl` | 表格 |
| `ClipArt` | `clipArt` | 剪贴画 |
| `Diagram` | `dgm` | 图示 |
| `Media` | `media` | 媒体 |
| `Picture` | `pic` | 图片 |
| `SmartArt` | `sldImg` | SmartArt |

### 形状位置 (xfrm) 属性
| 属性 | 单位 | 说明 |
|------|------|------|
| `off@x` | EMU | 水平偏移（左上角） |
| `off@y` | EMU | 垂直偏移（左上角） |
| `ext@cx` | EMU | 宽度 |
| `ext@cy` | EMU | 高度 |
| `rot` | 1/60000 度 | 旋转角度（正=顺时针） |

### EMU 换算
| 源 | → EMU | → cm |
|----|-------|------|
| 1 inch | 914400 | 2.54 |
| 1 pt | 12700 | 0.0353 |
| 1 cm | 360000 | 1.0 |
| 字号×100 | — | `"2800"` = 28pt → 28×12700 EMU |

## 判分点全表（含题库未考但可能出现的）

> 完整版（三科合并）在 `student_exam_app/tools/reference/判分点全表.md`；这里只放本科目。新题先对表，再写代码。

### 演示文稿与幻灯片

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `p:sldMasterIdLst` / `p:sldIdLst` | 母版/幻灯片列表 | 顺序以 sldIdLst 为准；`p:sldId@id` 是内部 id |
| `p:sldSz@cx/cy` | 幻灯片尺寸 | 16:9 = 12192000×6858000 EMU；4:3 = 9144000×6858000 |
| `p:notesSz` | 备注页尺寸 | 与幻灯片尺寸分开 |
| `p:defaultTextStyle` | 默认文本样式 | 新建文本框继承它 |
| `p:sectLst`（p14） | 节 | 2010 的节在扩展命名空间 |
| `p:sldId@show="0"` | 隐藏幻灯片 | 放映时跳过 |
| `p:cSld@name` | 幻灯片名称 | 大纲视图里的名字 |

### 形状与文本

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `p:sp` / `p:pic` / `p:graphicFrame` / `p:grpSp` / `p:cxnSp` | 形状种类 | 图片是 pic、表格图表是 graphicFrame、组合是 grpSp |
| `p:nvSpPr/p:cNvPr@name/id` | 形状名与 id | 动画按 spid 关联，名字重复会撞 |
| `p:spPr/a:xfrm/a:off|x` / `a:ext@cx/cy` / `@rot` | 位置尺寸旋转 | EMU；`rot` 单位 1/60000 度（90°=5400000） |
| `a:prstGeom@prst` / `a:custGeom` | 预设/自定义形状 | roundRect、ellipse、actionButton*、flowChart* |
| `a:solidFill/a:srgbClr|schemeClr` | 填充 | `schemeClr@val="accent1"` 要查主题 |
| `a:gradFill` + `a:gsLst` | 渐变 | 艺术字常用；判"两种颜色渐变"看 gs 数量 |
| `a:ln@w` + `a:solidFill` | 轮廓 | 宽度 EMU |
| `a:effectLst/a:softEdge@rad` | 柔化边缘 | 题库"柔化边缘矩形" |
| `a:effectLst/a:outerShdw` / `a:glow` | 阴影/发光 | 图片样式类考点 |
| `a:blipFill/a:blip@r:embed` | 图片引用 | media 文件 |
| `a:srcRect` | 裁剪 | 千分比 |
| `p:txBody/a:bodyPr@anchor/wrap/vert/anchorCtr` | 文本框属性 | 垂直居中 `anchor="ctr"`；竖排 `vert="vert270"` |
| `a:pPr@algn/marL/lvl/indent` | 段落 | 居中 `algn="ctr"`；层级 `lvl` |
| `a:buChar@char` / `a:buAutoNum@type` / `a:buNone` | 项目符号/编号 | "1. 2. 3."=`arabicPeriod`；`buNone` 是取消 |
| `a:lnSpc/a:spcPct@val` | 行距 | 100000=1 倍，120000=1.2 倍；`spcPts` 是固定磅值 |
| `a:spcBef` / `a:spcAft` | 段前段后 | `spcPts` 磅或 `spcPct` |
| `a:rPr@sz/b/i/u` + `a:latin/a:ea/a:cs` | 字号字体 | 字号×100；中文在 `a:ea` |
| `a:prstTxWarp@prst` | 艺术字变形 | textArchUp 等 |

### 表格、图表、SmartArt、媒体

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `a:tbl` + `a:tblPr@firstRow/bandRow` + `a:tblGrid/a:gridCol@w` | 表格 | 行高 `a:tr@h`；列宽在 gridCol |
| `p:graphicFrame` + `c:chart@r:id` | 图表 | 图表 XML 在 `ppt/charts/chartN.xml`，与 Excel 同构 |
| `dgm:relIds` | SmartArt | 内容在 `ppt/diagrams/dataN.xml` |
| `a:videoFile@r:link` / `a:audioFile` | 音视频 | 引用在 rels，文件在 `ppt/media/` |
| `p:nvPr/p:media` | 媒体关系 | 与 timing 配合做自动播放 |

### 动画、切换、放映

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `p:timing/p:tnLst/p:par/p:seq` | 动画序列 | 一个形状多个动画按 `p:cBhvr/p:tgtEl/p:spTgt@spid` 定位 |
| `p:animEffect@transition/filter` | 进入/退出效果 | `filter="barn(inVertical)"` 带方向；"擦除"=`wipe`、"飞入"=`slide(in)` |
| `p:anim@calcmode` / `p:animMotion` / `p:animScale` / `p:animRot` | 强调/动作路径/缩放旋转 | "按方框形缩小"=`animScale` 或 `animEffect filter="box(out)"` |
| `p:cTn@dur` / `@repeatCount` / `p:cond` | 时长/重复/触发 | "自动播放"=`p:cond delay="0"`；"单击时"=`evt="onClick"` |
| `p:transition@spd` + `p:blinds|p:dissolve|p:fade|p14:*` | 切换效果 | 百叶窗=`p:blinds`；速度 slow/med/fast |
| `p:sndAc` | 切换声音 | `p:stSnd` 引用 media |
| `p:showPr@showType` | 放映方式 | `present`/`browse`（观众自行浏览）/`kiosk`（在展台浏览） |
| `p:showPr@loop/useTimings/showAnimation` | 循环/使用计时 | 自动放映类考点 |

### 母版、版式、主题、备注

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `p:sldLayoutIdLst`（母版里） | 版式清单 | 版式名 `p:cSld@name`；`slideLayout4` 常是"两栏内容" |
| `p:txStyles/p:titleStyle|p:bodyStyle|p:otherStyle` | 母版文字样式 | 继承链第三级 |
| `theme/a:clrScheme` | 主题配色 | `dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink` |
| `theme/a:fontScheme@name` | 主题字体 | majorFont/minorFont 各含 latin/ea/cs |
| `theme/a:fmtScheme` | 主题效果 | 填充/线条/效果样式列表 |
| `p:bg` / `p:bgRef` | 背景 | 背景图 `a:blipFill`；`bgRef` 引主题样式 |
| `notesSlide/a:t` | 备注文字 | 独立部件，不在 slide |
| `handoutMaster` | 讲义母版 | 题库少见 |
| 超链接 `a:hlinkClick@r:id/action` | 链接/动作 | `action="ppaction://hlinksldjump"` 是跳页；`ppaction://program` 是运行程序 |

