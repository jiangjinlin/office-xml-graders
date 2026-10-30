# Word Open XML 关键模式参考

本文档提供Word `.docx` 文件内部XML结构的关键模式，用于格式检测和评分。

---

## 一、文件结构

```
文件名.docx/
├── word/
│   ├── document.xml      ← 主文档
│   ├── styles.xml        ← 样式
│   ├── settings.xml      ← 文档设置
│   ├── footnotes.xml     ← 脚注
│   ├── header1.xml      ← 页眉（可能多个）
│   ├── footer1.xml      ← 页脚
│   ├── _rels/           ← 关系文件
│   │   ├── document.xml.rels
│   │   └── header3.xml.rels
│   └── media/           ← 嵌入图片
│       ├── image1.png
│       └── 水印.jpg
```

---

## 二、命名空间

```python
WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WPNS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
ANS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PICNS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
VNS = 'urn:schemas-microsoft-com:vml'  # 关键：水印使用VML
```

---

## 三、核心Helper函数

```python
def T(tag):
    """标签名转完整命名空间"""
    if ':' not in tag: return tag
    p, l = tag.split(':', 1)
    ns_map = {'w': WNS, 'wp': WPNS, 'a': ANS, 'pic': PICNS, 'r': RNS, 'v': VNS}
    if p in ns_map:
        return '{' + ns_map[p] + '}' + l
    return tag

def child(parent, tag):
    """查找直接子元素，兼容ElementTree"""
    for c in parent:
        if c.tag == tag: return c
    return None

def children(parent, tag):
    """查找所有直接子元素"""
    return [c for c in parent if c.tag == tag]

def GA(elem, attr, default=None):
    """安全获取属性，自动处理命名空间
    ⚠️ 关键：ns_map存裸URI（无大括号），用'%s%s'拼接！
    错误：'{%s}%s' % (W, 'attr') → '{{http://...}}attr'（双层括号！）
    正确：'%s%s' % ('http://...', 'attr') → '{http://...}attr'
    """
    if elem is None: return default
    if ':' in attr:
        ns, local = attr.split(':', 1)
        ns_uris = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'v': 'urn:schemas-microsoft-com:vml',
            'o': 'urn:schemas-microsoft-com:office:office',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        }
        return elem.get('%s%s' % (ns_uris.get(ns, ''), local), default)
    return elem.get(attr, default)

---

## 四、常用XML模式

### 4.1 字符格式

**位置**：`w:rPr`（Run Properties）

| 格式 | XML标签 | 关键属性 |
|--------|----------|------------|
| 字体 | `w:rFonts` | `w:eastAsia`（中文字体） |
| 加粗 | `w:b` | （存在即表示加粗） |
| 倾斜 | `w:i` | （存在即表示倾斜） |
| 字号 | `w:sz` | `w:val`（半磅为单位，44=二号22pt） |
| 颜色 | `w:color` | `w:val`（十六进制颜色） |
| 字符边框 | `w:bdr` | `w:val`（样式）、`w:sz`（线宽）、`w:color`（颜色） |
| 下划线 | `w:u` | `w:val`（样式） |

### 4.2 段落格式

**位置**：`w:pPr`（Paragraph Properties）

| 格式 | XML标签 | 关键属性 |
|--------|----------|------------|
| 对齐 | `w:jc` | `w:val`（left/center/right/justify） |
| 缩进 | `w:ind` | `w:leftChars`、`w:rightChars`（字符数）、`w:firstLineChars`（首行缩进） |
| 间距 | `w:spacing` | `w:beforeLines`（段前，100=1行）、`w:afterLines`（段后）、`w:line`（行距）、`w:lineRule`（fixed/exact/atLeast） |
| 分栏 | `w:cols` | `w:num`（栏数）、`w:sep`（分隔线）、`w:equalWidth` |
| 边框 | `w:pBorders` | `w:top`、`w:left`等 |

### 4.3 图片

**位置**：`w:drawing` 或 `w:pict` (VML)

| 属性 | XML路径 | 单位/说明 |
|--------|-----------|------------|
| 尺寸 | `wp:extent` | `cx`/`cy`（EMU，1cm=914400*2.54/100） |
| 环绕 | `wp:wrapTight`等 | wrapTight/wrapSquare/wrapNone等 |
| 旋转 | `a:xfrm` | `rot`（60000=1度） |
| 对齐 | `wp:positionH` | `relativeFrom`（column/page/margin）、`wp:align`（center/left/right） |

**注意**：旧版Word可能使用`w:pict`（VML）而非`w:drawing`（DML）。需要同时检查两者。

### 4.3.1 VML文本框（w:pict）

**位置**：`w:pict`中的VML形状

**VML形状类型速查（type属性值）：**

| `shape.get('type')` | 中文名称 | `shape.get('id')` 示例 | 用途 |
|---------------------|---------|----------------------|------|
| `#_x0000_t202` | 文本框 | "文本框 2" | 标题、文字标注 |
| `#_x0000_t63` | 椭圆批注 | "椭圆批注 1" | 气泡标注 |
| `#_x0000_t136` | 圆角矩形 | "矩形 N" | 流程图形状 |
| `#_x0000_t32` | 弧形 | — | 弧形箭头等 |
| `#_x0000_t75` | 五边形 | — | 流程图 |
| — | 水印 | id含"Watermark" | 只在header中出现 |

```python
# VML文本框检测
picts = root.findall('.//{' + WNS + '}pict')
for pict in picts:
    # 检查VML形状
    shape = pict.find('.//{' + VNS + '}shape')
    if shape is not None:
        # 边框颜色
        strokecolor = shape.get('strokecolor', '')  # 如 "red", "#ff0000"
        # 边框宽度
        strokeweight = shape.get('strokeweight', '')  # 如 "2pt"
        # 填充颜色
        fillcolor = shape.get('fillcolor', '')
    
    # 检查VML stroke元素
    stroke = pict.find('.//{' + VNS + '}stroke')
    if stroke is not None:
        color = stroke.get('color', '')
        weight = stroke.get('weight', '')
    
    # 检查VML fill元素
    fill = pict.find('.//{' + VNS + '}fill')
    if fill is not None:
        type = fill.get('type', '')  # "gradient" 表示渐变
        fillcolor2 = fill.get('fillcolor2', '')  # 渐变第二色
        opacity = fill.get('opacity', '')  # 透明度
```

### 4.4 节设置

**位置**：`w:sectPr`（可能在`w:body`末尾，或段落内的`w:pPr/w:sectPr`）

| 设置 | XML标签 | 说明 |
|--------|----------|------|
| 页面边框 | `w:pgBorders` | 包含`w:top`/`w:left`等子元素 |
| 脚注设置 | `w:footnotePr` | 包含`w:numFmt`（编号格式） |
| 背景 | `w:background` | 渐变或纯色填充 |
| 页眉页脚链接 | `w:headerReference`/`w:footerReference` | 通过`r:id`引用 |

### 4.5 水印（VML）

**位置**：`header*.xml`中的VML形状

```python
# 水印是VML形状，不在标准Open XML中
for shape in root.findall('.//{' + VNS + '}shape'):
    sid = shape.get('id', '')
    if 'Watermark' in sid:
        # 找到水印形状
        for imgdata in shape.findall('{' + VNS + '}imagedata'):
            blacklevel = imgdata.get('blacklevel', '')  # 冲蚀效果
            gain = imgdata.get('gain', '')
```

---

## 五、单位转换

| 单位 | 说明 | 转换 |
|--------|------|------|
| EMU | English Metric Unit | 1cm = 914400 * 2.54 / 100 |
| 半磅 | 字号单位 | 44 = 22pt = 二号 |
| 1/20字符 | 字符宽度 | 200 = 10字符 |
| 1/8磅 | 边框线宽 | 12 = 1.5磅 |
| 1/100行 | 段落间距 | 100 = 1行 |
| 60000/度 | 旋转角度 | 10800000 = 180度 |

---

## 六、常见问题排查

### 6.1 检测不到元素

- [ ] 命名空间是否正确？使用`T()`函数
- [ ] 路径是否需要递归？使用`X(elem, '//tag:name')`
- [ ] 元素是否在`w:pPr`中？某些格式在段落属性而非运行属性

### 6.2 属性获取失败

- [ ] 是否使用`GA()`而非`elem.get()`？
- [ ] 属性是否带命名空间？如`w:val`需写作`{http://...}val`

### 6.3 图片相关

- [ ] 图片是否在`w:drawing`中？旧版Word可能用`w:pict`
- [ ] 关系文件是否加载？图片引用通过`*rels`文件
- [ ] 单位是否正确转换？EMU转厘米需除以914400*2.54/100

### 6.4 水印检测

- [ ] 是否检查了VML形状？水印使用`v:shape`
- [ ] 是否检查了所有页眉？水印可能在`header3.xml`
- [ ] 关系文件是否加载？水印图片通过`header*.xml.rels`引用

### 6.5 脚注/题注检测

- [ ] 是否加载了`word/footnotes.xml`文件？
- [ ] 是否检查了图片段落中的`w:footnoteReference`？
- [ ] 脚注内容是否可能在脚注而非段落中？
- [ ] 题注可能不单独成段，而以脚注形式附加在图片段落中

### 6.6 页面边框/背景检测

- [ ] 页面边框在`sectPr/pgBorders`中，不在普通段落中
- [ ] 页面背景`w:background`在body直接子级，不在sectPr中
- [ ] XML序列化后命名空间膨胀：`w:pBdr`→`{http://...}pBdr`，需用DOM查找

---

## 七、示例代码

### 7.1 解压并加载XML

```python
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

def unzip_docx(docx_path):
    xmldir = Path(docx_path).parent / ('_xml_' + Path(docx_path).stem)
    if xmldir.exists():
        shutil.rmtree(xmldir)
    with zipfile.ZipFile(docx_path, 'r') as z:
        z.extractall(xmldir)
    return xmldir

def load_document(xml_dir):
    doc_xml = xml_dir / 'word' / 'document.xml'
    if doc_xml.exists():
        return ET.parse(str(doc_xml)).getroot()
    return None
```

### 7.2 检测标题格式

```python
def check_title_format(doc_root):
    paras = XA(doc_root, 'w:body/w:p')
    if not paras:
        return 0
    title = paras[0]
    rPr = X(title, 'w:pPr/w:rPr')
    score = 0
    if rPr is not None:
        # 字体
        rf = X(rPr, 'w:rFonts')
        if rf is not None and '行楷' in GA(rf, 'w:eastAsia', ''):
            score += 2
        # 加粗
        if X(rPr, 'w:b') is not None:
            score += 2
        # 字号
        sz = X(rPr, 'w:sz')
        if sz is not None and int(GA(sz, 'w:val', '0')) == 44:
            score += 2
    return score
```

---

## 四、新增模式（2026-07-02）

### 4.6 段落边框 vs 文字边框

**关键区别**：
- 段落边框：`w:pPr/w:pBdr`（整个段落的边框）
- 文字边框：`w:rPr/w:bdr`（选中文字的边框）

```python
# 检测段落边框（如"水膜实验中"段落的方框）
pPr = para.find('w:pPr')
if pPr is not None:
    pBdr = pPr.find('w:pBdr')
    if pBdr is not None:
        # 检查各边边框
        for side in ['top', 'left', 'bottom', 'right']:
            border = pBdr.find('w:' + side)
            if border is not None:
                val = GA(border, 'w:val', '')  # 线条样式
                sz = GA(border, 'w:sz', '')   # 线宽
                color = GA(border, 'w:color', '')  # 颜色
```

### 4.7 段落底纹

**位置**：`w:pPr/w:shd`

```python
# 检测段落底纹
pPr = para.find('w:pPr')
if pPr is not None:
    shd = pPr.find('w:shd')
    if shd is not None:
        fill = GA(shd, 'w:fill', '')  # 底纹颜色
        val = GA(shd, 'w:val', '')    # 底纹样式（clear/solid）
```

### 4.8 样式继承检测

**问题**：页眉页脚的对齐方式可能通过样式继承，XML中无显式`w:jc`。

**解决方法**：检查样式定义（`styles.xml`）

```python
# 1. 加载styles.xml
styles_tree = ET.parse('word/styles.xml')
styles_root = styles_tree.getroot()

# 2. 获取段落样式ID
pStyle = pPr.find('w:pStyle')
if pStyle is not None:
    style_id = GA(pStyle, 'w:val', '')
    
    # 3. 在styles.xml中查找样式定义
    for style in styles_root.findall('.//{' + WNS + '}style'):
        if style.get('{' + WNS + '}styleId', '') == style_id:
            # 4. 检查样式中的jc
            style_jc = style.find('.//{' + WNS + '}jc')
            if style_jc is not None:
                val = GA(style_jc, 'w:val', '')
```

**常见样式ID**：
- `a`: 普通样式
- `a3`: 页眉
- `a4`: 页脚

### 4.19 脚注与尾注检测（重要！）

**问题**：图片题注可能以脚注形式实现，不在document.xml的段落中。

**位置**：`word/footnotes.xml` 和 `word/endnotes.xml`

```python
# 1. 加载脚注文件
fn_path = xml_dir / 'word' / 'footnotes.xml'
if fn_path.exists():
    footnotes_root = ET.parse(str(fn_path)).getroot()

# 2. 查找脚注内容
for fn in footnotes_root.findall('{' + WNS + '}footnote'):
    fn_id = fn.get('{' + WNS + '}id', '')
    if fn_id in ('-1', '0'):  # 跳过分隔符
        continue
    
    # 提取文本
    txt = ''.join(t.text or '' for t in fn.findall('.//{' + WNS + '}t'))
    
    # 检查格式
    rPr = fn.find('.//{' + WNS + '}rPr')
    if rPr is not None:
        rf = rPr.find('{' + WNS + '}rFonts')
        sz = rPr.find('{' + WNS + '}sz')
        b = rPr.find('{' + WNS + '}b')
        # 注意: b w:val="0" 表示不加粗
```

**关键**：
- 脚注引用在document.xml中：`w:footnoteReference w:id="1"`
- 脚注内容在footnotes.xml中：`w:footnote w:id="1"`
- 判断是否为图片题注：检查脚注引用所在的段落是否包含图片

### 4.20 页面边框与背景

**页面边框**：`w:sectPr/w:pgBorders`
```python
for sp in doc_root.findall('{' + WNS + '}sectPr'):
    pgBorders = sp.find('{' + WNS + '}pgBorders')
    if pgBorders is not None:
        for side in ['top', 'left', 'bottom', 'right']:
            border = pgBorders.find('{' + WNS + '}' + side)
            val = GA(border, 'w:val', '')     # double/single
            sz = GA(border, 'w:sz', '0')      # 6 = 0.75pt
            color = GA(border, 'w:color', '') # 0070C0 = 蓝色
```

**页面背景**：`w:background`（在document.xml的body中，不在sectPr中）
```python
bg = doc_root.find('{' + WNS + '}background')
if bg is not None:
    bg_color = bg.get('{' + WNS + '}color', '')
    theme_color = bg.get('{' + WNS + '}themeColor', '')
```

### 4.21 图片所在段落识别

```python
# 找到包含图片的段落
for p in doc_root.findall('.//{' + WNS + '}p'):
    if p.find('.//{' + WPNS + '}anchor') is not None or \
       p.find('.//{' + WPNS + '}extent') is not None:
        img_para = p  # 图片所在段落
        # 检查同一段落中是否有脚注引用
        fn_ref = p.find('.//{' + WNS + '}footnoteReference')
        break
```

`w:sectPr` 可以在两个位置：

1. **段落级**：`w:p/w:pPr/w:sectPr`（如分栏设置）
2. **文档级**：`w:body/w:sectPr`（如页面设置）

```python
# 方法1：查找所有sectPr
sectPrs = root.findall('.//{' + WNS + '}sectPr')
for sp in sectPrs:
    cols = sp.find('w:cols')
    if cols is not None:
        num = GA(cols, 'w:num', '1')
        
# 方法2：区分段落级和文档级
# 段落级：在w:pPr中
for p in root.findall('.//{' + WNS + '}p'):
    pPr = p.find('w:pPr')
    if pPr is not None:
        sp = pPr.find('w:sectPr')
        if sp is not None:
            # 这是段落级sectPr
            
# 文档级：在w:body末尾
body = root.find('{' + WNS + '}body')
if body is not None:
    sp = body.find('w:sectPr')
    if sp is not None:
        # 这是文档级sectPr
```

### 4.10 竖排文字检测

**方法1**：OpenXML（`w:txbxContent` 中的 `w:textDirection`）

```python
txbxs = root.findall('.//{' + WNS + '}txbxContent')
for txbx in txbxs:
    for p in txbx.findall('.//{' + WNS + '}p'):
        pPr = p.find('w:pPr')
        if pPr is not None:
            td = pPr.find('w:textDirection')
            if td is not None:
                vert = GA(td, 'w:val', '')  # "tbRl" 表示竖排
```

**方法2**：VML（`w:pict` 中的 `layout-flow`）

```python
picts = root.findall('.//{' + WNS + '}pict')
for pict in picts:
    xs = ET.tostring(pict, encoding='unicode')
    if 'layout-flow:vertical' in xs:
        # 竖排
```

### 4.11 表格格式检测

```python
# 表格属性
tbl = root.find('.//{' + WNS + '}tbl')
tblPr = tbl.find('w:tblPr')

# 表格居中
jc = tblPr.find('w:jc')
if jc is not None and GA(jc, 'w:val', '') == 'center':
    # 居中

# 表格宽度
tblW = tblPr.find('w:tblW')
if tblW is not None and GA(tblW, 'w:type', '') == 'auto':
    # 根据内容自动调整

# 表格边框
tblBorders = tblPr.find('w:tblBorders')
if tblBorders is not None:
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = tblBorders.find('w:' + side)
        if border is not None:
            val = GA(border, 'w:val', '')   # 线条样式
            sz = GA(border, 'w:sz', '')     # 线条宽度
            color = GA(border, 'w:color', '')  # 颜色

# 单元格底纹
cells = tbl.findall('.//{' + WNS + '}tc')
for cell in cells:
    tcPr = cell.find('w:tcPr')
    if tcPr is not None:
        shd = tcPr.find('w:shd')
        if shd is not None:
            fill = GA(shd, 'w:fill', '')  # 底纹颜色
```

---

### 4.12 超链接域代码检测（重要！）

**问题**：Word中超链接可能不使用`w:hyperlink`元素，而是用域代码`w:instrText`实现。

```python
# 方法1: 检查w:hyperlink元素（标准方式）
hyperlinks = doc_root.findall('.//{' + WNS + '}hyperlink')
for hl in hyperlinks:
    rid = hl.get('{' + RNS + '}id', '')
    if rid in main_rels:
        target = main_rels[rid]  # URL

# 方法2: 检查域代码（字段代码方式，更常见！）
instr_texts = doc_root.findall('.//{' + WNS + '}instrText')
for it in instr_texts:
    if it.text and 'HYPERLINK' in it.text.upper():
        url = it.text.strip()  # 如: HYPERLINK "http://cwca.org.cn/"

# 方法3: 搜索原始XML
xs = ET.tostring(doc_root, encoding='unicode')
if 'HYPERLINK' in xs.upper() and 'cwca' in xs.lower():
    # 找到超链接
```

**教训**：必须同时检查两种超链接实现方式。

### 4.13 样式继承深度检测（关键！）

**问题**：字体、字号、加粗等格式可能不在run/pPr中，而是继承自段落样式或默认样式。

```python
# 1. 获取段落样式ID
pStyle = pPr.find('{' + WNS + '}pStyle')
style_id = GA(pStyle, 'w:val', '')  # 如 "a3"

# 2. 在styles.xml中查找样式
for style in styles_root.findall('.//{' + WNS + '}style'):
    if style.get('{' + WNS + '}styleId', '') == style_id:
        s_rPr = style.find('{' + WNS + '}rPr')
        if s_rPr is not None:
            # 检查样式中的字号
            sz = s_rPr.find('{' + WNS + '}sz')
            # 检查样式中的加粗
            b = s_rPr.find('{' + WNS + '}b')
            # 注意: b val="0" 表示不加粗
        # 3. 检查basedOn基础样式
        basedOn = style.find('{' + WNS + '}basedOn')
        if basedOn is not None:
            base_id = GA(basedOn, 'w:val', '')  # 如 "a"
            # 递归检查基础样式...

# 4. 检查默认段落样式（docDefaults）
# 在styles.xml中: w:docDefaults/w:rPrDefault/w:rPr
```

**常见样式链**：
- `a3` (Normal Web) → basedOn `a` (Normal) → docDefaults
- `a4` (header) → basedOn `a` (Normal) → docDefaults
- 默认样式`a`可能有 `w:b`（加粗）和 `w:rFonts w:eastAsia="楷体"`

### 4.14 中文字号对照表（修正版）

| 中文字号 | 磅值(pt) | w:sz值 | 说明 |
|---------|---------|--------|------|
| 初号 | 42 | 84 | 最大 |
| 小初 | 36 | 72 | |
| 一号 | 26 | 52 | 标题常用 |
| 小一 | 24 | 48 | |
| 二号 | 22 | 44 | |
| 小二 | 18 | 36 | |
| 三号 | 16 | 32 | |
| 小三 | 15 | 30 | |
| 四号 | 14 | 28 | 表格常用 |
| 小四 | 12 | 24 | 正文常用 |
| 五号 | 10.5 | 21 | 页眉常用 |
| 小五 | 9 | 18 | |
| 六号 | 7.5 | 15 | |
| 小六 | 6.5 | 13 | |
| 七号 | 5.5 | 11 | |
| 八号 | 5 | 10 | 最小 |

**注意**：五号 = sz 21（不是22！），小四 = sz 24，四号 = sz 28

### 4.15 正文段落定位（重要！）

**问题**：题目说"第五段"，但XML中段落编号包含标题，需要跳过标题计算。

```python
paras = doc_root.findall('.//{' + WNS + '}p')
# 段落0 = 标题
# 段落1 = 正文第1段
# 段落5 = 正文第5段（总段落6）
target_para = paras[5]  # 正文第五段
```

**备选策略**：搜索含特定关键词且有目标格式的段落
```python
for p in paras:
    if '关键词' in GT(p):
        pPr = p.find('{' + WNS + '}pPr')
        if pPr is not None:
            has_bdr = pPr.find('{' + WNS + '}pBdr') is not None
            has_shd = pPr.find('{' + WNS + '}shd') is not None
            if has_bdr or has_shd:
                target = p
                break
```

### 4.16 表格自动调整检测

```python
tblPr = tbl.find('{' + WNS + '}tblPr')
tblW = tblPr.find('{' + WNS + '}tblW')
tblLayout = tblPr.find('{' + WNS + '}tblLayout')

# "根据内容自动调整表格" = tblW type="auto" 且 无tblLayout或tblLayout type!="fixed"
auto_fit = False
if tblW is not None and GA(tblW, 'w:type', '') == 'auto':
    if tblLayout is None or GA(tblLayout, 'w:type', '') != 'fixed':
        auto_fit = True
```

### 4.17 表格排序验证

```python
# 提取某列的值，检查是否升序
rows = tbl.findall('{' + WNS + '}tr')
values = []
for row in rows[1:]:  # 跳过表头
    cells = row.findall('{' + WNS + '}tc')
    if len(cells) >= target_col:
        text = GT(cells[target_col])
        nums = re.findall(r'\d+', text)
        if nums:
            values.append(int(nums[0]))
if values == sorted(values):
    # 升序排列
```

### 4.18 段落底纹颜色对照

| 颜色名 | RGB值 | 说明 |
|--------|-------|------|
| 浅蓝色 | 00B0F0 | Word"浅蓝"底纹 |
| 浅蓝色 | CCE5FF | 另一种浅蓝 |
| 黄色 | FFFF00 | 边框常用 |
| 红色 | FF0000 | 字体颜色 |
| 水绿色 | 4BACC6 | accent5 |
| 橄榄色 | C2D69B | accent3 tint40% |
| 紫色 | CCC0D9 | accent4 tint60% |

### 4.22 GA函数实现注意事项（重要！）

**致命陷阱**：命名空间常量 `W = '{http://...}'` 已含大括号，格式化时不能再用 `{` 和 `}`。

```python
# ❌ 错误：W 本身是 '{http://...}'，'{%s}%s' % (W, 'tag') → '{{http://...}}tag'（双层括号！）
return elem.get('{%s}%s' % (W, 'tag'), default)

# ✅ 正确：ns_map 存裸URI（无大括号），'%s%s' 拼接
ns_uris = {'w': 'http://schemas.openxmlformats.org/...'}
return elem.get('%s%s' % (ns_uris[ns], 'tag'), default)
```

**症状**：`GA(elem, 'w:val', '0')` 永远返回默认值 `'0'`，所有属性检测静默失败，无报错。

### 4.23 文本跨 w:t 分段检测策略

**问题**：文本在XML中被切分到多个 `<w:t>` 元素，字符串搜索失效。

```
# XML实际结构：
<w:r><w:t>表格</w:t></w:r>
<w:r><w:t> 1 </w:t></w:r>
<w:r><w:t>2010年</w:t></w:r>

# 错误：'表格 1' NOT IN doc_xml（因为中间有XML标签）
# 正确：用 get_text() 合并所有 w:t 文本后再搜索
full_text = ''.join(t.text or '' for t in paragraph.iter(W + 't'))
if '表格 1' in full_text:  # ✓ 正确
```

### 4.24 表格内双级别 rPr 检测

表格单元格的字体有两种设置方式，必须都检查：

```python
# 方式1: 运行级 rPr（w:tc/w:p/w:r/w:rPr）
for r in cell.iter(W + 'r'):
    rPr = child(r, W + 'rPr')
    if rPr: check_font(rPr)

# 方式2: 段落级 rPr（w:tc/w:p/w:pPr/w:rPr）——表格常见！
for p in children(cell, W + 'p'):
    pPr = child(p, W + 'pPr')
    if pPr is not None:
        pRPr = child(pPr, W + 'rPr')
        if pRPr is not None: check_font(pRPr)
```

### 4.25 页面背景图检测（w:background + VML fill）

图片页面背景不在header水印中，而在 `w:background` 元素链：

```python
bg = child(doc_root, W + 'background')  # 直接子级，不在sectPr中
if bg is not None:
    # 检查 VML fill（图片填充）
    for fill in bg.iter(V + 'fill'):
        title = fill.get(O + 'title', '')  # 图片标题如"背景"
        rtype = fill.get('type', '')       # "frame" = 图片框架填充
        rid = fill.get(R + 'id', '')       # 关系ID
    # 检查 VML background 元素
    for vbg in bg.iter(V + 'background'):
        # 确认存在
        pass
    # 也可检查 VML imagedata
    for im in bg.iter(V + 'imagedata'):
        title = im.get(O + 'title', '')
```

---

*最后更新：2026-07-03（补充GA函数双括号Bug、VML形状类型表、文本跨段检测、表内双级别rPr、页面背景fill检测）*
