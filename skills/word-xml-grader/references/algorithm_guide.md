# Word XML 评分算法思路

> 基于 Word 2010 操作题评分实战经验总结，覆盖 11 种常见格式检测的完整算法指南。

## 一、核心流程

```
1. 解压 .docx (zipfile) → 临时目录
2. 加载关键 XML: document.xml, styles.xml, settings.xml, footnotes.xml, 
   header*.xml, footer*.xml, diagrams/*.xml, _rels/document.xml.rels
3. 逐项检测格式要求 (req1 ~ reqN)
4. 生成 HTML 评分报告
```

## 二、辅助函数体系（必须使用）

### 命名空间常量
```python
WNS   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
WPNS  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
ANS   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PICNS = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
RNS   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
VNS   = 'urn:schemas-microsoft-com:vml'
W14NS = 'http://schemas.microsoft.com/office/word/2010/wordml'
```

### 五大辅助函数
```python
def T(tag):
    """转换 'w:val' → '{http://...}val'"""
    if ':' not in tag: return tag
    p, l = tag.split(':', 1)
    ns_map = {'w': WNS, 'wp': WPNS, 'a': ANS, 'pic': PICNS, 'r': RNS, 'v': VNS, 'w14': W14NS}
    return '{' + ns_map[p] + '}' + l if p in ns_map else tag

def X(parent, path):
    """查找单个子元素，支持 / 逐级和 // 递归"""
    if parent is None: return None
    if '//' in path:
        pre, post = path.split('//', 1)
        ctx = parent if pre in ('.', '') else X(parent, pre)
        if ctx is None: return None
        return ctx.find('.//' + T(post))
    parts = path.split('/')
    node = parent
    for p in parts[:-1]:
        node = node.find(T(p)) if node is not None else None
    return node.find(T(parts[-1])) if node is not None else None

def XA(parent, path):
    """查找所有子元素，支持 / 逐级和 // 递归"""
    if parent is None: return []
    if '//' in path:
        pre, post = path.split('//', 1)
        ctx = parent if pre in ('.', '') else X(parent, pre)
        if ctx is None: return []
        return ctx.findall('.//' + T(post))
    parts = path.split('/')
    node = parent
    for p in parts[:-1]:
        node = node.find(T(p)) if node is not None else None
    return node.findall(T(parts[-1])) if node is not None else []

def GA(elem, attr, default=None):
    """安全获取带命名空间的属性
    ⚠️ ns_map 存裸URI（无大括号），用 '%s%s' 拼接！
    错误：'{%s}%s' % (W, 'attr') → '{{http://...}}attr'（双层括号，必失败）
    正确：'%s%s' % ('http://...', 'attr') → '{http://...}attr'
    """
    if elem is None: return default
    if ':' in attr:
        ns, local = attr.split(':', 1)
        ns_uris = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'v': 'urn:schemas-microsoft-com:vml',
            'o': 'urn:schemas-microsoft-com:office:office',
        }
        return elem.get('%s%s' % (ns_uris.get(ns, ''), local), default)
    return elem.get(attr, default)

def GT(elem):
    """获取元素内所有 w:t 文字"""
    if elem is None: return ""
    return ''.join(t.text or '' for t in elem.findall('.//{' + WNS + '}t'))
```

## 三、XML 结构速查

### .docx 内部文件
```
[docx]/
├── word/
│   ├── document.xml              # 正文主文档
│   ├── styles.xml                # 所有样式定义
│   ├── settings.xml              # 文档设置
│   ├── stylesWithEffects.xml     # 样式效果（Word 2010+）
│   ├── header1.xml / header2.xml / header3.xml   # 页眉（可多个）
│   ├── footer1.xml / footer2.xml / footer3.xml   # 页脚（可多个）
│   ├── _rels/document.xml.rels   # 主文档关系（超链接、图片、页眉页脚引用）
│   ├── theme/theme1.xml          # 主题颜色
│   └── media/                    # 嵌入图片
├── [Content_Types].xml
├── _rels/.rels
└── docProps/
```

## 四、中文字号对照表（关键！）

| 中文字号 | 磅值(pt) | w:sz值 | 用途 |
|---------|---------|--------|------|
| 初号 | 42 | 84 | |
| 小初 | 36 | 72 | |
| 一号 | 26 | 52 | 标题 |
| 小一 | 24 | 48 | |
| 二号 | 22 | 44 | |
| 小二 | 18 | 36 | |
| 三号 | 16 | 32 | |
| 小三 | 15 | 30 | |
| 四号 | 14 | 28 | 表格文字 |
| 小四 | 12 | 24 | 正文 |
| 五号 | 10.5 | **21** | 页眉页脚 |
| 小五 | 9 | 18 | |
| 六号 | 7.5 | 15 | |

> **易错**: 五号 = 21（不是22！），小四 = 24，四号 = 28，一号 = 52

## 五、常见格式检测逻辑

### 1. 页面设置
**XML位置**: `document.xml` 中的 `w:sectPr`

```python
for sp in doc_root.findall('.//{' + WNS + '}sectPr'):
    # 页面大小 A4: w=11906 h=16838, code=9
    pgSz = sp.find('{' + WNS + '}pgSz')
    w = GA(pgSz, 'w:w', '')  # 必须用 GA()！
    code = GA(pgSz, 'w:code', '')

    # 页边距 (单位: 缇, 1cm=567缇)
    pgMar = sp.find('{' + WNS + '}pgMar')
    top = GA(pgMar, 'w:top', '0')      # 3cm = 1701
    bottom = GA(pgMar, 'w:bottom', '0')
    left = GA(pgMar, 'w:left', '0')    # 3.5cm = 1985
    right = GA(pgMar, 'w:right', '0')
    gutter = GA(pgMar, 'w:gutter', '0')  # 0.6cm = 340

    # 文档网格
    docGrid = sp.find('{' + WNS + '}docGrid')
    linePitch = GA(docGrid, 'w:linePitch', '0')   # 每页42行 = 319
    charSpace = GA(docGrid, 'w:charSpace', '0')    # 每行40字符 = -4285
    break  # 只检查第一个sectPr
```

**单位换算**: 1cm = 567缇, 1磅 = 20半磅, 1cm = 914400/2.54 EMU

### 2. 替换文字
```python
xs = ET.tostring(doc_root, encoding='unicode')
if '错词' not in xs and '正确词' in xs:
    score = 满分  # 已替换
```

### 3. 标题格式（含文本效果）
**XML位置**: `document.xml` 段落1（paras[0]）的 `w:pPr/w:rPr`

```python
paras = doc_root.findall('.//{' + WNS + '}p')
title = paras[0]
pPr = title.find('{' + WNS + '}pPr')
rPr = pPr.find('{' + WNS + '}rPr') if pPr else None

# 字体: w:rFonts w:eastAsia
rf = rPr.find('{' + WNS + '}rFonts')
font = GA(rf, 'w:eastAsia', '')  # 如 "黑体"

# 加粗: w:b
rPr.find('{' + WNS + '}b') is not None

# 字号: w:sz (一号=52)
sz = rPr.find('{' + WNS + '}sz')
size = int(GA(sz, 'w:val', '0'))

# 居中: w:jc w:val="center"
jc = pPr.find('{' + WNS + '}jc')

# 文本效果（艺术字）: w14 命名空间
textFill = rPr.find('.//{' + W14NS + '}textFill')      # 填充
textOutline = rPr.find('.//{' + W14NS + '}textOutline')  # 轮廓
shadow = rPr.find('.//{' + W14NS + '}shadow')            # 阴影
```

### 4. 正文段落格式
**XML位置**: 段落2（paras[1]）的 `w:pPr`

```python
body = paras[1]
pPr = body.find('{' + WNS + '}pPr')

# 字体字号: 可能在 pPr/rPr 或段落样式中
rPr = pPr.find('{' + WNS + '}rPr')
# 如果 rPr 中没有 sz，检查样式链
pStyle = pPr.find('{' + WNS + '}pStyle')
style_id = GA(pStyle, 'w:val', '')  # 如 "a3"
# → 在 styles.xml 中查找该样式的 rPr/sz

# 缩进 (单位: 1/20字符)
ind = pPr.find('{' + WNS + '}ind')
firstLineChars = GA(ind, 'w:firstLineChars', '0')  # 首行缩进2字符=200
leftChars = GA(ind, 'w:leftChars', '0')            # 左缩进0.5字符=50
rightChars = GA(ind, 'w:rightChars', '0')

# 间距 (单位: 1/100行)
spacing = pPr.find('{' + WNS + '}spacing')
beforeLines = GA(spacing, 'w:beforeLines', '0')  # 段前1行=100
afterLines = GA(spacing, 'w:afterLines', '0')    # 段后0.5行=50
```

### 5. 图片设置
**XML位置**: `w:drawing` 中的 `wp:anchor`（浮动）或 `wp:inline`（嵌入式）

```python
for dwg in doc_root.findall('.//{' + WNS + '}drawing'):
    blip = dwg.find('.//{' + ANS + '}blip')
    if blip is None: continue

    # 尺寸 (EMU → cm: 除以914400*2.54)
    ext = dwg.find('.//{' + WPNS + '}extent')
    cx = int(GA(ext, 'cx', '0')) / 914400 * 2.54
    cy = int(GA(ext, 'cy', '0')) / 914400 * 2.54

    # 环绕方式
    dwg.find('.//{' + WPNS + '}wrapSquare') is not None    # 四周型
    dwg.find('.//{' + WPNS + '}wrapTight') is not None     # 紧密型
    dwg.find('.//{' + WPNS + '}wrapTopAndBottom') is not None  # 上下型

    # 图片样式（柔化边缘椭圆）
    pic = dwg.find('.//{' + PICNS + '}pic')
    spPr = pic.find('.//{' + ANS + '}spPr') if pic else None
    # 椭圆形状: <a:prstGeom prst="ellipse">
    # 柔化边缘: <a:effectLst><a:softEdge rad="..."/></a:effectLst>
```

### 6. 页眉页脚
**XML位置**: `header*.xml`, `footer*.xml`

```python
# 遍历所有页眉
for nm, root in headers.items():
    txt = GT(root)
    if '保护野生动物' in txt:
        # 检查字体、字号、加粗（可能在 run/rPr 或样式继承中）
        rPr = root.find('.//{' + WNS + '}rPr')
        # 微软雅黑: w:rFonts w:eastAsia
        # 五号: w:sz w:val="21" (不是22！)
        # 加粗: w:b

# 遍历所有页脚
for nm, root in footers.items():
    txt = GT(root)
    # 罗马数字页码: 检查 "Ⅰ" 或 "upperRoman" 或 PAGE 域
    xs = ET.tostring(root, encoding='unicode')
    if 'Ⅰ' in xs or 'II' in xs or 'upperRoman' in xs:
        # 罗马数字格式
```

> **关键**: 五号字体 sz=21（不是22）。加粗可能继承自默认样式 "a"(Normal)，需检查 `basedOn` 链。

### 7. 段落边框和底纹（重点！）
**XML位置**: 目标段落的 `w:pPr`

```python
# 段落定位（重要！）
# "第五段" = 正文第五段 = 总段落6（跳过标题）
# 方法1: 直接索引 paras[5]
# 方法2: 搜索含关键词且有格式的段落
target = None
for p in paras:
    if '关键词' in GT(p):
        pPr = p.find('{' + WNS + '}pPr')
        if pPr is not None:
            if pPr.find('{' + WNS + '}pBdr') or pPr.find('{' + WNS + '}shd'):
                target = p
                break

pPr = target.find('{' + WNS + '}pPr')

# 字体颜色: 检查 pPr/rPr 和 run/rPr
rPr = pPr.find('{' + WNS + '}rPr')
color = rPr.find('{' + WNS + '}color') if rPr else None
c = GA(color, 'w:val', '')  # FF0000 = 红色
# 也检查 run 级别
for run in target.findall('.//{' + WNS + '}r'):
    rrPr = run.find('{' + WNS + '}rPr')
    rc = rrPr.find('{' + WNS + '}color') if rrPr else None
    if rc and 'FF0000' in GA(rc, 'w:val', '').upper():
        # run 级别红色

# 段落边框
pBdr = pPr.find('{' + WNS + '}pBdr')
for side in ['top', 'left', 'bottom', 'right']:
    border = pBdr.find('{' + WNS + '}' + side)
    val = GA(border, 'w:val', '')     # single, double
    sz = GA(border, 'w:sz', '0')      # 12 = 1.5磅
    color = GA(border, 'w:color', '') # FFFF00 = 黄色
    shadow = GA(border, 'w:shadow', '0')  # 1 = 有阴影

# 段落底纹
shd = pPr.find('{' + WNS + '}shd')
fill = GA(shd, 'w:fill', '')  # 00B0F0 = 浅蓝色
```

> **易错**: 不是 `w:rPr/w:bdr`（文字边框），而是 `w:pPr/w:pBdr`（段落边框）
> **颜色**: 浅蓝底纹=00B0F0（不是CCE5FF！）

### 8. 表格边框
**XML位置**: `w:tbl/w:tblPr/w:tblBorders`

```python
tbl = doc_root.findall('.//{' + WNS + '}tbl')[0]
tblPr = tbl.find('{' + WNS + '}tblPr')
tblBorders = tblPr.find('{' + WNS + '}tblBorders')

# 外边框
for side in ['top', 'left', 'bottom', 'right']:
    border = tblBorders.find('{' + WNS + '}' + side)
    val = GA(border, 'w:val', '')    # double = 双实线
    sz = GA(border, 'w:sz', '0')     # 12 = 1.5磅
    color = GA(border, 'w:color', '')  # CCC0D9 = 紫色淡色60%

# 内边框
for side in ['insideH', 'insideV']:
    border = tblBorders.find('{' + WNS + '}' + side)
    val = GA(border, 'w:val', '')    # single = 单实线
    sz = GA(border, 'w:sz', '0')     # 24 = 3磅
    color = GA(border, 'w:color', '')  # C2D69B = 橄榄色淡色40%
```

### 9. 表格内容格式（重点！）
**XML位置**: `w:tbl` 中的单元格 `w:tc`

> ⚠️ **双级别 rPr**：表格字体经常在段落级 `w:p/w:pPr/w:rPr` 而非运行级 `w:r/w:rPr`，必须两者都检查！
```python
# 表格居中
jc = tblPr.find('{' + WNS + '}jc')
GA(jc, 'w:val', '') == 'center'

# 自动调整（关键判断！）
tblW = tblPr.find('{' + WNS + '}tblW')
tblLayout = tblPr.find('{' + WNS + '}tblLayout')
auto_fit = (GA(tblW, 'w:type', '') == 'auto' and
            (tblLayout is None or GA(tblLayout, 'w:type', '') != 'fixed'))

# 字体（楷体）- 可能继承自默认样式！
# 1. 先获取默认段落样式的字体
default_font = ''
for style in styles_root.findall('.//{' + WNS + '}style'):
    if (style.get('{' + WNS + '}type') == 'paragraph' and
        style.get('{' + WNS + '}default') == '1'):
        rPr = style.find('{' + WNS + '}rPr')
        rf = rPr.find('{' + WNS + '}rFonts') if rPr else None
        default_font = GA(rf, 'w:eastAsia', '')  # "楷体"

# 2. 检查单元格 run 的字体 AND 段落级 rPr
for cell in tbl.findall('.//{' + WNS + '}tc'):
    # 运行级
    for run in cell.findall('.//{' + WNS + '}r'):
        rPr = run.find('{' + WNS + '}rPr')
        run_font = GA(rPr.find('{' + WNS + '}rFonts'), 'w:eastAsia', '') if rPr else ''
        if run_font == '黑体':
            font_ok = True
    # 段落级（表格常见！）
    for p in cell.findall('{' + WNS + '}p'):
        pPr = p.find('{' + WNS + '}pPr')
        if pPr is not None:
            pRPr = pPr.find('{' + WNS + '}rPr')
            if pRPr is not None:
                pfont = GA(pRPr.find('{' + WNS + '}rFonts'), 'w:eastAsia', '')
                sz = GA(pRPr.find('{' + WNS + '}sz'), 'w:val', '')
                if pfont == '黑体' and sz == '24':
                    font_ok = True

# 字号: 四号 = sz 28
# 颜色: 水绿色 = 4BACC6

# 水平居中: 段落的 w:jc w:val="center"
# 垂直居中: 单元格的 w:tcPr/w:vAlign w:val="center"

# 排序验证
import re
rows = tbl.findall('{' + WNS + '}tr')
years = []
for row in rows[1:]:  # 跳过表头
    cells = row.findall('{' + WNS + '}tc')
    if len(cells) >= 3:
        text = GT(cells[2])  # 第三列
        nums = re.findall(r'\d+', text)
        if nums:
            years.append(int(nums[0]))
if years == sorted(years):
    # 升序排列 ✓
```

> **关键**: 楷体通常不在run的rPr中，而是继承自默认样式"a"(Normal)的 `w:rFonts w:eastAsia="楷体"`

### 10. 超链接（极易错！）
**超链接有两种实现方式**:

```python
# 方法1: w:hyperlink 元素（较少见）
hyperlinks = doc_root.findall('.//{' + WNS + '}hyperlink')
for hl in hyperlinks:
    rid = hl.get('{' + RNS + '}id', '')
    if rid in main_rels:
        target = main_rels[rid]  # URL

# 方法2: 域代码 w:instrText（更常见！）
instr_texts = doc_root.findall('.//{' + WNS + '}instrText')
for it in instr_texts:
    if it.text and 'HYPERLINK' in it.text.upper():
        # 找到超链接: HYPERLINK "http://cwca.org.cn/"
        url = it.text.strip()

# 方法3: 原始XML搜索（兜底）
xs = ET.tostring(doc_root, encoding='unicode')
if 'HYPERLINK' in xs.upper() and 'cwca' in xs.lower():
    # 找到超链接
```

> **必须同时检查两种方式**，标准答案常用域代码而非w:hyperlink元素。

### 11. 页面边框
**XML位置**: `w:sectPr/w:pgBorders`

```python
for sp in doc_root.findall('.//{' + WNS + '}sectPr'):
    pgBorders = sp.find('{' + WNS + '}pgBorders')
    if pgBorders is not None:
        for side in ['top', 'left', 'bottom', 'right']:
            border = pgBorders.find('{' + WNS + '}' + side)
            val = GA(border, 'w:val', '')     # single
            sz = GA(border, 'w:sz', '0')      # 24 = 3磅
            color = GA(border, 'w:color', '') # 0070C0 = 蓝色
            shadow = GA(border, 'w:shadow', '0')  # 1 = 阴影
    break
```

### 12. 艺术字 / 文本效果
**XML位置**: `w:pPr/w:rPr` 中的 w14 命名空间元素

```python
rPr = pPr.find('{' + WNS + '}rPr')
# 文本填充: w14:textFill
textFill = rPr.find('.//{' + W14NS + '}textFill')
# 文本轮廓: w14:textOutline
textOutline = rPr.find('.//{' + W14NS + '}textOutline')
# 文本阴影: w14:shadow
shadow = rPr.find('.//{' + W14NS + '}shadow')
```

## 六、样式继承深度检测（核心难点）

### 问题
字体、字号、加粗等格式可能不在 run/pPr 中，而是继承自段落样式或默认样式。

### 样式继承链
```
run/w:rPr → pPr/w:rPr → pStyle引用的样式 → basedOn引用的基础样式 → docDefaults
```

### 检测算法
```python
def get_style_property(styles_root, style_id, prop_path):
    """从样式链中查找属性"""
    visited = set()
    current_id = style_id
    while current_id and current_id not in visited:
        visited.add(current_id)
        for style in styles_root.findall('.//{' + WNS + '}style'):
            if style.get('{' + WNS + '}styleId', '') == current_id:
                # 在当前样式中查找
                elem = X(style, prop_path)
                if elem is not None:
                    return elem
                # 检查 basedOn
                basedOn = style.find('{' + WNS + '}basedOn')
                current_id = GA(basedOn, 'w:val', '') if basedOn else None
                break
        else:
            break
    return None

# 使用示例: 查找样式 a3 的字号
sz = get_style_property(styles_root, 'a3', 'w:rPr/w:sz')
if sz is not None:
    size = int(GA(sz, 'w:val', '0'))
```

### 常见样式ID
| 样式ID | 名称 | 说明 |
|--------|------|------|
| `a` | Normal | 默认段落样式，可能有 `w:b`（加粗）和 `w:rFonts w:eastAsia="楷体"` |
| `a3` | Normal (Web) | 正文段落常用，basedOn `a` |
| `a4` | header | 页眉样式，basedOn `a` |
| `a7` | Table Grid | 表格样式 |

> **注意**: 默认样式 "a" 中 `<w:b/>` 表示默认加粗，但如果页眉 run 中 `<w:b w:val="0"/>` 则取消加粗。

## 七、常见坑点

### 0. GA 函数双括号陷阱（致命！）
```python
# ❌ 错误：W = '{http://...}' 已含{}，'{%s}%s' % (W, 'attr') → '{{http://...}}attr'
# ✅ 正确：ns_map 存裸URI，'%s%s' % ('http://...', 'attr') → '{http://...}attr'
```
**症状：所有 GA() 返回默认值，全部检测静默失败，无报错。**

### 1. 属性获取必须用 GA()
```python
# 错误 ❌
val = elem.get('w:val')  # 返回 None！
val = elem.get('val')    # 可能返回 None！

# 正确 ✓
val = GA(elem, 'w:val', '')  # 自动处理命名空间
```

### 2. 段落定位
- "第五段" = 正文第五段 = 总段落6（标题是段落1）
- 不要用 `paras[4]`，要用 `paras[5]`
- 备选: 搜索含关键词且有目标格式的段落

### 3. 中文字号易错
- 五号 = 21（不是22！）
- 小四 = 24，四号 = 28，一号 = 52

### 4. 超链接两种实现
- `w:hyperlink` 元素（较少见）
- `w:instrText` 域代码（更常见！）

### 5. 样式继承
- 字体可能在默认样式 "a" 中: `w:rFonts w:eastAsia="楷体"`
- 加粗可能在默认样式 "a" 中: `<w:b/>`
- 字号可能在段落样式 "a3" 中: `<w:sz w:val="24"/>`

### 6. 底纹颜色
- 浅蓝色 = 00B0F0（不是CCE5FF！）

### 7. 表格自动调整
- `tblW type="auto"` + 无 `tblLayout` 或 `tblLayout type!="fixed"`

### 8. VML vs DML
- VML: 旧版Word，元素在 `w:pict` 中
- DML: 新版Word，元素在 `w:drawing` 中
- 需同时检查两者

### 9. sectPr 位置
- 段落级: `w:p/w:pPr/w:sectPr`（分栏等）
- 文档级: `w:body/w:sectPr`（页面设置等）

### 10. 段落边框 vs 文字边框
- 段落边框: `w:pPr/w:pBdr`
- 文字边框: `w:rPr/w:bdr`

## 八、颜色值速查表

| 颜色名 | RGB值 | 用途 |
|--------|-------|------|
| 红色 | FF0000 | 字体颜色 |
| 黄色 | FFFF00 | 段落边框 |
| 浅蓝色 | 00B0F0 | 段落底纹 |
| 浅蓝色 | CCE5FF | 备选浅蓝 |
| 水绿色 | 4BACC6 | 表格文字（accent5） |
| 橄榄色 | C2D69B | 表格内边框（accent3 tint40%） |
| 紫色 | CCC0D9 | 表格外边框（accent4 tint60%） |
| 蓝色 | 0070C0 | 页面边框 |
| 灰色 | 333333 | 正文文字 |

## 九、调试技巧

### 1. 打印段落列表
```python
paras = doc_root.findall('.//{' + WNS + '}p')
for i, p in enumerate(paras):
    txt = GT(p)
    print(f'段落{i+1}: {txt[:60]}')
```

### 2. 打印元素完整XML
```python
xml_str = ET.tostring(elem, encoding='unicode')
print(xml_str[:2000])
```

### 3. 检查所有段落的格式
```python
for i, p in enumerate(paras):
    pPr = p.find('{' + WNS + '}pPr')
    info = []
    if pPr:
        if pPr.find('{' + WNS + '}pBdr'): info.append('有边框')
        if pPr.find('{' + WNS + '}shd'): info.append('有底纹')
        rPr = pPr.find('{' + WNS + '}rPr')
        if rPr:
            c = rPr.find('{' + WNS + '}color')
            if c: info.append(f'颜色={GA(c, "w:val", "")}')
    print(f'段落{i+1}: {info}')
```

### 4. 在原始XML中搜索
```python
xs = ET.tostring(doc_root, encoding='unicode')
if 'HYPERLINK' in xs:
    # 找到超链接域代码
```

### 5. 检查脚注内容
```python
# 加载脚注文件
fn_tree = ET.parse('word/footnotes.xml')
fn_root = fn_tree.getroot()
for fn in fn_root.findall('.//{' + WNS + '}footnote'):
    fn_id = fn.get('{' + WNS + '}id', '')
    if fn_id in ('-1', '0'): continue  # 跳过分隔符
    txt = ''.join(t.text or '' for t in fn.findall('.//{' + WNS + '}t'))
    print(f'脚注{fn_id}: [{txt[:60]}]')
```

### 6. 搜索所有XML文件中的关键词
```python
# 当document.xml找不到时，搜索所有XML文件
with zipfile.ZipFile(docx_path, 'r') as z:
    for name in z.namelist():
        if name.endswith('.xml'):
            content = z.read(name).decode('utf-8', errors='ignore')
            if '关键词' in content:
                print(f'找到关键词在: {name}')
```

## 十、算法生成模板

```python
def reqX(self):
    """检测XX格式"""
    log = []
    score = 0

    # 1. 找到目标元素
    paras = self.doc_root.findall('.//{' + WNS + '}p')
    if not paras:
        log.append('  NG 未找到段落')
        return {'req': X, 'name': 'XX', 'score': 0, 'full': Y, 'detail': log, 'status': '未通过'}

    target = paras[N]  # 或通过关键词搜索

    # 2. 检查格式属性
    pPr = target.find('{' + WNS + '}pPr')
    if pPr is not None:
        # 检查属性1
        elem = pPr.find('{' + WNS + '}elemName')
        if elem is not None:
            val = GA(elem, 'w:val', '')
            if val == '期望值':
                score += N
                log.append('  OK 描述 (+N)')

        # 检查样式继承
        pStyle = pPr.find('{' + WNS + '}pStyle')
        if pStyle is not None:
            style_id = GA(pStyle, 'w:val', '')
            # → 在 styles.xml 中查找

    # 3. 返回结果
    st = '通过' if score >= 8 else ('部分通过' if score >= 4 else '未通过')
    return {'req': X, 'name': 'XX', 'score': min(score, Y), 'full': Y,
            'detail': log, 'status': st}
```

## 十一、评分器架构

```python
class XMLGrader:
    def __init__(self, docx_path):
        self.docx = Path(docx_path)
        # ...

    def unzip(self):
        """解压 .docx 到临时目录"""

    def load_all(self):
        """加载所有 XML 文件"""
        # document.xml, styles.xml, settings.xml
        # header*.xml, footer*.xml
        # document.xml.rels

    def grade(self):
        """逐项评分"""
        for i in range(1, N+1):
            result = getattr(self, 'req%d' % i)()
            self.results.append(result)

    def req1(self): ...  # 每个要求一个方法
    def req2(self): ...
    # ...

    def report(self):
        """生成 HTML 报告"""
```

## 十二、标准答案验证策略

1. **先分析标准答案**: 运行前先用 Python 脚本检查标准答案的XML结构
2. **确认答案是标准的**: 用户确认后，评分器必须匹配标准答案的实际XML实现
3. **区分度验证**: 标准答案应得满分，原始素材应得接近0分
4. **常见问题排查顺序**:
   - 段落定位是否正确？（第N段 vs paras[N]）
   - 属性是否用 GA() 获取？
   - 格式是否在样式继承链中？
   - 超链接是否用域代码？
   - 题注是否在脚注中？（检查`footnotes.xml`）
   - 搜索范围是否覆盖所有XML文件？（header, footer, diagrams, footnotes）

---

*最后更新：2026-07-03（GA函数双括号Bug修复、表格双级别rPr、VML形状类型、精确匹配铁律）*
