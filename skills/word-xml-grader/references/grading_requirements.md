# Word 操作题评分要求详解

本文档详细说明11项评分要求的XML实现方式。

---

## 要求1：标题格式（10分）

### 检测点
- 字体：华文行楷
- 字形：加粗、倾斜
- 字号：二号（22pt = 44半磅）
- 颜色：深红色（C00000 或 800000）

### XML位置
`w:body/w:p[0]/w:pPr/w:rPr`

### 代码示例
```python
rPr = X(title, 'w:pPr/w:rPr')
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
    # 颜色
    co = X(rPr, 'w:color')
    if co is not None and GA(co, 'w:val', '').upper() in ['C00000', '800000']:
        score += 2
```

---

## 要求2：标题边框间距（10分）

### 检测点
- 字符边框：单实线、1.5磅、蓝色
- 段前间距：0.5行（50 = 50/100行）
- 段后间距：1行（100 = 1行）
- 对齐：居中

### XML位置
- 边框：`w:pPr/w:rPr/w:bdr`
- 间距：`w:pPr/w:spacing`
- 对齐：`w:pPr/w:jc`

### 代码示例
```python
# 字符边框
bdr = X(rPr, 'w:bdr')
if bdr is not None:
    val = GA(bdr, 'w:val', '')      # 样式
    sz = GA(bdr, 'w:sz', '0')       # 线宽（1/8磅）
    col = GA(bdr, 'w:color', '')     # 颜色
    if val in ['single', 'thick']:
        score += 2
    if abs(int(sz) - 12) < 4:      # 12 = 1.5磅
        score += 2
    if '00' in col.upper() and 'C0' in col.upper():
        score += 2

# 间距
sp = X(pPr, 'w:spacing')
if sp is not None:
    bl = GA(sp, 'w:beforeLines', '')  # 段前
    al = GA(sp, 'w:afterLines', '')   # 段后
    if bl == '50':
        score += 2
    if al == '100':
        score += 2

# 居中
jc = X(pPr, 'w:jc')
if jc is not None and GA(jc, 'w:val', '') == 'center':
    score += 2
```

---

## 要求3：正文格式（10分）

### 检测点
- 字体：华文中宋
- 字号：四号（14pt = 28半磅）
- 首行缩进：2字符（200 = 200/100字符）
- 行距：固定20磅（400 = 400/2 半磅，lineRule=exact）

### XML位置
`w:body/w:p[1:]/w:pPr`

### 注意
正文格式可能在：
1. `w:pPr/w:rPr`（段落默认格式）
2. `w:r/w:rPr`（文本级格式）

需检查两者。

### 代码示例
```python
for para in body_paras[:8]:
    pPr = X(para, 'w:pPr')
    if pPr is None: continue
    
    # 检查段落默认格式
    rp_p = X(pPr, 'w:rPr')
    if rp_p is not None:
        rf = X(rp_p, 'w:rFonts')
        if rf is not None and '中宋' in GA(rf, 'w:eastAsia', ''):
            f_font += 1
        
        sz = X(rp_p, 'w:sz')
        if sz is not None and int(GA(sz, 'w:val', '0')) == 28:
            f_sz += 1
    
    # 首行缩进
    ind = X(pPr, 'w:ind')
    if ind is not None:
        flc = GA(ind, 'w:firstLineChars', '0')
        if int(flc) >= 150:
            f_ind += 1
    
    # 行距
    sp = X(pPr, 'w:spacing')
    if sp is not None:
        ln = GA(sp, 'w:line', '0')
        rule = GA(sp, 'w:lineRule', '')
        if int(ln) == 400 and rule == 'exact':
            f_ln += 1
```

---

## 要求4：第3段图片设置（10分）

### 检测点
- 尺寸：14.5cm × 5cm
- 环绕方式：紧密型（wrapTight）
- 旋转：180度（rot=10800000，因为60000/度）
- 对齐：相对于栏居中

### XML位置
`w:body//w:drawing`

### 单位转换
```python
# EMU转厘米
cx = int(GA(ext, 'cx', '0'))
wcm = cx / 914400 * 2.54

# 旋转角度
rot = int(GA(xfrm, 'rot', '0'))
deg = rot / 60000.0  # 60000 = 1度
```

### 代码示例
```python
drawings = XA(doc_root, 'w:body//w:drawing')
for dwg in drawings:
    # 尺寸
    ext = X(dwg, '//wp:extent')
    if ext is not None:
        cx = int(GA(ext, 'cx', '0'))
        wcm = cx / 914400 * 2.54
        if abs(wcm - 14.5) < 1:
            score += 3
    
    # 环绕
    if X(dwg, '//wp:wrapTight') is not None:
        score += 2
    
    # 旋转
    xfrm = X(dwg, '//a:xfrm')
    if xfrm is not None:
        rot = int(GA(xfrm, 'rot', '0'))
        deg = rot / 60000.0
        if abs(deg - 180) < 10:
            score += 2
    
    # 对齐
    pH = X(dwg, '//wp:positionH')
    if pH is not None:
        rf = pH.get('relativeFrom', '')
        al = X(pH, 'wp:align')
        if rf == 'column' and al is not None and al.text == 'center':
            score += 1
```

---

## 要求5：脚注设置（5分）

### 检测点
- 编号格式：圆圈数字（numFmt = circle 或 encloseCircle）
- 脚注引用：文档中有`w:footnoteReference`
- 脚注内容：包含URL（`http`或`https`）

### XML位置
- 设置：`w:sectPr/w:footnotePr/w:numFmt`
- 引用：`w:body//w:footnoteReference`
- 内容：`footnotes.xml`

### 代码示例
```python
# 检查编号格式
for sp in XA(doc_root, 'w:body//w:sectPr'):
    fp = X(sp, 'w:footnotePr')
    if fp is not None:
        nf = X(fp, 'w:numFmt')
        if nf is not None:
            v = GA(nf, 'w:val', '')
            if 'Circle' in v:
                score += 2

# 检查引用
refs = XA(doc_root, 'w:body//w:footnoteReference')
if refs:
    score += 1

# 检查内容
if footnotes_root is not None:
    xs = ET.tostring(footnotes_root, encoding='unicode')
    if 'http' in xs.lower():
        score += 2
```

---

## 要求6：图片水印（5分）⭐关键

### 检测点
- 水印图片：存在于页眉中
- 冲蚀效果：blacklevel 或 gain 属性

### XML位置
`header*.xml`中的VML形状

### 关键洞察
Word的水印是**VML形状**（不是标准Open XML），需要：
1. 使用VML命名空间：`urn:schemas-microsoft-com:vml`
2. 查找`v:shape`元素（id包含"Watermark"）
3. 检查`v:imagedata`子元素的`blacklevel`/`gain`属性

### 代码示例
```python
VNS = 'urn:schemas-microsoft-com:vml'

for nm, root in headers.items():
    for shape in root.findall('.//{' + VNS + '}shape'):
        sid = shape.get('id', '')
        if 'Watermark' in sid:
            score += 3
            for imgdata in shape.findall('{' + VNS + '}imagedata'):
                blacklevel = imgdata.get('blacklevel', '')
                gain = imgdata.get('gain', '')
                if blacklevel or gain:
                    score += 2
                    break
            break
```

---

## 要求7：第2段分栏（10分）

### 检测点
- 左右缩进：各2字符（leftChars=rightChars=200）
- 分栏数：2栏（num=2）
- 分隔线：有（sep=1）
- 栏宽：不等宽（equalWidth=0），第1栏宽14字符（w=2940，因为1/20字符）

### XML位置
- 缩进：`w:body/w:p[1]/w:pPr/w:ind`
- 分栏：`w:body//w:sectPr/w:cols`

### 注意
分栏设置（`w:cols`）可能在：
1. 文档末尾的`w:sectPr`中
2. 某个段落的`w:pPr/w:sectPr`中（段落标记的分节符）

需检查所有`w:sectPr`。

### 代码示例
```python
# 检查缩进
paras = XA(doc_root, 'w:body/w:p')
for idx in range(1, min(6, len(paras))):
    para = paras[idx]
    ind = X(para, 'w:pPr/w:ind')
    if ind is not None:
        lc = GA(ind, 'w:leftChars', '0')
        rc = GA(ind, 'w:rightChars', '0')
        if int(lc) >= 150 and int(rc) >= 150:
            score += 4
            break

# 检查分栏
for sp in XA(doc_root, 'w:body//w:sectPr'):
    cols = X(sp, 'w:cols')
    if cols is not None:
        num = GA(cols, 'w:num', '1')
        if num == '2':
            score += 2
            sep = GA(cols, 'w:sep', '0')
            if sep == '1':
                score += 1
            ew = GA(cols, 'w:equalWidth', '1')
            if ew == '0':
                score += 1
            break
```

---

## 要求8：页眉页脚（15分）

### 检测点
- 奇偶页不同：启用（settings.xml中有`w:evenAndOddHeaders`）
- 奇数页眉："冬奥会"（或类似文字）
- 偶数页眉："体育精神"（或类似文字）
- 页脚：包含PAGE域（页码）、页码加粗

### XML位置
- 设置：`settings.xml/w:settings/w:evenAndOddHeaders`
- 页眉：`header*.xml`
- 页脚：`footer*.xml`

### 代码示例
```python
# 检查奇偶页不同
if settings_root is not None:
    if X(settings_root, 'w:evenAndOddHeaders') is not None:
        score += 3

# 检查页眉
for nm, root in headers.items():
    txt = GT(root)  # 提取所有文本
    if '冬奥' in txt:
        score += 3
    if '体育' in txt:
        score += 3

# 检查页脚
for nm, root in footters.items():
    xs = ET.tostring(root, encoding='unicode')
    if 'PAGE' in xs:
        score += 3
    if '<w:b' in xs:
        score += 2
```

---

## 要求9：图片题注（10分）

### 检测点
- 冰墩墩图片：存在于文档中
- 题注标签："图片"（SEQ域）
- 题注居中：题注段落有`w:jc w:val="center"`

### XML位置
- 图片：`w:body//w:drawing`
- 题注：`w:body//w:p`中包含`SEQ`和"图片"

### 注意
题注可能在**多个段落**中，需要检查所有含"SEQ"和"图片"的段落，而不仅仅第一个。

### 代码示例
```python
# 查找冰墩墩图片
drawings = XA(doc_root, 'w:body//w:drawing')
for dwg in drawings:
    cNv = X(dwg, '//pic:cNvPr')
    if cNv is not None:
        nm = GA(cNv, 'name', '')
        if '冰墩' in nm or '雪容' in nm:
            score += 3

# 查找题注
found_caption = False
for para in XA(doc_root, 'w:body//w:p'):
    xs = ET.tostring(para, encoding='unicode')
    if 'SEQ' in xs and '图片' in xs:
        score += 2
        found_caption = True
        # 检查居中
        jc = X(para, 'w:pPr/w:jc')
        if jc is not None and GA(jc, 'w:val', '') == 'center':
            score += 2
            break  # 找到居中的题注就停
```

---

## 要求10：页面边框背景（10分）

### 检测点
- 页面边框：存在（`w:pgBorders`）
- 阴影：有（w:shadow="1"）
- 渐变背景：存在（`w:background`）

### XML位置
- 边框：`w:body//w:sectPr/w:pgBorders`
- 背景：`w:background`

### 代码示例
```python
# 检查页面边框
for sp in XA(doc_root, 'w:body//w:sectPr'):
    pgb = X(sp, 'w:pgBorders')
    if pgb is not None:
        score += 3
        has_shadow = False
        for side in ['top', 'left', 'bottom', 'right']:
            b = X(pgb, 'w:' + side)
            if b is not None:
                sh = GA(b, 'w:shadow', '0')
                if sh == '1':
                    has_shadow = True
        if has_shadow:
            score += 3
        break

# 检查背景
bg = X(doc_root, 'w:background')
if bg is not None:
    score += 2
    xs = ET.tostring(bg, encoding='unicode')
    if 'grad' in xs.lower():
        score += 2
```

---

## 要求11：超链接（5分）

### 检测点
- 超链接文字："丝绸之路"（或类似文字）
- 目标URL：百度百科（baike.baidu.com）

### XML位置
`w:body//w:hyperlink`

### 代码示例
```python
hls = XA(doc_root, 'w:body//w:hyperlink')
for hl in hls:
    rid = GA(hl, 'r:id', '')
    txt = GT(hl)
    if '丝绸' in txt:
        score += 2
    if rid in main_rels:
        tgt = main_rels[rid]['target']
        if 'baike' in tgt.lower():
            score += 3
```

---

## 总结

| 要求 | 满分 | 关键XML标签 | 特殊注意 |
|--------|------|------------|----------|
| 1. 标题格式 | 10 | `w:rPr` | 字体/字号/颜色 |
| 2. 边框间距 | 10 | `w:bdr`, `w:spacing` | 字符边框非段落边框 |
| 3. 正文格式 | 10 | `w:pPr` | 检查pPr/rPr和run/rPr |
| 4. 第3段图片 | 10 | `w:drawing` | 单位转换（EMU） |
| 5. 脚注 | 5 | `w:footnotePr` | 编号格式 |
| 6. 水印 | 5 | `v:shape` | **VML**，非标准XML |
| 7. 分栏 | 10 | `w:cols` | 可能在段落标记的sectPr中 |
| 8. 页眉页脚 | 15 | `header*.xml`, `footer*.xml` | 奇偶页不同 |
| 9. 题注 | 10 | `w:drawing`+SEQ域 | 检查所有题注段落 |
| 10. 边框背景 | 10 | `w:pgBorders`, `w:background` | 阴影效果 |
| 11. 超链接 | 5 | `w:hyperlink` | 通过rels查找URL |

---

*最后更新：2026-07-03（全部检测改为 == 精确匹配，新增 VML形状、页面背景填充、表格行级边框）*
