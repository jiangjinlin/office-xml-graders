---
name: Word-XML-评分器
description: 通过解析.docx文件的底层Open XML结构，对Word 2010操作题进行精准评分。当用户需要自动评分或检查Word文档格式时使用此技能（如标题、边框、图片、水印、分栏、页眉页脚、题注、页面边框、超链接等）。通过直接检查XML而非使用python-docx等高层API，实现对标准答案文件100%检测通过率。
agent_created: true
---

# Word XML 评分器

通过直接解析.docx文件内部的Open XML结构，精准检测Word文档格式，实现操作题自动评分。

## 概述

本技能通过解析.docx文件内部的Open XML结构，精准检测Word文档格式，实现操作题自动评分。

**核心工作流：** Word比较文档生成修订标记 → 解析修订XML → 对照标准答案 → 生成评分代码

**核心原则：所有检测参数必须从标准答案XML中提取精确值，用 `==` 比较，不用 `in` 模糊匹配。**

## 何时使用此技能

当用户提出以下需求时使用：
- "评分Word文档" 或 "打分Word文件"
- "检查.docx格式"
- "做Word评分器" 或 "创建Word评分程序"
- "解析Word XML" 或 "查看docx内部结构"
- 任何需要检测python-docx无法处理的Word格式的任务

**关键指标：**
- 目标文件是.docx（Word 2007+）
- 需要检测高级格式（水印、环绕方式、旋转等）
- 需要高准确率（匹配"标准答案"文件）

## 工作流程

### 第一步：Word 生成修订版（用户操作）

1. Word → `审阅` → `比较` → `比较（比较文档的两个版本）`
2. **原文档** = `2素材.docx`，**修订的文档** = `1结果.docx`
3. 点开 `更多>>`，确认所有修订选项已勾选，修订显示位置选「原文档」
4. 确定 → Word 自动对比并在原文档上标记所有修订
5. 保存覆盖回 `2素材.docx`

**产出：** XML 中含 150+ 插入标记、100+ 删除标记、200+ 属性变更标记。

### 第二步：运行通用分析器 → 一步出报告

使用技能内置的 `scripts/analyze.py`，自动完成答案分析和修订分析：

```bash
python scripts/analyze.py N
# N = 文件夹编号，如 analyze.py 8
```

**输出：** `N/analysis.md` — 包含两部分：
1. **答案文档完整结构**（段落、节属性、表格、页眉页脚、超链接、编号）
2. **修订标记差异**（节属性变更、段落属性变更、运行属性变更、插入/删除内容）

这一步替换了之前每次手写的 200-500 行分析脚本，**代码量降为零**。

**修订元素速查：**

| 元素 | 语义 | 读取方法 |
|------|------|---------|
| `w:ins` | 插入的新内容 | `get_text(ins)` + 读内部 `w:rPr` 获取格式 |
| `w:del` | 删除的旧内容 | `get_text(del)` |
| `w:rPrChange` | 运行属性变更 | OLD = `rPrChange/rPr`，NEW = 外层 `rPr`（需跳过rPrChange） |
| `w:pPrChange` | 段落属性变更 | OLD = `pPrChange/pPr`，NEW = 外层 `pPr`（需跳过pPrChange） |
| `w:sectPrChange` | 节属性变更 | OLD = `sectPrChange/sectPr`，NEW = 外层 `sectPr` |

### 第三步：对照标准答案 → 提取精确值 → 写评分代码

**铁律：每个检测参数必须从标准答案 XML 中提取精确值，评分代码用 `==` 比较。**

修订标记告诉你“改了什么”，标准答案告诉你“精确值是多少”。二者结合生成评分代码：

| 检测项 | 从标准答案XML取精确值 | 评分代码示例 |
|--------|---------------------|-------------|
| 字体 | `w:eastAsia="华文新魏"` | `if font == '华文新魏':` |
| 字号 | `w:val="24"` (小四) | `if sz == '24':` |
| 颜色 | `w:color="FF0000"` | `if color.upper() == 'FF0000':` |
| 线型 | `w:val="double"` | `if val == 'double':` |
| 线宽 | `w:sz="8"` (1pt) | `if sz_v == '8':` |
| VML类型 | `type="#_x0000_t63"` | `if stype == '#_x0000_t63':` |
| VML填充色 | `fillcolor="white [3201]"` | `if fill == 'white [3201]':` |
| 样式ID | `w:val="2"` (标题2) | `if sid == '2':` |
| 底纹色 | `w:fill="DAEEF3"` | `if fill == 'DAEEF3':` |
| 超链接 | `Target="https://..."` | `if url == 'https://...':` |

**禁止使用 `in` 模糊匹配。** 即使标准答案中 `sz_v` 恰好是 `'8'`，也不能写成 `sz_v in ('4', '8')`。

### 第三步（产出A）：生成题目要求（必做）

修订标记 + 标准答案精确值 → 自动生成自然语言题目描述。写入 `N/题目要求_分析生成.md`。

**生成规则：**
- 节属性变更 → 页面设置类要求（纸张大小、页边距、行数、字符数）
- 段落的 `rPrChange` + `pPrChange` → 字体/字号/加粗/倾斜/颜色/对齐/间距类要求
- `w:ins` 插入的 VML 形状 → 图形插入类要求（标注类型、位置、文字、填充）
- `w:ins` 插入的文本 → 标题/正文内容类要求
- `w:ins` 插入的表格 → 文字转表格类要求
- 表格属性 → 边框/底纹/字体/题注类要求
- `w:ins` 的 `w:hyperlink` → 超链接类要求
- Header/Footer 变更 → 页眉页脚/页码类要求
- 水印（`v:shape` id含Watermark）→ 水印类要求

**示例输出：** `7/题目要求_分析生成.md`、`8/题目要求_分析生成.md`、`10/题目要求_分析生成.md`

### 第四步：生成评分代码 → 测试 → 100%

```bash
python word_grader_N.py  # 默认测试 1结果.docx
```

必须 **100/100**。不通过 → 检查哪个参数值不与标准答案XML一致 → 修正 → 重测。

---

### 备选流程：无修订标记时的手动分析

仅在无法获得修订对比文档时使用：解压标准答案 → 分析XML → 写检测代码（同样用 `==` 精确匹配）→ 测试 → 修复。

**与主流程的关键区别：** 缺少修订标记的“从A到B”语义，需靠人工判断哪些 XML 差异是有意义的考题改动。

## 核心辅助函数

评分器使用四个核心辅助函数（定义在`scripts/word_grader_template.py`中）：

```python
def T(tag):
    """转换'w:tag'为完整命名空间URI"""
    # 返回'{http://...}tag'

def X(parent, path):
    """查找单个元素，支持//递归"""
    # 示例：X(root, 'w:body//w:drawing')

def XA(parent, path):
    """查找所有匹配元素，支持//递归"""
    # 示例：XA(root, 'w:body//w:hyperlink')

def GA(elem, attr, default=None):
    """获取属性，处理命名空间——注意URI已含{}大括号！"""
    # ⚠️ 关键：ns_map值如'http://...'不含{}，用'%s%s'拼接
    # 错误写法：'{%s}%s' % (W, 'attr') → '{{http://...}}attr'（多一层括号）
    # 正确写法：'%s%s' % ('http://...', 'attr') → '{http://...}attr'
    if ':' in attr:
        ns, local = attr.split(':', 1)
        ns_map = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'v': 'urn:schemas-microsoft-com:vml',
            'o': 'urn:schemas-microsoft-com:office:office',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        }
        return elem.get('%s%s' % (ns_map.get(ns, ''), local), default)
    return elem.get(attr, default)
```

**重要：** 始终使用这些辅助函数，而不是手动`find()`/`get()`调用。

## 关键XML概念

### Open XML结构

.docx文件是包含XML文件的ZIP压缩包：

```
document.docx/
├── word/
│   ├── document.xml      ← 主文档
│   ├── settings.xml      ← 文档设置
│   ├── styles.xml        ← 样式定义
│   ├── footnotes.xml     ← 脚注
│   ├── header1.xml       ← 页眉（可能有多个）
│   ├── footer1.xml       ← 页脚
│   ├── diagrams/         ← SmartArt定义
│   ├── _rels/            ← 关系文件
│   └── media/            ← 嵌入图片
```

### 命名空间

Word使用多个XML命名空间。评分器定义：

```python
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
PIC = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
V = '{urn:schemas-microsoft-com:vml}'  # ⭐ 关键：水印使用VML
```

### VML水印

标准Open XML无法访问水印。它们存储为页眉文件中的VML形状。

**常见VML形状类型（type属性）：**

| type值 | 中文名 | shape.get('id') 特征 |
|--------|-------|---------------------|
| `#_x0000_t202` | 文本框 | "文本框 N" |
| `#_x0000_t63` | 椭圆批注 | "椭圆批注 N" |
| `#_x0000_t136` | 矩形/圆角矩形 | PowerPack |
| （水印） | 图片/文字水印 | id含"Watermark" |

```python
# 水印是id包含"Watermark"的VML形状
for shape in root.findall('.//{' + V + '}shape'):
    sid = shape.get('id', '')
    if 'Watermark' in sid:
        # 检查冲蚀效果
        for imgdata in shape.findall('{' + V + '}imagedata'):
            blacklevel = imgdata.get('blacklevel', '')
            gain = imgdata.get('gain', '')
```

### 文字水印

文字水印使用VML的`v:textpath`元素：

```python
# 在header中搜索文字水印
for shape in root.findall('.//{' + V + '}shape'):
    if 'Watermark' in shape.get('id', ''):
        textpath = shape.find('{' + V + '}textpath')
        if textpath is not None:
            print('水印文字:', textpath.get('string', ''))
```

### SmartArt

SmartArt引用在`document.xml`中，但文本在`word/diagrams/drawing1.xml`中：

```python
# 检测SmartArt引用
if 'diagrams' in self.doc_xml:
    print('找到SmartArt')

# 读取SmartArt文本
with zipfile.ZipFile(docx_path, 'r') as z:
    drawing = z.read('word/diagrams/drawing1.xml').decode('utf-8')
    if 'Metaverse' in drawing:
        print('SmartArt文本正确')
```

## 资源文件

本技能包含以下资源：

### scripts/

- **`analyze.py`** — ⭐ 通用修订分析器，一次调用替换 1254 行手写代码：
  ```bash
  python scripts/analyze.py N    # 分析文件夹 N 的答案+修订
  ```
  输出 `N/analysis.md`，含完整答案结构 + 修订差异报告。

- **`word_grader_template.py`** — 评分器模板：
  - `XMLGrader`类，包含`unzip()`, `load_all()`, `grade()`, `report()`方法
  - 示例`req1()`-`req11()`方法，覆盖常见Word格式要求
  - 辅助函数`T()`, `X()`, `XA()`, `GA()`
  - **重要：** `load_all()`方法会加载所有相关XML到`self.doc_xml`，用于字符串搜索

  **复用方法：** 复制此脚本并修改`req`方法以匹配具体要求。

### references/

- **`xml_patterns.md`** - XML模式快速参考：
  - 文件结构
  - 命名空间定义
  - 辅助函数代码
  - 常见XML模式（字符格式、段落格式、图片、分节属性、VML水印、VML文本框、段落边框、样式继承）
  - 单位转换（EMU、半角磅、1/20字符等）
  - 调试清单
  - **新增（2026-07-02）：** 段落边框vs字符边框、VML文本框检测、样式继承、垂直文本检测、表格格式

- **`grading_requirements.md`** - 11个示例要求的详细实现指南：
  - 检测内容（检查点）
  - XML位置
  - 代码示例

- **`algorithm_guide.md`** - 创建Word XML评分器的算法思路指南：
  - XML结构快速参考
  - 11种常见格式类型的检测逻辑
  - 调试技巧
  - 常见陷阱
  - 算法生成模板
  - **使用场景：** 创建新评分器时需要理解如何在XML中找到对应逻辑

## 高准确率技巧

1. **逆向工程答案文件：** 解压标准答案并检查其XML，理解确切的标签/属性使用方式。

2. **对嵌套元素使用递归（`//`）：** 图片在段落中嵌套很深。使用`X(dwg, '//wp:extent')`而不是`X(dwg, 'wp:extent')`。

3. **检查段落和运行属性：** 格式可以在段落级（`w:pPr/w:rPr`）或运行级（`w:r/w:rPr`）设置。都要检查。

4. **对水印使用VML：** 标准Open XML无法访问水印。使用VML命名空间（`urn:schemas-microsoft-com:vml`）。

5. **加载关系文件：** 图片引用（包括水印）存储在`_rels/*.xml.rels`文件中。加载它们以映射`r:id`到实际文件路径。

6. **先在标准答案上测试：** 评分器应该在标准答案文件上得100%分数。如果没有，调试直到通过。

7. **确保doc_xml完整：** `self.doc_xml`需要包含所有相关XML文件（document.xml, header*.xml, footer*.xml, diagrams/*.xml），才能用字符串搜索检测所有特征。

8. **使用宽松容差比较浮点数：** 尺寸比较需要用宽松容差（如0.2cm），避免5.10 vs 5.1的精度问题。

9. **XML字符串搜索作为兜底：** 当DOM解析复杂时，可以用XML字符串搜索作为兜底方案（如搜索'ellipse', 'softEdge'等）。

10. **题目要求的特定参数必须精确匹配：** URL、颜色值、字体名称、水印文字等明确指定的参数，用 `==` 精确比较，不能用 `in` 模糊匹配。

11. **警惕sectPr循环重复加分：** 文档可能有多个sectPr（如分栏段落产生段落级sectPr），`for sp in ...sectPr` 循环必须加 `break` 或 `found` 标志，防止同一要求重复计分。

12. **评分逻辑必须严格：** 如果某个参数不正确，总分必须相应扣减，不能让其他项的溢分掩盖错误（例如：URL错误但编号格式分有盈余时仍显示满分）。

13. **表格字体必须同时检查段落级和运行级 rPr**：表格单元格的字体经常在 `w:p/w:pPr/w:rPr`（段落级）而非 `w:r/w:rPr`（运行级）设置，两个级别都要遍历。

14. **页面背景图在w:background/v:fill中**：图片页面背景不在header水印中，而在 `w:background/v:background/v:fill` 元素链中。检查 `bg.iter(V+'fill')` 获取 `o:title`、`type="frame"`、`r:id`。

## 常见错误和解决方案

### 错误1：属性读取失败

**问题：** `elem.get('w:attr')`返回None

**原因：** 属性名需要带完整命名空间URI

**解决：** 使用`GA(elem, 'w:attr')`或`elem.get('{http://...}attr')`

### 错误2：水印检测失败

**问题：** 搜索document.xml找不到水印

**原因：** 水印在header*.xml中

**解决：** 在`load_all()`中把header/footer的XML也加入`self.doc_xml`

### 错误3：SmartArt文本检测失败

**问题：** 搜索document.xml找不到SmartArt文本

**原因：** SmartArt文本在`word/diagrams/drawing1.xml`中

**解决：** 把diagrams文件的XML也加入`self.doc_xml`

### 错误4：浮点数比较失败

**问题：** 5.10 vs 5.1比较失败

**原因：** 浮点数精度问题

**解决：** 使用宽松容差`abs(cx - 5.1) < 0.2`

### 错误5：cx/cy属性读取失败

**问题：** `GA(ext, 'cx')`返回0

**原因：** `cx`/`cy`属性**不带命名空间**（Word特例）

**解决：** 改用`ext.get('cx', '0')`

### 错误6：题注检测失败（在脚注中！）

**问题：** 搜索document.xml的所有段落找不到题注文字

**原因：** 图片题注可能以**脚注(footnote)形式**实现——脚注引用在图片段落中（`w:footnoteReference`），脚注内容在`word/footnotes.xml`中

**解决：** 
1. 加载`word/footnotes.xml`文件
2. 找到图片段落中的`w:footnoteReference w:id="N"`
3. 在footnotes.xml中查找对应id的脚注内容
4. 检查脚注格式（字体、字号、加粗）

### 错误7：页面边框/背景搜索失败

**问题：** 在`self.doc_xml`中搜索`pgBorders`找不到

**原因：** XML序列化后命名空间前缀被展开——`w:pgBorders` → `{http://...}pgBorders`

**解决：** 使用DOM直接查找而非字符串搜索：`sp.find(W + 'pgBorders')`

### 错误8：sectPr循环重复加分

**问题：** 标准答案100分，但改了某参数还是100分

**原因：** 文档有多个sectPr（如分栏产生段落级sectPr），`for sp in sectPr` 为每个sectPr都加一次分，溢分掩盖了其他错误

**解决：** 循环加 `break` 或 `found` 标志：`if found: break`，确保只计一次

### 错误9：PyInstaller打包后stdout崩溃

**问题：** exe启动报 `'NoneType' object has no attribute 'buffer'`

**原因：** `--noconsole`模式下 `sys.stdout.buffer` 为None，评分器模块的 `io.TextIOWrapper(sys.stdout.buffer)` 崩溃

**解决：** 创建带 `buffer` 属性的安全stdout代理类，导入前替换 sys.stdout

### 错误10：URL含in模糊匹配导致错误不扣分

**问题：** 改了URL仍然满分

**原因：** `'baike' in url` 只要URL含此文本就给分，不检查精确匹配

**解决：** 用 `url == 'https://baike.baidu.com/'` 精确比较，且该项分数占分足够大，确保不匹配时总分下降

### 错误11：GA函数属性全部返回默认值（严重！）

**问题：** `GA(elem, 'w:val', '0')` 永远返回 '0'，所有属性检测失败但代码无报错

**原因：** 命名空间URI如 `W = '{http://...}'` 已含大括号，用 `'{%s}%s' % (W, 'attr')` 会多一层括号变成 `'{{http://...}}attr'` —— 与XML中的 `'{http://...}attr'` 不匹配

**解决：** GA内部不能用 `{W}` 形式（W本身已有{}），改为用裸URI字符串+ `'%s%s'` 拼接：
```python
# 错误写法（W已含{}会导致双层括号）
return elem.get('{%s}%s' % (W, local), default)

# 正确写法（ns_map存裸URI无大括号）
ns_map = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
return elem.get('%s%s' % (ns_map[ns], local), default)
```

### 错误12：文本跨w:t分段致字符串搜索失败

**问题：** `'表格 1' in self.doc_xml` 返回 False，但文档中确实有此文字

**原因：** Word XML 中文本被切分到多个 `<w:t>` 元素中，如 `<w:t>表格</w:t><w:t> 1 </w:t><w:t>2010年</w:t>` —— 空白/数字经常独立分段，导致字符串搜索失效

**解决：** 用 DOM `get_text()` 逐段合并文本后再搜索，不要直接在原始XML字符串中查找中文短语

### 错误13：表格只检查运行级rPr忽略段落级rPr

**问题：** 表格单元格中调用 `cell.iter(W+'r')` 遍历运行时找不到字体设置

**原因：** 表格格式经常设置在段落级 rPr（`w:p/w:pPr/w:rPr`）而非运行级（`w:r/w:rPr`），继承链不同

**解决：** 表格字体检测必须同时检查两种级别：
```python
# 1. 运行级 rPr
for r in cell.iter(W + 'r'):
    rPr = child(r, W + 'rPr')
    if rPr: check_font(rPr)

# 2. 段落级 rPr（表格常见）
for p in children(cell, W + 'p'):
    pPr = child(p, W + 'pPr')
    if pPr:
        pRPr = child(pPr, W + 'rPr')
        if pRPr: check_font(pRPr)
```

### 错误14：标题格式在 pPr/rPr 而非 r/rPr（文件夹8）

**问题：** 检查 `paras[0]` 的子运行 `w:r/w:rPr` 找不到字体/字号/颜色

**原因：** 标题格式常设置在段落级 rPr（`w:pPr/w:rPr`），对所有运行生效，不在单个 `w:r` 内

**解决：** 标题检测先从 `pPr/rPr` 读取，再兜底检查 `r/rPr`

### 错误15：字符缩放用 w:w 而非 w:spacing（文件夹8）

**问题：** 搜索 `w:spacing` 检查字符缩放150%找不到

**原因：** 字符间距缩放是 `w:rPr/w:w w:val="150"`，元素名为 `w:w`，不是 `w:spacing`

**解决：** `child(rPr, W+'w')` 获取，`GA(we, 'w:val', '')=='150'`

### 错误16：变量Shadow — f变量被字体名字符串覆盖（文件夹10）

**问题：** `f=10`（满分）在循环中被 `f=GA(rPr...,'w:eastAsia','')` 覆盖为"宋体"，导致 `min(score, '宋体')` TypeError

**解决：** 满分变量用 `full`，字体变量用 `font_name`。**永远不要把循环内的临时变量命名为 `f`。**

### 错误17：编号格式 DOM 查找失效（文件夹10）

**问题：** `GA(nf,'w:val','')` 对 `numbering.xml` 的 numFmt 元素返回空

**原因：** numbering.xml 在某些 Word 版本中使用不同的 namespace 前缀声明

**解决：** 号码格式改用 `'chineseCountingThousand' in self.all_xml` 字符串搜索作为兜底

### 错误18：页眉多 Run 问题（文件夹10）

**问题：** 页眉第一个 `w:r` 的 rPr 为空或仅有 noProof，字体信息在后续 run 中

**解决：** 用 `for r in hroot.iter(W+'r')` 遍历所有 run，**不要 break 在第一个无 rPr 的 run 上**。只 break 在找到 rFonts 之后。

### 错误19：首行缩进值因字号而异（文件夹10）

**问题：** 题目说"缩进2字符"，用 `w:firstLine='480'`（小四默认值）检查失败

**原因：** `w:firstLine` 的单位是缇(twips)，与字号相关。五号=480、小四=496、四号=560

**解决：** 用 `w:firstLineChars`（字符单位）检查：`GA(ind,'w:firstLineChars','')=='200'`

## 一次性成功检查清单

每次写新评分器前，逐项确认以下规则，避免多次迭代：

### 代码层面
- [ ] `GA()` 用裸URI字符串 + `'{'+uri+'}'+local` 拼接，不用 `{%s}`
- [ ] 满分变量命名为 `full`，不用单字母（防止shadow）
- [ ] 所有参数比较用 `==`，禁用 `in`
- [ ] 循环检查 `sectPr`/`tbl` 时加 `break` 防止重复加分

### 检测顺序（按标准答案XML层级）
- [ ] **标题/正文字体**: 先查 `pPr/rPr`，再查 `r/rPr`，最后查 style 继承
- [ ] **表格字体**: 同时查运行级 `r/rPr` 和段落级 `pPr/rPr`
- [ ] **分栏**: `body.iter(W+'sectPr')` 递归查找，不限于直接子元素
- [ ] **页眉字体**: 遍历所有 `r`，不过早 break
- [ ] **缩进**: 用 `firstLineChars` 不用 `firstLine`
- [ ] **字符缩放**: 用 `w:w` 元素，不用 `w:spacing`
- [ ] **编号**: DOM失败时 `all_xml` 字符串兜底
- [ ] **水印**: header 的 VML `v:shape id含Watermark`
- [ ] **页面背景图**: `w:background/v:background/v:fill` 链
- [ ] **文本搜索**: 跨 `w:t` 分段时用 `get_text()` 合并，不直接搜 XML 字符串
- [ ] **段落定位**: 跳过标题和题注（`'表格' not in txt`）

### 标准产出
- [ ] `word_grader_N.py` — 评分代码
- [ ] `report_N.html` — 测试报告
- [ ] `题目要求_分析生成.md` — 自然语言题目描述

## 示例用法

**用户说：** "为这个操作题做Word评分器。答案文件是1结果.docx。"

**操作步骤：**

1. 解压`1结果.docx`并检查XML以理解要求
2. 复制`scripts/word_grader_template.py`到工作目录
3. 修改`req1()`-`req11()`以匹配11个要求
4. 运行`python word_grader.py "学生文件.docx"`
5. 检查输出：应打印得分并生成HTML报告

**结果：** 一个工作的Word评分器，可以准确评分学生提交文件。

---

**技能创建：** 2026-07-02  
**作者：** WorkBuddy AI  
**版本：** 3.7（2026-07-05）  
**更新内容：**
- v3.4：**新增 `scripts/analyze.py`** — 通用修订分析器（1条命令替代 1254 行手写代码）
- **工作流第二步**改为运行分析器 → 直接读 `analysis.md` 出题+评分
- 文件夹7/8/10 全部完成，积累 7 套评分器实战经验
- **3.6：Word 9 实战补全** — 文字边框/图片3D效果/样式级居中/表格全宽居中/字符缩放/行倍距
- **3.7：ISO 29500 架构级参考** — 复杂类型定义/简单类型枚举/默认值/度量单位矩阵/容差配置

---

## 架构级参考（ISO/IEC 29500:2016 字符级提取）

### CT_OnOff（布尔开关）默认行为

```python
# ⚠️ 关键：缺失 ≠ false！缺失 = 继承/默认
# <w:b />          → 粗体开启（默认 true）
# <w:b w:val="0"/> → 粗体关闭
# <w:b w:val="1"/> → 粗体开启
# 无 <w:b/>        → 继承段落/样式设置值
```

所有格式属性：b, i, u, strike, dstrike, outline, shadow, emboss, imprint, caps, smallCaps, vanish, webHidden, specVanish, bdr, shd 都遵循此规则。

### CT_Border 完整属性集
| 属性 | 数据类型 | 默认值 | 单位 |
|------|---------|--------|------|
| `w:val` | ST_Border | `"single"` | 枚举 |
| `w:sz` | ST_EighthPointMeasure | — | 1/8 磅 |
| `w:space` | ST_PointMeasure | `0` | 磅 |
| `w:color` | ST_HexColor | `"auto"` | RRGGBB |
| `w:shadow` | ST_OnOff | 省略=off | — |
| `w:frame` | ST_OnOff | 省略=off | — |
| `w:themeColor` | ST_ThemeColor | — | Office 2010+ |
| `w:themeTint` | ST_UcharHexNumber | — | Office 2010+ |
| `w:themeShade` | ST_UcharHexNumber | — | Office 2010+ |

### CT_Font 完整属性集 (w:rFonts)
| 属性 | 说明 | 示例 |
|------|------|------|
| `w:ascii` | 西文字体 | `"Arial Black"` |
| `w:hAnsi` | 高ANSI字体 | `"Arial"` |
| `w:eastAsia` | **东亚字体(中文)** | `"楷体"`, `"华文行楷"` |
| `w:cs` | 复杂脚本 | `"Times New Roman"` |
| `w:asciiTheme` | ASCII主题 | `"minorHAnsi"` |
| `w:eastAsiaTheme` | 东亚主题 | `"minorEastAsia"` |
| `w:hint` | 字体提示 | `"eastAsia"` |

### CT_Color 属性 (w:color)
```xml
<w:color w:val="C00000" />                       <!-- 直接颜色 -->
<w:color w:val="C00000" w:themeColor="accent2"/>  <!-- 颜色+主题映射 -->
<w:color w:themeColor="accent2" w:themeTint="99"/> <!-- 主题色着色 -->
```
检测时优先用 `val` 直接比较十六进制；`themeColor` 为可选回退。

### 度量单位矩阵（按 XML 属性）
| 元素 | 属性 | 单位 | 换算 |
|------|------|------|------|
| `w:sz` | w:val | 半磅 | ÷2 = pt |
| `w:szCs` | w:val | 半磅 | ÷2 = pt |
| `w:spacing` | w:before/w:after | TWIPS | ÷20 = pt, ÷567 = cm |
| `w:spacing` | w:line | TWIPS/行 | `"240"`=单倍, `"360"`=1.5倍, `"400"`=20pt固定 |
| `w:spacing` | w:beforeLines/w:afterLines | 1/100 行 | `"50"`=0.5行 |
| `w:ind` | w:left/w:right/w:firstLine | TWIPS | ÷20 = pt |
| `w:ind` | w:leftChars/w:rightChars/w:firstLineChars | 1/100 字符 | `"200"`=2字符 |
| `w:bdr` | w:sz | 1/8 磅 | ÷8 = pt |
| `w:pgSz` | w:w, w:h | TWIPS | ÷1440 = inch |
| `w:pgMar` | w:top等 | TWIPS | ÷567 = cm |
| `w:tblW` | w:w (type=pct) | 1/50% | `"5000"`=100% |
| `w:tblW` | w:w (type=dxa) | TWIPS | ÷20 = pt |

### 容差配置建议
| 检测类别 | 单位 | 建议容差 |
|----------|------|---------|
| 字号 (w:sz) | 半磅 | ±0（精确） |
| 行距 (w:line) | TWIPS | ±20（≈1pt） |
| 缩进 (w:ind) | TWIPS | ±100（≈5pt） |
| 页边距 (w:pgMar) | TWIPS | ±50（≈0.09cm） |
| 图片尺寸 (wp:extent) | EMU | ±180000（≈0.5cm） |
| 颜色 RGB | 十六进制 | 不区分大小写 (upper) |
| 字体名 | 字符串 | 直接 == 或包含 in |

---

## 新增陷阱（Word 9 实战）

### 陷阱20：文字边框 ≠ 段落边框
```python
# ❌ 错误：搜 w:pBdr
if pPr.find(f'{{{W}}}pBdr') is not None:
    → 标题的"文字方框"可能不在这里

# ✅ 正确：文字方框存储在 rPr 的 w:bdr 元素中
# XML: <w:rPr><w:bdr w:val="single" w:sz="12" w:space="0" w:color="0070C0"/></w:rPr>
bdr = rPr.find(f'{{{W}}}bdr')
if bdr is not None:
    val = GA(bdr, 'w:val', '')   # "single"
    sz = GA(bdr, 'w:sz', '')     # "12" = 1.5pt (÷8)
    cl = GA(bdr, 'w:color', '')  # "0070C0" 蓝色
```
`w:bdr` 是字符级边框，"文字方框"用这个；`w:pBdr` 是段落级边框。

### 陷阱21：图片3D效果 ≠ prstMaterial
```python
# Word 2010+ 的棱台/亚光效果存储在 scene3d/sp3d/bevelT 中
# XML: <a:scene3d><a:sp3d contourW="6350">
#         <a:bevelT w="50800" h="16510"/>
#         <a:contourClr><a:srgbClr val="C0C0C0"/></a:contourClr>
#       </a:sp3d></a:scene3d>

# ✅ 用字符串搜索，不依赖单个元素名
if 'scene3d' in dwg_xml or 'sp3d' in dwg_xml or 'bevelT' in dwg_xml:
    → 棱台效果存在
```

### 陷阱22：页眉居中通过样式继承
```python
# 页眉 XML 中常见: <w:pPr><w:pStyle w:val="a3"/></w:pPr>
# — 没有 w:jc！但 a3 是内置页眉样式，默认居中

# ✅ 检查 jc 时同时检查 pStyle
pStyle = pPr.find(f'{{{W}}}pStyle')
if pStyle is not None:
    sv = GA(pStyle, 'w:val', '')
    if sv in ('a3', 'a4', 'a5'):  # 内置页眉/页脚样式默认居中
        → 居中判定通过
```

### 陷阱23：表格居中 → tblW auto 等价
```python
# 表格 XML 中无 <w:jc w:val="center"/> 但视觉上居中
# 原因: <w:tblW w:type="auto" w:w="0"/> → 表格自动填满页宽=等价居中

# ✅ 三级检测
if tblPr.find(f'{{{W}}}jc') is not None:            # 显式居中
    → score += 2
elif 父段落 pPr/jc == "center":                      # 父段落居中
    → score += 2
elif tblW type="auto" and w="0":                     # 全宽等价居中
    → score += 2
```

### 陷阱24：beforeLines/afterLines 行倍距
```xml
<!-- 段前0.5行 + 段后1行 = beforeLines="50" afterLines="100" -->
<w:spacing w:beforeLines="50" w:before="156" w:afterLines="100" w:after="312"/>
```
`beforeLines`/`afterLines` 以 1/100 行为单位。同时存在 `before`/`after`（TWIPS备份值）。

### 陷阱25：字符缩放 w:w ≠ w:spacing
```xml
<!-- 字符间距缩放150% → w:rPr/w:w val="150" -->
<w:rPr>
  <w:w w:val="150"/>      ← 这是缩放比例(%)
  <w:spacing w:val="8"/>  ← 这是字符间距(TWIPS)
</w:rPr>
```
`w:w` 控制缩放，`w:spacing` 控制间距。两者不同，不要混淆。

### 陷阱26：段落索引从正文算起
"正文第2段" = 跳过标题，从第二段正文开始。`paras[2]` 不是 `paras[1]`。

## 检查清单（新增）
- [ ] 标题边框: 先查 pPr/rPr/w:bdr（文字方框），再查 w:pBdr（段落边框）
- [ ] 图片3D: 搜 `scene3d`/`sp3d`/`bevelT`，不限于 `prstMaterial`
- [ ] 页眉居中: 检查 pStyle val 是否为 a3/a4/a5（内置页眉样式）
- [ ] 表格居中: tblPr/jc → 父段落/jc → tblW auto=0 三级兜底
- [ ] 段间距: 用 beforeLines/afterLines（1/100行单位）
- [ ] 字符缩放: w:rPr/w:w val（%值）
- [ ] 段落编号: 标题不计入"第N段"

---

## 精确 XML 属性参考（基于 ISO/IEC 29500:2016）

### 字体属性 (`w:rPr` 子元素)

| 元素 | 属性 | 说明 | 常见值 |
|------|------|------|--------|
| `<w:rFonts/>` | `w:eastAsia` | 东亚字体 | `"华文楷体"`, `"宋体"`, `"华文行楷"`, `"黑体"` |
| | `w:ascii` | 西文字体 | `"Calibri"`, `"Arial"` |
| | `w:hAnsi` | 高ANSI字体 | 同 ascii |
| | `w:cs` | 复杂脚本字体 | |
| `<w:b/>` | `w:val` | 加粗 | `"true"`, `"1"` (可自闭合) |
| `<w:i/>` | `w:val` | 倾斜 | `"true"`, `"1"` |
| `<w:u/>` | `w:val` | 下划线 | `"single"`, `"double"`, `"dotted"`, `"wave"`, `"none"` |
| `<w:strike/>` | `w:val` | 删除线 | `"true"` |
| `<w:color/>` | `w:val` | 字体颜色 | RRGGBB 六位十六进制 `"FF0000"`(红), `"002060"`(深蓝) |
| | `w:themeColor` | 主题色 | `"accent1"`~`"accent6"` |
| `<w:sz/>` | `w:val` | 字号(半磅) | `22`=11pt, `24`=12pt(小四), `28`=14pt(四号), `32`=16pt(三号), `36`=18pt(小二), `44`=22pt(二号) |
| `<w:szCs/>` | `w:val` | 复杂脚本字号 | 同上 |
| `<w:highlight/>` | `w:val` | 高亮色 | `"yellow"`, `"cyan"`, `"red"`, `"none"` |
| `<w:vertAlign/>` | `w:val` | 上下标 | `"superscript"`, `"subscript"` |
| `<w:spacing/>` | `w:val` | 字符间距 | 正数=扩展, 负数=紧缩(TWIPS) |
| `<w:position/>` | `w:val` | 字符位置 | 正=提升, 负=降低(半磅) |
| `<w:shd/>` | `w:fill` | 文字底纹填充 | `"FFC000"`(橙色), `"DAEEF3"`(浅蓝) |
| | `w:val` | 底纹类型 | `"clear"`, `"solid"` |
| `<w:lang/>` | `w:eastAsia` | 语言 | `"zh-CN"` |
| | `w:val` | | `"en-US"` |

### 段落属性 (`w:pPr` 子元素)

| 元素 | 属性 | 说明 | 常见值 |
|------|------|------|--------|
| `<w:jc/>` | `w:val` | 水平对齐 | `"left"`, `"center"`, `"right"`, `"both"`(两端) |
| `<w:spacing/>` | `w:before` | 段前间距(TWIPS) | `120`=5pt, `240`=10pt |
| | `w:after` | 段后间距 | 同上 |
| | `w:line` | 行距值 | `240`=单倍, `360`=1.5倍, `400`=20pt固定 |
| | `w:lineRule` | 行距规则 | `"auto"`, `"exact"`(固定), `"atLeast"` |
| `<w:ind/>` | `w:left` | 左缩进(TWIPS) | `567`≈2字符(小四) |
| | `w:right` | 右缩进 | 同上 |
| | `w:firstLine` | 首行缩进 | `480`(五号2字符), `560`(四号2字符) |
| | `w:firstLineChars` | 首行缩进(字符) | `"200"`=2字符 |
| `<w:pBdr/>` | 子元素 | 段落边框 | 含 `<w:top>`, `<w:left>`, `<w:bottom>`, `<w:right>` |
| | `w:val` | 边框线型 | `"single"`, `"double"`, `"dashed"`, `"dotted"`, `"thick"` |
| | `w:sz` | 线宽(1/8pt) | `6`=0.75pt, `8`=1pt, `12`=1.5pt, `24`=3pt |
| | `w:color` | 线色 | `"0000FF"`(蓝), `"FF0000"`(红) |
| | `w:shadow` | 阴影 | `"1"` |
| `<w:shd/>` | `w:fill` | 段落底纹 | `"FFC000"` |
| | `w:val` | 底纹模式 | `"clear"`(清除), `"solid"`(纯色) |
| `<w:numPr/>` | | 编号属性 | `w:ilvl`(级别), `w:numId`(编号ID) |
| `<w:outlineLvl/>` | `w:val` | 大纲级别 | `"0"`(正文)~`"9"` |
| `<w:keepLines/>` | | 段中不分页 | 布尔 |
| `<w:pageBreakBefore/>` | `w:val` | 段前分页 | `"true"` |

### 节属性 (`w:sectPr` 子元素)

| 元素 | 属性 | 说明 | 常见值 |
|------|------|------|--------|
| `<w:pgSz/>` | `w:w` | 页面宽度(TWIPS) | `11906`=A4, `12240`=Letter |
| | `w:h` | 页面高度 | `16838`=A4 |
| | `w:code` | 纸张编号 | `9`=A4 |
| | `w:orient` | 方向 | `"portrait"`, `"landscape"` |
| `<w:pgMar/>` | `w:top` | 上边距(TWIPS) | `1440`=1in, `1418`≈2.5cm |
| | `w:bottom` | 下边距 | 同上 |
| | `w:left` | 左边距 | `1701`≈3cm |
| | `w:right` | 右边距 | 同上 |
| | `w:gutter` | 装订线 | `340`≈0.6cm |
| `<w:cols/>` | `w:num` | 分栏数 | `"1"`, `"2"`, `"3"` |
| | `w:space` | 栏间距(TWIPS) | |
| | `w:sep` | 分隔线 | `"1"`=有, `"0"`=无 |
| | `w:equalWidth` | 等宽 | `"1"`=是, `"0"`=否 |
| `<w:docGrid/>` | `w:linePitch` | 每页行数 | `320`≈44行, `326`≈42行 |
| | `w:charSpace` | 字符间距 | |

### 表格属性 (`w:tblPr`, `w:tcPr`)

| 元素 | 属性 | 说明 | 常见值 |
|------|------|------|--------|
| `<w:tblBorders/>` | 含 `w:top`, `w:left`, `w:bottom`, `w:right`, `w:insideH`, `w:insideV` | | |
| `<w:top/>` | `w:val` | 线型 | `"single"`, `"double"` |
| | `w:sz` | 线宽(1/8pt) | `12`=1.5pt |
| | `w:color` | 颜色 | `"1F497D"`(深蓝), `"B4C6E7"`(浅蓝) |
| `<w:jc/>` | `w:val` | 表格对齐 | `"center"`, `"left"` |
| `<w:tblW/>` | `w:type` | 列宽类型 | `"auto"`, `"pct"`(百分比), `"dxa"` |
| `<w:shd/>` | `w:fill` | 单元格底纹 | `"D9E2F3"`(浅蓝灰) |
| `<w:vAlign/>` | `w:val` | 垂直对齐 | `"top"`, `"center"`, `"bottom"` |

### 文本框/形状 VML 属性 (`v:shape`, `v:textbox`)

| 属性 | 说明 | 常见值 |
|------|------|--------|
| `id` | 形状标识 | `"文本框 2"`, 含水印则含 `"Watermark"` |
| `type` | VML形状类型 | `"#_x0000_t202"`(文本框), `"#_x0000_t136"`(圆角矩形) |
| `strokecolor` | 边框颜色 | `"#c00000"`(红), `"#1f497d"`(深蓝) |
| `strokeweight` | 边框磅值 | `"2pt"` |
| `fillcolor` | 填充色 | `"white [3201]"` |
| `style` | CSS样式 | 含宽/高/定位信息 |

### Watermark 检测要点

```xml
<!-- 图片水印: header XML → v:shape id="PowerPlusWaterMarkObjectxxx" -->
<v:shape id="wordpicturewatermark329474312" ...>
  <v:imagedata blacklevel="13107f" gain="19661f"/>  <!-- 冲蚀=降低对比度 -->
</v:shape>

<!-- 文字水印: header XML → v:textpath -->
<v:textpath string="科技日报" style="font-family:微软雅黑;font-size:1pt"/>
```

### 单位换算

| 单位 | 到 TWIPS | 到 cm |
|------|----------|-------|
| 1 TWIP | 1 | 1/567 |
| 1 磅 (pt) | 20 | 1/28.35 |
| 1 英寸 (in) | 1440 | 2.54 |
| 1 cm | 567 | 1 |
| 1 字符(小四) | ~280 TWIPS | ~0.5 |

### 字号 → half-point → pt 对照表

| 字号 | val(half-pt) | pt |
|------|-------------|-----|
| 五号 | 21 | 10.5 |
| 小四 | 24 | 12 |
| 四号 | 28 | 14 |
| 小三 | 30 | 15 |
| 三号 | 32 | 16 |
| 小二 | 36 | 18 |
| 二号 | 44 | 22 |
| 一号 | 52 | 26 |
| 小初 | 72 | 36 |

---

## 边框线型完整枚举（ISO/IEC 29500 BorderValues）

```
Nil, None, Single, Dashed, Dotted, Double,
Thick, Wave, DoubleWave, DashDotStroked,
DashSmallGap, DotDash, DashDotDot, DotDotDash,
Triple, ThinThickSmallGap, ThickThinSmallGap, ThinThickThinSmallGap,
ThinThickMediumGap, ThickThinMediumGap, ThinThickThinMediumGap,
ThinThickLargeGap, ThickThinLargeGap, ThinThickThinLargeGap,
DashedSmallGap, ThreeDEmboss, ThreeDEngrave, Outset, Inset
```

**评分最常用**: `Single`(单实线), `Double`(双实线), `Dashed`(虚线), `Dotted`(点线), `Thick`(粗线), `None`(无)

### 线宽对照表 (w:sz, 1/8pt 单位)
| sz | pt |
|----|-----|
| 4 | 0.5 |
| 6 | 0.75 |
| 8 | 1.0 |
| 12 | 1.5 |
| 18 | 2.25 |
| 24 | 3.0 |
| 36 | 4.5 |

### Settings 属性速查
| 属性 | 说明 |
|------|------|
| `<w:evenAndOddHeaders/>` | 奇偶页不同 |
| `<w:mirrorMargins/>` | 对称页边距 |
| `<w:noLineBreaksAfter/>` | 避头尾 |
| `<w:noLineBreaksBefore/>` | 避尾 |
| `<w:defaultTabStop/>` | 默认制表位 |
| `<w:autoHyphenation/>` | 自动断字 |

### 修订跟踪 w:rPrChange 属性速查
```xml
<!-- 运行属性变更: OLD=rPrChange的子rPr, NEW=外层rPr -->
<w:rPr>
  <w:rPrChange w:id="1" w:author="xx" w:date="2024-01-01T00:00:00Z">
    <w:rPr>  ← OLD 值
      <w:rFonts w:eastAsia="宋体"/>
      <w:sz w:val="21"/>
    </w:rPr>
  </w:rPrChange>
  <w:rFonts w:eastAsia="华文楷体"/>  ← NEW 值
  <w:sz w:val="24"/>
  <w:b/>
</w:rPr>
```

### 完整单位换算公式
| 转换 | 公式 |
|------|------|
| inch → TWIPS | ×1440 |
| inch → EMU | ×914400 |
| pt → TWIPS | ×20 |
| pt → EMU | ×12700 |
| cm → TWIPS | ×567 |
| cm → EMU | ×360000 |
| 字号半磅值 → pt | ÷2 (val="24"=12pt) |
| 边框1/8pt → pt | ÷8 (val="12"=1.5pt) |
| DrawingML 字号 → pt | ÷100 (val="2800"=28pt) |
| tint/shade 百分比 → 值 | ×1000 (50%=50000) |
