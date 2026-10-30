---
name: Word-XML-评分器
description: 通过解析.docx文件的底层Open XML结构，对Word 2010操作题进行精准评分。当用户需要自动评分或检查Word文档格式时使用此技能（如标题、边框、图片、水印、分栏、页眉页脚、题注、页面边框、超链接等）。通过直接检查XML而非使用python-docx等高层API，实现对标准答案文件100%检测通过率。
agent_created: true
---

# Word XML 评分器

通过直接解析.docx文件内部的Open XML结构，精准检测Word文档格式，实现操作题自动评分。

## 权威参考（查结构先看这里）

- 规范：ECMA-376 Part 1 §17（WordprocessingML）。页眉页脚 §17.6.5、节属性 §17.6.17、
  页码域 §17.16.5、表格 §17.4、编号 §17.9。
- MS 文档：https://learn.microsoft.com/zh-cn/office/open-xml/word/overview
  SDK 类名即元素名：`SectionProperties`=`w:sectPr`、`FieldCode`=`w:instrText`。
- ooxml.org：https://ooxml.org/wordml/ 有 210 个元素的父/子/属性/条款号，查结构最顺手；
  但它是前端渲染的，脚本抓正文要带渲染，`urllib` 只能拿到空壳。
- 本地速查表：`student_exam_app/tools/reference/OOXML-速查.md`。

## 概述

本技能通过解析.docx文件内部的Open XML结构，精准检测Word文档格式，实现操作题自动评分。

**核心工作流：** Word比较文档生成修订标记 → 解析修订XML → 对照标准答案 → 生成评分代码

**核心原则：所有检测参数必须从标准答案XML中提取精确值，用 `==` 比较，不用 `in` 模糊匹配。**

## 考点覆盖表（按题库 10 套实测，缺什么一眼看到）

| 考点 | XML 证据 | 判定要点 | 已有实现 |
| --- | --- | --- | --- |
| 邮件合并 | `settings.xml` 的 `w:mailMerge` + `w:mainDocumentType@val="formLetters"` + `w:connectString`；正文 `w:instrText` 里的 `MERGEFIELD 姓名` | 三件齐才算：合并开关 + 主文档类型 + 合并域。数据源路径里应出现目标表名（不同学生路径不同，只比文件名） | `_core.mail_merge_info` / `merge_fields`，word10 |
| 图片尺寸 | `wp:extent@cx/@cy`，单位 EMU，1 厘米 = 360000 | 7 厘米 = 2520000；只查 `w:drawing` 存在不够，必须比 extent | word2/4/6/7/9 |
| 图片环绕 | `wp:anchor` 下的 `wp:wrapSquare` / `wrapTopAndBottom` / `wrapTight` / `wrapThrough` / `wrapNone` | `wp:inline` = 嵌入型，没有 wrap 子元素；"上下型"是 `wrapTopAndBottom`，"四周型"是 `wrapSquare` | word2 已按此判 |
| 图片裁剪/亮度 | `a:srcRect@l/t/r/b`、`a:lum@bright/@contrast` | 裁剪比例是千分比（`l="25000"` = 裁掉 25%） | 题库未考，新题按此写 |
| 旧式图片(VML) | `w:pict > v:shape > v:imagedata@r:id` | 剪贴画、老模板图走 VML，`w:drawing` 里找不到，必须两路都查 | `_core.pictures` 已兼容 |
| 艺术字 | `w:pict > v:shape` 里带 `v:textpath@string`，或 DrawingML 的 `a:prstTxWarp` | 文字内容在 `textpath/@string`，不在 `w:t` 里 | word2、word7 |
| 文本框 | `w:txbxContent`；竖排看 `w:textDirection@val="tbRlV"` | 文本框里的字用 `gt()` 也能取到，但样式要进 `txbxContent` 里查 | word2 |
| 形状/椭圆标注/箭头 | VML `v:shape@type="#_x0000_t9"`（椭圆）、`t13`（箭头）；或 `a:prstGeom@prst="ellipse"/"rightArrow"` | 改颜色看 `v:fill@color` 或 `a:solidFill` | word7 |
| SmartArt | `word/diagrams/data*.xml` + `dgm:relIds` | 内容在 diagrams 部件里，document.xml 只有引用；判文字要读 data 部件 | word4 |
| 公式 | `m:oMath`（新版）或域代码 `= ...`（老版公式） | `oMath` 在 `w:p` 里，命名空间是 math，不是 w | word7 用域判 |
| 题注 | `w:fldSimple@w:instr=" SEQ 图 \\* ARABIC "` 或段落里的 `Caption` 样式 | 只查"图 1"文字会出现假阳性 | word1/7/8 |
| 水印 | 页眉里的 `v:shape@id="PowerPlusWaterMarkObject..."` + `v:textpath@string` | 文字水印在页眉部件；`w:color` 和 `rotation:-45` 控制半透明/角度 | word1/2/4/5/6/10 |
| 脚注尾注 | `w:footnoteReference@w:id` + `footnotes.xml` 的 `w:footnote` | 编号在 `w:footnote@w:id`，正文只留引用 | word1/5/6/8/9 |
| 首字下沉 | `w:framePr@w:dropCap="drop"` + `w:lines` | 题库未考，新题按此写 | 待补 |
| 分栏 | `w:cols@w:num`，等宽看 `w:equalWidth`，分隔线 `w:sep` | 多节文档要按节取，不能只看第一个 sectPr | word2/5/9 |
| 页面边框 | `w:pgBorders` + 四个 `w:top/bottom/left/right@w:val/@w:sz/@w:color` | `w:offsetFrom="page"` 才是页面边框，`text` 是文字边框 | word9 |
| 目录 | `w:instrText` 里 `TOC \\o "1-3"`，或 `w:sdt` 包着的目录 | 题库未考，新题按此写 | 待补 |

## 图片与形状判定的三个坑

1. **同一张图有两种存法**：Office 2010 默认 DrawingML（`w:drawing`），但剪贴画、老模板、部分
   粘贴操作落成 VML（`w:pict`）。只查一种就漏判，`_core.pictures()` 两路都收。
2. **`wp:extent` 是 EMU**：1 厘米 360000、1 磅 12700、1 英寸 914400。题干写"高 7 厘米"
   要换算成 2520000，允许 ±5% 误差（缩放会四舍五入）。
3. **环绕方式看父子关系**：`wp:inline` 里的 `extent` 是嵌入型；`wp:anchor` 里的
   `behindDoc="1"` 是衬于文字下方，`wrapTopAndBottom` 才是上下型。只看 `anchor` 存在会误判。

## 邮件合并判定细则（word10 实战）

```
settings.xml
<w:mailMerge>
  <w:mainDocumentType w:val="formLetters"/>      信函
  <w:dataType w:val="native"/>                   Excel/Access 数据源
  <w:connectString w:val="Provider=Microsoft.ACE.OLEDB.12.0;...Data Source=...学生成绩表.xlsx;..."/>
  <w:query w:val="SELECT * FROM `学生成绩表$`"/>
</w:mailMerge>

document.xml
<w:instrText> MERGEFIELD 姓名 </w:instrText>     每个域一对 instrText
```

判定顺序：① `w:mailMerge` 存在（3 分）② mainDocumentType=formLetters（2 分）
③ 数据源里出现目标表名（2 分）④ MERGEFIELD 数量 ≥5（1.5 分）⑤ 域名覆盖
姓名/语文/数学/英语/物理/化学/总分 ≥6 个（1.5 分）。

坑：不能搜 "MERGEFIELD" 字符串就完事——学生可能只插了一个域；也不能要求数据源路径
完全一致，每人保存的绝对路径都不同，只比文件名。

## 新补的陷阱（2026-10-30，对照 ECMA-376 校验后）

- **页码域别搜字符串**：`'PAGE' in xml` 会命中 `NUMPAGES`、正文里的英文单词。
  统一用 `_core.checks.has_field(xml, 'PAGE')`，它按 `w:instrText` 和
  `w:fldSimple@w:instr` 里的域代码判断。word1/2/4/6/8 都已切换。
- **ET.tostring 会加 `ns0:` 前缀**：解析后再序列化，正则里的 `<w:instrText` 就匹配不到，
  has_field 的正则写成 `<(?:\w+:)?instrText`，同时兼容两种。
- **页眉页脚有三态**：`w:headerReference@w:type` 是 default/first/even；
  奇偶页不同还要 `settings.xml` 里有 `w:evenAndOddHeaders`，只看 header 部件数量会误判。
- **字符单位与 twips 别混**：`w:ind@w:firstLineChars="200"` 是 2 字符，
  而 `w:pgMar@w:top="1418"` 是 2.5cm（twips）。判断容差按各自单位取。
- **样式链要合并**：docDefaults → 样式表 basedOn 链 → pPr/rPr → run 的 rPr，
  只看 run 会漏掉样式给的字体字号（word3 的"小四"就来自 a3 样式）。
- **底纹看 themeFill+tint**：`w:shd@w:fill` 为空而 `w:themeFill="background2" w:themeFillTint="80"`
  才是"强调文字颜色5 淡色80%"，只比 fill 会判错。

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

## 统一内核（新评分器直接用）

`student_exam_app/graders/_core/` 里的 `WordDoc` 已经把段落、样式继承、节、页眉页脚、表格读完：

```python
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import WordDoc, Report, extract_word_parts

doc = WordDoc(path)
for p in doc.paragraphs:
    p.index, p.text, p.style
    p.props['jc'], p.props['firstLineChars'], p.props['beforeLines']
    p.runs[0].props['font'], p.runs[0].props['size'], p.runs[0].props['bold']

doc.page_setup()            # 纸张/页边距/分栏/行数，默认取 body 级 sectPr
doc.header_texts(); doc.footer_texts()
doc.tables[0].cell_texts(); doc.numbering; doc.instr_texts()
doc.text_all()              # 正文+页眉页脚+脚注尾注+表格，用于全文查找
```

`run.props` 已经把 docDefaults → 样式链（basedOn）→ 段落级 `pPr/rPr` → 运行级 `rPr` 合并过一遍，不用再自己判断"字体到底设在哪一级"。

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

### 错误20：`extractall()` 每次判分解出上百个文件

**问题：** 判分服务跑久了 `uploads/` 里堆满 `_xml_学生文件/word/media/*`，每份 docx 解出 70+ 个文件；清理时还会被批量删除保护拦住。

**原因：** `z.extractall(dir)` 把图片、缩略图、docProps 全解出来了，评分根本用不到 `word/media/`。

**解决：** 只解评分需要的 XML：

```python
from _core import extract_word_parts
extract_word_parts(self.docx, self.xmldir)
# 文档/样式/设置/编号/脚注/尾注 + header*.xml + footer*.xml + _rels/*.rels + diagrams/theme
```

实测第 2 套从 70 个文件降到 15 个左右，判分更快，也不再触发清理限制。`self.media` 只是打印用，留空不影响任何评分项。

### 错误21：`page_setup()` 取错了 sectPr

**问题：** 分栏文档里 `for sp in doc.iter(sectPr)` 有 3 个，最后一个可能是段落级 sectPr（分栏用），页边距和题目要求对不上。

**解决：** 优先取 body 的直接子元素：

```python
sect = doc.body.find(q('sectPr'))   # body 级 = 全文档页面设置
```

`_core.WordDoc.page_setup()` 已按这个顺序取（body 级 → 最后一级兜底）。

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

## 回归与部署

```bash
python tools/selftest.py     # 全部评分器对标准答案跑一遍，必须全 100
python tools/core_sweep.py   # 核心库遍历 10 套 Word 答案，确认无解析异常
python deploy.py --dry-run   # 发布前预检 + 打包
```

Word 评分器的满分常见是 110（每项分值加起来不为 100），`selftest.py` 和 `app.py` 都按百分比折算，不要为了凑 100 去改分值——改分值会让每项的权重失衡。

---

**技能创建：** 2026-07-02  
**作者：** WorkBuddy AI  
**版本：** 3.8（统一内核 + 解压/分节修正 + 回归部署）  
**更新内容：**
- 3.8：`_core.WordDoc` 用法（样式链合并、body 级 sectPr、页眉页脚文本）；错误20（extractall 解出上百文件）、错误21（多 sectPr 取错）；回归与部署命令
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

## 判分点全表（含题库未考但可能出现的）

> 完整版（三科合并）在 `student_exam_app/tools/reference/判分点全表.md`；这里只放本科目。新题先对表，再写代码。

### 文档与节

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `w:pgSz@w/h/orient` | 纸张大小方向 | A4=11906×16838 twips；横向看 `orient="landscape"`（宽高会互换，别只比 w） |
| `w:pgMar@top/right/bottom/left/header/footer/gutter` | 页边距 | 2.8 厘米=1588、3 厘米=1701；`header/footer` 是页眉页脚距边界，不是边距 |
| `w:pgBorders` + 四条 `w:top…` | 页面边框 | `offsetFrom="page"` 是页面边框，`text` 是文字边框；艺术型看 `w:val` |
| `w:cols@num/space/sep/equalWidth` | 分栏 | 多节文档要按 `sectPr` 逐节取；`w:cols` 出现在 `sectPr` 里，不在 pPr |
| `w:docGrid@type/linePitch/charSpace` | 文档网格 | "每页 42 行每行 39 字" = `type="linesAndChars" linePitch=… charSpace=…` |
| `w:pgNumType@fmt/start` | 页码格式/起始 | 起始页 1 = `start="1"`；罗马数字看 `fmt="lowerRoman"` |
| `w:titlePg` | 首页不同 | 必须配合 `headerReference@type="first"` 才生效 |
| `w:evenAndOddHeaders`（settings） | 奇偶页不同 | 在 settings.xml，不在 document.xml；缺它则 even 页眉不生效 |
| `w:sectType@val` | 分节符类型 | `nextPage`/`continuous`/`evenPage`/`oddPage` |

### 段落与字符

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `w:pStyle@val` | 段落样式 | 样式名是 styleId，中文名在同 id 的 `w:name`；"标题 1" 的 id 常是 `1` |
| `w:jc@val` | 对齐 | 两端对齐=`both`，分散对齐=`distribute`；不要拿 `center` 一把梭 |
| `w:ind@firstLine/firstLineChars/left/right/hanging` | 缩进 | `firstLineChars="200"` 是 2 字符；混用 twips 与字符单位会差 100 倍 |
| `w:spacing@before/after/line/lineRule` | 段前段后/行距 | 固定值 18 磅 = `line="360" lineRule="exact"`；倍数行距是 `lineRule="auto"` + `line="240"×倍数` |
| `w:snapToGrid` | 对齐网格 | 关闭才允许非整数行距，题干说"固定值"时答案常关它 |
| `w:keepNext/keepLines/pageBreakBefore/widowControl` | 分页控制 | 常见"标题不与正文分离"=`keepNext` |
| `w:outlineLvl@val` | 大纲级别 | 0=1 级；目录靠它抓取，样式标题也自带 |
| `w:numPr`→`w:num`→`w:abstractNumId` | 编号 | 段落必须真引用 numId；只建 abstractNum 不算 |
| `w:pBdr` / `w:shd` | 段落边框底纹 | `w:shd@val` 是图案（clear/solid），`@fill` 才是颜色 |
| `w:framePr@dropCap/lines/h/wrap` | 首字下沉/文本框 | `dropCap="drop"` + `lines="2"`；文本框也用它定位 |
| `w:rFonts@ascii/hAnsi/eastAsia/cs` | 字体 | 中文看 `eastAsia`，西文 `ascii`；只判 ascii 会漏"华文行楷" |
| `w:b/@bCs` `w:i/@iCs` | 加粗/斜体 | `w:b w:val="0"` 是**取消**加粗，别只看标签存在 |
| `w:sz@val` `w:szCs` | 字号 | 半磅单位：二号=44（22 磅），小四=24；中文复杂文种看 szCs |
| `w:color@val/themeColor/themeTint` | 颜色 | 主题色要看 `themeColor` + `themeTint`，`val="auto"` 是自动 |
| `w:u@val` | 下划线 | 单线=`single`，波浪=`wave`，"始终无下划线"是链接样式不是 u |
| `w:vertAlign@val` | 上下标 | `superscript`/`subscript`，不是 position |
| `w:spacing@val`（rPr 内） | 字符间距加宽 | 单位 twips，加宽 2 磅 = `val="40"` |
| `w:em@val` / `w:highlight@val` | 着重号 / 突出显示 | 着重号是 em，不是下划线 |

### 表格

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `w:tblW@w/type` | 表宽 | `type="pct"` 是百分比（5000=100%），`dxa` 是 twips |
| `w:jc@val`（tblPr 内） | 表格对齐 | `center`；与内容居中无关 |
| `w:tblBorders` / `w:tcBorders` | 边框 | 单元格级覆盖表级；"外框双线内框单线"要分别判 top/insideH |
| `w:tblLayout@type` | 固定/自动布局 | "根据内容调整"=`autofit` |
| `w:tblCellMar` / `w:tblCellSpacing` | 单元格边距/间距 | 单位 twips |
| `w:tblStyle@val` + `w:tblLook` | 表格样式 | 样式 id 与显示名不同；套用后局部底纹会覆盖 |
| `w:trHeight@val/@hRule` | 行高 | 1 厘米=567；`hRule="exact"` 是固定值 |
| `w:gridSpan` / `w:vMerge` | 合并单元格 | 横向 gridSpan，纵向 vMerge(restart/continue)；没有 mergeCell 概念 |
| `w:tcW` / `w:shd` / `w:vAlign` | 列宽/底纹/垂直对齐 | `shd@fill` 颜色，`vAlign="center"` |
| `w:tblHeader` | 标题行重复 | 跨页表头 |

### 域、引用、批注

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `w:fldSimple@w:instr` / `w:instrText` | 域代码 | 页码 `PAGE`、总页 `NUMPAGES`、日期 `DATE`、目录 `TOC`、题注 `SEQ`、邮件合并 `MERGEFIELD`、交叉引用 `REF` |
| `w:fldChar@type` | 复合域分界 | begin/separate/end 三段，缓存结果在中间段 — 判"是不是域"看 instrText，判"显示什么"看缓存文本 |
| `w:bookmarkStart/@w:name` | 书签 | 交叉引用/目录跳转依赖它 |
| `w:hyperlink@r:id/@w:anchor` | 超链接 | 外链看 r:id→rels 的 Target；站内锚点看 `w:anchor` |
| `w:footnoteReference` + `footnotes.xml` | 脚注 | 编号在 `w:footnote@w:id`；尾注同理 endnotes |
| `w:commentReference` + `comments.xml` | 批注 | 批注正文独立部件 |
| `w:ins` / `w:del` / `w:rPrChange` | 修订 | 修订未接受时文字在 `w:ins` 里，`w:delText` 是被删内容 |
| `w:sdt` + `w:sdtPr` | 内容控件 | 下拉/日期控件：`w:dropDownList`、`w:date`；题干"控件"类考点 |

### 图形、图表、其他

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `w:drawing/wp:inline|wp:anchor` | 图片/图形容器 | inline=嵌入型；anchor 带 `behindDoc`、`wrap*`、`distT/B` |
| `wp:extent@cx/cy` | 尺寸 | EMU，1 厘米=360000；7 厘米=2520000 |
| `a:blip@r:embed` | 图片引用 | 指向 rels 里的 media 文件 |
| `a:srcRect@l/t/r/b` | 裁剪 | 千分比；`l="25000"` 裁掉左侧 25% |
| `a:lum@bright/contrast` | 亮度对比度 | 百分比 ×1000 |
| `a:prstGeom@prst` | 形状类型 | rect/roundRect/ellipse/rightArrow… |
| `a:prstTxWarp@prst` | 艺术字变形 | textArchUp 等 |
| `w:pict/v:shape/v:imagedata` | VML 旧式图形 | 剪贴画、水印、老模板；`v:textpath@string` 是艺术字文字 |
| `w:drawing` 里 `c:chart@r:id` | 图表 | 图表数据在 `word/charts/chartN.xml`，不在 document.xml |
| `dgm:relIds` | SmartArt | 内容在 `word/diagrams/dataN.xml` |
| `m:oMath` | 公式 | math 命名空间；老版公式是 `EQ` 域 |
| `w:settings` 里 `w:documentProtection` | 文档保护 | 只读/仅表单 |
| `w:mailMerge` | 邮件合并 | 见上文专章 |

