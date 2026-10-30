---
name: Excel-XML-评分器
description: 通过解析.xlsx文件的底层Open XML结构，对Excel 2010操作题进行精准评分。当用户需要自动评分或检查Excel文档格式时使用此技能（如单元格格式、公式函数、图表、条件格式、数据条、合并单元格、高级筛选、VLOOKUP等）。通过直接检查XML而非使用openpyxl等高层API，实现对标准答案文件100%检测通过率。
agent_created: true
---

# Excel XML 评分器

通过直接解析.xlsx文件内部的Open XML结构，精准检测Excel文档格式和内容，实现操作题自动评分。

## 权威参考（查结构先看这里）

- 规范：ECMA-376 / ISO-IEC 29500。Excel 部分看 Part 1 §18（SpreadsheetML），
  DrawingML（图表、形状、字体）看 §21。
- MS 文档：https://learn.microsoft.com/zh-cn/office/open-xml/spreadsheet/overview
  SDK 类名与元素一一对应（`CellFormat` = `xf`、`NumberingFormat` = `numFmt`）。
- ooxml.org：https://ooxml.org/ 只放了 Word（210）和 DrawingML（142）的元素页，
  SpreadsheetML 写着 Coming soon；Excel 的坑别在那儿找。
- 本地速查表：`student_exam_app/tools/reference/OOXML-速查.md`（条款号 + 易错点一页看完）。

## 概述

与Word评分器共享同一内核：ZIP解压 → XML解析 → 格式检查 → 评分。但Excel有其独特的数据结构。

**核心能力：** `sharedStrings`查文本 → `styles.xml`查格式 → `sheet.xml`查单元格 → `all_xml`兜底

## 统一内核（优先用这套，别再手写加载器）

项目里的 `student_exam_app/graders/_core/` 已经把三种文档的读取写完了，新评分器直接复用：

```python
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _core import ExcelBook, Report

book = ExcelBook(path)
sheet = book.sheet('员工工资表')          # 找不到全名时自动按包含匹配
cell = book.cell('A1', sheet)             # value / formula / style_index
fmt = book.styles.get(cell.style_index)   # 已解析到 font/fill/border/alignment/numfmt
fmt['font']['bold'], fmt['font']['name'], fmt['font']['size']
fmt['alignment'].get('horizontal'), fmt['alignment'].get('vertical')
fmt['numfmt']                             # '"¥"#,##0.0'
book.charts(); book.tables(); book.pivot_tables(); book.all_xml()
```

`ExcelBook` 内部已经处理掉：工作表按 `workbook.xml.rels` 映射（sheetId 陷阱不会再有）、单元格字典索引（`cell()` 是 O(1)，不用每次遍历全表）、样式索引解析到真实字体/填充/边框/对齐/数字格式。

新评分器建议用 `Report` 输出，后端和回归脚本都认这个格式：

```python
def grade(self):
    report = Report('Excel 第N套')
    report.add(1, '工作表+标题', 8, 10, ['改名 OK', 'A1 标题不符'])
    return report

from _core import run
if __name__ == '__main__':
    sys.exit(run(ExcelGraderN, title='Excel 第N套'))
```

`Report` 会输出 `总分: 88 / 100   得分率: 88%`，考试系统优先读"得分率"，读不到再按总分比例折算，所以满分不是 100 的评分器（Word 常见 110 分制）也能正确显示。

## 工作流程

```
素材.xlsx + 答案.xlsx
      │
      ▼  第一步：对比文件列表 + 单元格diff（⭐ 必须先做）
      │    → 找出新增的chart/table/公式/样式
      │    → 确认答案的精确实现方式（如AVERAGE vs SUMIF/COUNTIF）
      │
      ▼  第二步：加载XML → 字符串搜索兜底
      │    sharedStrings（文本索引）
      │    + styles.xml（字体/边框/填充/颜色/对齐）
      │    + sheet1.xml（单元格值、公式、合并）
      │    = self.all_xml（全量拼接）
      │
      ├──→ 产出A: 题目要求（自然语言）
      └──→ 产出B: 评分代码 → 测试 100/100
```

### ⚠️ 第一步必做：对比两个文件

**不要跳过！** Excel没有修订标记，需要自己写diff。

```python
# 1. 对比文件列表 → 找到新增文件（chart、table等）
# 2. 对比 sheet1 的单元格 → 找到所有新增/修改的cell
# 3. 对比 sharedStrings → 找到新增文本
```

## Excel vs Word 核心差异

| | Word | Excel |
|---|---|---|
| 文本读取 | `gt()` 合并 `w:t` | `sharedStrings[索引]` |
| XML解析 | `ET.fromstring(d.encode('utf-8'))` ✅ | `ET.fromstring(d)` ⚠️ 不能 re-encode！ |
| 修订标记 | ✅ `w:ins`/`w:del` | ❌ 需手动 diff |
| 格式检测 | 属性直接 `GA()` | **都在 styles.xml 里**，sheet 只存索引 |
| 边框 | `w:pBdr` / `w:tblBorders` | `styles.xml` → `border` 元素 |
| 字体/颜色 | pPr/rPr 内联 | `styles.xml` → `font` / `fill` 元素 |
| 对齐 | `w:jc` / `w:vAlign` | `styles.xml` → `alignment horizontal/vertical` |
| 标签颜色 | 不适用 | ⚠️ 在 **sheet1.xml** 不是 workbook.xml |
| 列宽 | 不适用 | `cols/col width customWidth` |
| 公式 | `w:instrText` | `c/f` 元素 |

## 考点覆盖表（按题库 10 套实测）

| 考点 | XML 证据 | 判定要点 | 已有实现 |
| --- | --- | --- | --- |
| 图表类型 | `chart1.xml` 的 `c:barChart`/`c:lineChart`/`c:pieChart`，三维看 `c:view3D` 或 `c:bar3DChart` | 图表 XML 不在 workbook 里，必须读 `xl/charts/chartN.xml`；`c:barDir@val="col"` 才是柱形 | excel1/2/6/8/9 |
| 图表标题/图例 | `c:title` 内有 `a:t`；`c:legendPos@val="r"` | 标题在 `c:title/c:tx/c:rich`；`legendPos` 缺省=右侧，判"靠右"要显式 r | excel1/6/9 |
| 图表位置 | `xdr:twoCellAnchor` 的 `xdr:from/xdr:to` | "在 A18:G30 生成"要用锚点的行列反推，不能只看 `drawing1.xml` 存在 | excel1 |
| 数据透视表 | `xl/pivotTables/pivotTable1.xml` + `pivotCacheDefinition` 的 `cacheField@name` | 字段名在 cacheField，不在 pivotTable；行/列区域看 `pivotField@axis="axisRow"` | excel5/6/9 |
| 条件格式 | `conditionalFormatting@sqref` + `cfRule@type` | 范围看 sqref，优先级看 priority；一个 sqref 可挂多条规则 | excel1/6/8 |
| 数据条 | `cfRule@type="dataBar"`；2010 的渐变条在 `extLst` 的 `x14:dataBar` | 基础 dataBar 只是兼容壳，判"渐变填充"要认 x14 或颜色 | excel1/6/8 |
| 迷你图 | `x14:sparklineGroups` 在 `extLst` 里 | 不在 `sheetData`；`x14:sparkline@f` 是数据范围，`sqref` 是放置位置 | excel5 |
| 筛选/高级筛选 | `autoFilter@ref` + `filterColumn`；高级筛选结果看行 `hidden="1"` + 顶部条件区 | 判"筛出了哪些行"最稳的是 hidden 属性；判"用了高级筛选"要 `_FilterDatabase` 定义名或 autoFilter 缺失 | excel1/2/3/4/6/8/10 |
| 排序 | `sortState@ref` + `sortCondition@descending` | sortState 是 worksheet 直接子节点；`descending` 缺省=升序；按两列排要看两个 sortCondition 顺序 | excel2/3/5/6/10 |
| 分类汇总 | 行上的 `outlineLevel` + `SUBTOTAL` 公式 | `outlinePr` 在 sheetPr 里是概要方向，`outlineLevel` 才在行上 | excel3/4/6/8/9 |
| 合并/跨列居中 | `mergeCell@ref` vs `alignment@horizontal="centerContinuous"` | 题干说"跨列居中"不能判 mergeCell，反过来也是 | excel2/5/6/8 |
| 数字格式 | `numFmt@formatCode` + `cellXfs@applyNumberFormat="1"` | 只写 numFmt 不 apply 不生效；货币 ¥ 是 `"¥"#,##0.0` | excel1/5 |
| 表样式 | `xl/tables/table1.xml` 的 `tableStyleInfo@name` | 名称是中文样式名（"表样式中等深浅 6"），不是 styleId | excel1/4/5/6/8 |
| 页面设置 | `pageSetup@paperSize/orientation` + `pageMargins` + `printOptions` | 纸张 A4 是 `paperSize="9"`；水平居中在 `printOptions@horizontalCentered` | excel3/4/9 |
| 页眉页脚 | `oddHeader/oddFooter` 里的 `&amp;L`、`&amp;P` | `&` 会转义成 `&amp;`，页码是 `&amp;P`，别按原文匹配 | excel3/4/7/9/10 |
| 冻结窗格 | `pane@xSplit/@ySplit/@topLeftCell` | 题库未考，新题按此写 | 待补 |
| 数据有效性 | `dataValidation@type/@sqref` + `formula1` | 题库未考，新题按此写 | 待补 |
| 打印标题 | `definedName@name="_xlnm.Print_Titles"` | 题库未考，新题按此写 | 待补 |

## 图表判定的四个坑

1. **chart XML 不在 workbook 里**：`xl/charts/chart1.xml` 是独立部件，遍历 `sheetData` 永远找不到。
2. **图表标题常是富文本**：`c:title/c:tx/c:rich/a:p/a:r/a:t`，直接取 `c:title` 的 text 会是空。
3. **图例位置缺省靠右**：`c:legendPos` 缺失时 Excel 默认右侧，判"右侧显示图例"不能要求节点存在。
4. **三维柱形是 `c:bar3DChart` + `c:view3D`**，不是 `c:barChart` 加个属性。

## 透视表判定的要点

- 表结构在 `pivotTable1.xml`：`pivotField@axis` 标识行/列/页字段，`dataField@name` 是值字段。
- 字段名在 `pivotCacheDefinition1.xml` 的 `cacheField@name`，中文名常在这里。
- 判"按某字段汇总"要同时看 axis 和 dataField，只看 pivotTable 存在等于没判。

## 筛选与排序的判定取向

- 优先判**结果**：筛选后不符合的行 `hidden="1"`，这比判"用了哪个菜单"更贴近题干要求，
  也兼容学生用高级筛选用自动筛选两种做法。
- 排序看 `sortState/sortCondition@descending`：`descending="1"` 是降序，`0` 或缺失是升序；
  两列排序时 `sortCondition` 的顺序就是主次关键字的顺序。

## 新补的陷阱（2026-10-30，对照 ECMA-376 校验后）

- **全文搜字体/字号是错的**：`'华文楷体' in all_xml` 只要工作簿里任何单元格用到就成立。
  一律用 `book.format_in(root, 'A1')` 按目标单元格取。已经因此修掉 excel1/2 的标题与数据区判断。
- **数据条要看 x14 扩展**：Excel 2010 的渐变/实心数据条写在 `extLst` 里的 `x14:dataBar`，
  基础 `<dataBar>` 只是兼容壳。判"渐变填充"时要认 `x14:dataBar` 的 `gradient` 属性或直接查颜色。
- **numFmt 要配套 `applyNumberFormat`**：只写 `numFmt` 而 cellXfs 里 `applyNumberFormat="0"`，
  单元格显示不会变，不能算数。
- **sortState 是 worksheet 的直接子节点**，不在 `sheetData`、也不在 `table` 里；
  `sortCondition@descending` 缺省等于升序，判降序必须显式 `descending="1"`。
- **conditionalFormatting 的判定范围看 `sqref`**，`cfRule@priority` 决定叠加顺序；
  多个规则时不要取第一个匹配就返回。
- **合并单元格与跨列居中是两回事**：`mergeCell` 合并；`alignment@horizontal="centerContinuous"`
  是不合并的跨列居中，题干说"跨列居中"要判后者。

## ⚠️ 致命陷阱（必读）

### 陷阱0：ET.fromstring 参数形式

### 陷阱0.5：sheetId ≠ XML文件编号！必须用 _rels 映射

**⚠️ 严重：** Excel中 sheetId（用户可见的序号）不一定等于 XML 文件名编号！

```python
# ❌ 错误：直接用 sheetId 拼 XML 路径
for name, sid in sheet_names:
    fname = f'xl/worksheets/sheet{int(sid)}.xml'  # sheetId=6 时会找不存在的 sheet6.xml!
    d = self.zf.read(fname).decode('utf-8')

# ✅ 正确：通过 workbook.xml.rels 中的 r:id → Target 映射
# workbook.xml: <sheet name="销售订单" sheetId="1" r:id="rId2"/>
# _rels/workbook.xml.rels: <Relationship Id="rId2" Target="worksheets/sheet2.xml"/>
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

# 1. 解析 rels 建立 rId→target 映射
rels_data = self.zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
rels_root = ET.fromstring(rels_data)
rel_map = {}
for rel in rels_root.iter(f'{{{PKG}}}Relationship'):
    r_id = rel.get('Id', '')
    target = rel.get('Target', '')
    if 'worksheet' in rel.get('Type', ''):
        # target: "worksheets/sheet1.xml"
        rel_map[r_id] = target

# 2. 通过 r:id 找到正确的 XML 文件
for sh in wb.iter(f'{{{S}}}sheet'):
    name = sh.get('name', '')
    r_id = sh.get(f'{{{R_NS}}}id', '')  # 获取 r:id
    target = rel_map.get(r_id, '')
    if target:
        fname = 'xl/' + target  # xl/worksheets/sheet1.xml
        d = self.zf.read(fname).decode('utf-8')
```

**什么情况下 sheetId ≠ XML 编号：**
- 新增数据透视表 → 可能分配新的 sheetId 但复用旧的 XML 编号
- 工作表重排序 → sheetId 不变但 XML 编号不变
- **始终使用 rels 映射，不要依赖 sheetId**

```python
# ❌ 错误：Excel XML re-encode 后命名空间全部丢失，所有元素查找返回 None
root = ET.fromstring(xml_str.encode('utf-8'))

# ✅ 正确：decode 后的字符串直接传入
root = ET.fromstring(xml_str)
```

### 陷阱1：样式全部在 styles.xml

**99% 的格式检测失败都是因为只搜了 sheet XML。**

```python
# ❌ 错误：只在 sheet1 搜边框
'border' in sheet1_xml  → False

# ✅ 正确：搜 all_xml（含 styles.xml）
'border' in self.all_xml  → True
```

`s3_xml` 改 `self.all_xml` 是本次 R10 从 7→10 的关键。

### 陷阱2：tabColor 在 sheet1.xml

```python
# ❌ 错误：在 workbook.xml 的 sheet 元素下找
sheet.find(f'{{{S}}}tabColor')  → None

# ✅ 正确：直接搜 all_xml
'tabColor' in self.all_xml  → True
```

### 陷阱3：函数名不唯一

答案可能用 `SUMIF+COUNTIF` 代替 `AVERAGE`，用 `SUM` 代替 `+`。检测时用字符串包含判断而非精确函数名。

### 陷阱4：字体大小用 `<sz val="16"/>`

`sz="16"` 作为属性名不存在于Excel中。Excel将字号存在 `<sz val="16"/>` 子元素里。字符串搜索时注意格式差异。

### 陷阱5：跨列居中 ≠ 合并居中

**这是两个完全不同的概念！**

```python
# ❌ 错误：跨列居中搜 mergeCell
if 'mergeCell' in s1_xml and 'A1:L1' in s1_xml.replace(' ', ''):
    → 题目要求"跨列居中"但答案可能根本没有 mergeCell！

# ✅ 正确：跨列居中 = styles.xml → alignment horizontal="centerContinuous"
if 'centerContinuous' in self.all_xml:
    score += 2  # 这才是跨列居中
```

**辨别规则：**
- "合并居中" → `mergeCell` 存在于 sheet1.xml
- "跨列居中" → `alignment horizontal="centerContinuous"` 存在于 styles.xml
- 两者互斥，不可混淆检测

### 陷阱6：>= 在XML中是 `&gt;=`

Excel XML中 `>` 被编码为 `&gt;`。

```python
# ❌ 错误：直接搜 ">=95"
'>95' in self.all_xml  → False (即使单元格显示 >=95)

# ✅ 正确：搜 "&gt;=95" 或直接 get_cell value
'&gt;=95' in self.all_xml  → True
# 或直接读取单元格值（已解析的 sharedStrings）
self.sv(root_filter, 'A2')  → '>=95'
```

### 陷阱8：& 在页眉/页脚中编码为 `&amp;`

```python
# ❌ 错误：直接搜 &10 或 &R&P
'&10' in s2_xml → False
'&R&P' in s2_xml → False

# ✅ 正确：& 在 XML 中为 &amp;
'&amp;10' in s2_xml  → True   # 字号10
'&amp;R&amp;P' in s2_xml → True  # 靠右页码
```

页眉/页脚特殊编码：
- `&C` = 居中, `&R` = 靠右, `&P` = 页码, `&10` = 10号字
- XML中全部编码为 `&amp;C`, `&amp;R`, `&amp;P`, `&amp;10`

### 陷阱9：分类汇总的 outlineLevel 在行上，不是 outlinePr

```python
# ❌ 错误：搜 outlinePr 元素
'outlinePr' in s4_xml  → 可能为 False（元素可能不存在）

# ✅ 正确：检查行的 outlineLevel 属性
has_outline = any(row.get('outlineLevel', '') != ''
                  for row in root.iter(f'{{{S}}}row'))
# outlineLevel=1 → 汇总行, outlineLevel=2 → 明细行
```

### 陷阱10：条件格式可能未实现——但代码逻辑不应因此退化

**核心教训（来自文件夹3实战）：** 评分器代码逻辑正确，问题在答案文件本身。

```python
# 我们的代码在 all_xml 中搜索 conditionalFormatting——逻辑完全正确
has_cf = 'conditionalFormatting' in self.all_xml
# 第一次测试：97分（R5条件格式3分未得分）
# 原因：答案文件缺少 conditionalFormatting → sheet1.xml 中无此元素

# 用户修改答案文件后：100分
# 答案文件新增: fill[2] bgColor=FFFF0000 (红色底纹)
#               dxf[0] font color=FF002060 (深蓝字体) + fill bgColor=FFFF0000
#               cfRule type=cellIs operator=greaterThanOrEqual formula=1000 sqref=E2:E44
```

**经验法则：**
1. **信任XML证据** — 全量搜索zip中所有文件，找不到就是答案没做
2. **不要猜答案** — 不要因为"题干要求"就给分，以XML实际内容为准
3. **代码不用改** — 答案没做导致未得分，是答案的问题，不是评分器的问题
4. **对照题干** — 评分器输出结果要能解释：哪些得分（有XML证据）、哪些失分（缺XML证据）

### 陷阱11：`'<b' in all_xml` 会把边框当成加粗

`<border>`、`<bottom>`、`<bgColor>`、`<brk>` 全部以 `<b` 开头。用子串找加粗，等于"只要表里有边框就算加粗"。

```python
# ❌ 危险：表里有边框就为真
if '<b' in self.all_xml: score += 1.5

# ✅ 方式一：查单元格自己的样式
if self.book.format_in(root, 'A1').get('font', {}).get('bold'): score += 1.5

# ✅ 方式二：字符串兜底时限定标签边界
if re.search(r'<b(?:\s[^>]*)?/?>', self.all_xml): score += 1.5
```

实战案例：第 3 套标准答案 A1 只设了微软雅黑 16 号红色，**没有加粗**，题干却要求加粗。旧代码靠这条误判每份文件都白送 1 分，掩盖了答案缺陷。改成样式判断后立刻暴露，补上答案的 `<b/>` 才回到 100。

### 陷阱12：`'val="16"' 不等于字号 16

`val` 是所有元素通用的属性名，`<numFmt val=...>`、`<family val=...>`、`<charset val=...>` 都可能命中。字号必须限定在 `<sz>` 上。

```python
# ❌ 'val="16"' in all_xml
# ✅ '<sz val="16"' in all_xml
```

同理："垂直居中"不要写 `'vertical' in all_xml`（任何 alignment 的 vertical 都会命中），要写 `'vertical="center"'`，或者直接读 `fmt['alignment']['vertical']`。

### 陷阱13：字符串匹配会把整张表的格式算成目标单元格的

`all_xml` 是全表拼接，命中即给分等于"只要范围内任一处有该格式就算完成"。题目要求"标题 A1 加粗"时，数据区某列加粗也会让检查通过。有明确单元格的要求，一律走 `format_in(root, ref)`；只有"整个区域统一格式"这类要求才用全局搜索。

## Excel XML 结构

```
文件名.xlsx/
├── xl/
│   ├── workbook.xml          ← 工作表名称
│   ├── sharedStrings.xml     ← ⭐ 所有文本的索引表
│   ├── styles.xml            ← ⭐ 字体/边框/填充/对齐/数字格式
│   ├── calcChain.xml         ← 公式计算链
│   ├── theme/theme1.xml      ← 主题颜色
│   ├── worksheets/
│   │   ├── sheet1.xml        ← 主工作表（单元格、合并、条件格式）
│   │   ├── sheet2.xml        ← 辅助表（如VLOOKUP查找表）
│   │   └── sheet3.xml        ← 高级筛选目标表
│   ├── charts/chart1.xml     ← 图表定义
│   ├── tables/table1.xml     ← 套用表样式
│   ├── drawings/             ← 图表嵌入关系
│   └── _rels/                ← 关系文件（图片、图表引用）
```

## 核心代码模板

```python
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

class ExcelGrader:
    def __init__(self, path):
        self.ss = []          # sharedStrings
        self.sheets = {}      # {name: {'root': Element, 'xml': str}}
        self.all_xml = ''     # ⭐ 全量拼接 — 格式检测首选
        self.styles_xml = ''
        self.namelist = []

    def load(self):
        self.zf = zipfile.ZipFile(self.path)
        self.namelist = self.zf.namelist()
        all_parts = []

        # ⚠️ ET.fromstring(xml_str) — 不能 encode('utf-8')！
        # 1. SharedStrings
        d = self.zf.read('xl/sharedStrings.xml').decode('utf-8')
        all_parts.append(d)
        root = ET.fromstring(d)
        for si in root.iter(f'{{{S}}}si'):
            t = si.find(f'{{{S}}}t')
            if t is not None:
                self.ss.append(t.text if t.text else '')
            else:
                self.ss.append(''.join(tt.text or '' for tt in si.iter(f'{{{S}}}t')))

        # 2. Workbook → sheet names
        d = self.zf.read('xl/workbook.xml').decode('utf-8')
        all_parts.append(d)
        for sheet in ET.fromstring(d).iter(f'{{{S}}}sheet'):
            name = sheet.get('name', '')
            self.sheets[name] = {'wb_el': sheet}

        # 3. Each sheet
        for sid in range(1, 10):
            fname = f'xl/worksheets/sheet{sid}.xml'
            if fname in self.namelist:
                d = self.zf.read(fname).decode('utf-8')
                all_parts.append(d)
                # Find name from workbook info
                for name, info in self.sheets.items():
                    if 'root' not in info:
                        info['root'] = ET.fromstring(d)
                        info['xml'] = d
                        break

        # 4. Styles
        if 'xl/styles.xml' in self.namelist:
            d = self.zf.read('xl/styles.xml').decode('utf-8')
            all_parts.append(d); self.styles_xml = d

        self.all_xml = '\n'.join(all_parts)

    def sv(self, sheet_root, ref):
        """Get cell value — handles shared strings"""
        for row in sheet_root.iter(f'{{{S}}}row'):
            for c in row.iter(f'{{{S}}}c'):
                if c.get('r', '') == ref:
                    v = c.find(f'{{{S}}}v')
                    if v is None or v.text is None: return ''
                    if c.get('t', '') == 's':
                        idx = int(v.text)
                        return self.ss[idx] if idx < len(self.ss) else ''
                    return v.text
        return ''

    def get_cell(self, sheet_root, ref):
        """Get raw cell Element — for formula access"""
        for row in sheet_root.iter(f'{{{S}}}row'):
            for c in row.iter(f'{{{S}}}c'):
                if c.get('r', '') == ref: return c
        return None

    def get_formula(self, sheet_root, ref):
        """Get cell formula — short for get_cell + find f"""
        c = self.get_cell(sheet_root, ref)
        if c is not None:
            f = c.find(f'{{{S}}}f')
            if f is not None and f.text:
                return f.text
        return ''
```

## 常见考题检测模式

### 1. 工作表改名 + 标签颜色
```python
# 改名：workbook.xml 的 sheet name 属性
# 标签颜色：⚠️ 在 sheet1.xml，用 all_xml 搜
if '员工工资表' in self.sheets: score += 2
if 'tabColor' in self.all_xml: score += 2
```

### 2. 跨列居中 + 合并居中 + 字体格式

```python
# 跨列居中：⚠️ centerContinuous，不是 mergeCell！
# 跨列居中是 alignment 属性，在 styles.xml 中
if 'centerContinuous' in self.all_xml: score += 2  # 跨列居中

# 合并居中：sheet1.xml → mergeCells
if 'mergeCell' in s1_xml and 'A1:L1' in s1_xml.replace(' ', ''): score += 2

# ⚠️ 字体/字号/颜色：都在 styles.xml，用 all_xml 搜
if '华文楷体' in self.all_xml: score += 1.5
if '楷体' in self.all_xml.replace('华文楷体', ''): score += 1.5
if '<b' in self.all_xml: score += 1.5
if '7030A0' in self.all_xml.upper(): score += 1.5  # 紫色
```

### 3. 行高 + 列宽 + 垂直居中
```python
# 行高：row 的 ht 属性
for row in root.iter(f'{{{S}}}row'):
    if row.get('r') == '1' and abs(float(row.get('ht', '0')) - 32) < 2:
        score += 1.5; break

# 列宽：cols/col 的 customWidth 属性
if 'customWidth' in s1_xml: score += 1

# 垂直居中：⚠️ 在 styles.xml alignment 元素中
if 'vertical' in self.all_xml.lower(): score += 0.5
```

### 4. 对齐 + 货币格式 + 底纹填充

```python
# 对齐：在 styles.xml 的 alignment 元素
if 'horizontal="center"' in self.all_xml: score += 3
if 'horizontal="right"' in self.all_xml: score += 2

# ￥：在 styles.xml 的 numFmt
if '¥' in self.styles_xml: score += 3
if '0.0' in self.styles_xml: score += 2  # 1位小数

# 底纹/填充：⚠️ patternFill solid + fgColor theme/tint
# "水绿色，强调文字颜色5，淡色80%" → theme=8, tint=0.8
if 'patternFill' in self.styles_xml and 'solid' in self.styles_xml:
    if 'theme="8"' in self.styles_xml:      # 强调文字颜色5
        score += 5
    if 'tint="0.7999' in self.styles_xml or '0.8' in self.styles_xml:  # 淡色80%
        score += 5
# 其他主题色映射：theme=5(强调1/蓝色), theme=6(强调2/红色),
# theme=7(强调3/绿色), theme=8(强调5/水绿色), theme=9(强调6/橙色)
```

### 5. 公式/函数
```python
# 公式在 c/f 元素中
for c in sheet_root.iter(f'{{{S}}}c'):
    f = c.find(f'{{{S}}}f')
    if f is not None and f.text:
        # VLOOKUP: "F3*VLOOKUP(D3,加班费!$A$2:$B$5,2,FALSE)"
        # SUM:     "SUM(E3,G3)"
        # IF:      "IF(F3>=2.5,\"优秀\",\"加油干\")"
        # RANK:    "RANK(J3,$J$3:$J$31)"
        # SUMIF:   "SUMIF($B$2:$B$16,\"女\",$C$2:$C$16)/COUNTIF(...)"
```

### 6. 平均年龄（SUMIF/COUNTIF 替代 AVERAGE）
```python
# ⚠️ 答案可能用 SUMIF/COUNTIF，不是 AVERAGE/AVERAGEIF
has_avg = ('SUMIF' in s1_xml.upper() and 'COUNTIF' in s1_xml.upper())
if has_avg: score += 6
```

### 7. 套用表样式 + 条件格式 + 数据条
```python
# 表样式：xl/tables/table1.xml 存在
if 'xl/tables/table1.xml' in self.namelist: score += 3

# 条件格式：conditionalFormatting 元素
if 'conditionalFormatting' in self.all_xml: score += 2
if '1250' in self.all_xml: score += 3  # 条件值

# 数据条：extLst 中的 dataBar
if 'dataBar' in self.all_xml: score += 2
```

### 8. 图表

```python
# ⚠️ 图表内容在 chart1.xml，不在 all_xml 中！
if 'xl/charts/chart1.xml' not in self.namelist:
    return # 没有图表
chart = self.zf.read('xl/charts/chart1.xml').decode('utf-8')

# 图表类型
if 'bar3DChart' in chart: score += 3     # 三维簇状柱形图
elif 'barChart' in chart: score += 2     # 簇状柱形图(非三维)

# 标题 — 搜 chart XML
if '学生平均分比较图' in chart: score += 3    # 标题文本

# 图例 — legendPos val="r" = 靠右
if 'legendPos' in chart and 'val="r"' in chart: score += 2  # 图例靠右
elif 'legend' in chart.lower(): score += 1                    # 图例存在(位置未确认)

# 折线图检测：lineChart + marker
if 'lineChart' in chart:
    score += 2   # 折线图
    if 'marker' in chart.lower(): score += 1  # 带数据标记

# 图表标题字体：⚠️ 在 chart1.xml 的 rich text 中
# "楷体、加粗、字号12" → 字号按100倍存储: 12pt = 1200
if '楷体' in chart: score += 1   # 标题字体
if 'b=' in chart or '<a:b' in chart: score += 1  # 加粗
if '1200' in chart: score += 1   # 12号 (12×100)
```

### 9. 排序

```python
# ⚠️ 排序状态在 sheet1.xml → sortState 元素
# 每个 sortCondition 有 ref(列范围) 和 descending(1=降序)

for sortState in sheet_root.iter(f'{{{S}}}sortState'):
    ref = sortState.get('ref', '')          # 数据范围 如 A3:L15
    conditions = list(sortState.iter(f'{{{S}}}sortCondition'))
    
    # 主关键字: 班级(C列) 升序
    sc1 = conditions[0]
    if 'C' in sc1.get('ref', '') and sc1.get('descending') != '1':
        score += 4
    
    # 次关键字: 总分(K列) 降序
    sc2 = conditions[1]
    if 'K' in sc2.get('ref', '') and sc2.get('descending') == '1':
        score += 4

# 兜底：如无sortState，验证数据是否已排序
if not sort_found:
    classes = [sv(f'C{r}') for r in range(start, end+1)]
    if classes == sorted(classes):
        score += 5  # 数据已排序(但未设置sortState)
```

### 10. 高级筛选 + 边框

### 11. 边框颜色参考

```python
# Excel XML 边框颜色使用 ARGB 格式：AARRGGBB
# 常见颜色：
BORDER_COLORS = {
    '蓝色': '0070C0',
    '橙色': 'FFC000',
    '紫色': '7030A0',  # ⚠️ Excel紫色不是800080
    '红色': 'FF0000',
    '黑色': '000000',
}

# 边框样式：thin, medium, thick, double, dotted, dashed, hair
# 检测外边框：thick + 颜色
if 'thick' in self.all_xml and '0070C0' in self.all_xml.upper():
    score += 5  # 蓝色粗线外边框
# 检测内边框：double + 颜色
if 'double' in self.all_xml and 'FFC000' in self.all_xml.upper():
    score += 5  # 橙色双线内边框
```

### 12. 页面设置（页边距/页眉页脚/横向/起始页码）

```python
# ⚠️ 页面设置都在 sheet2.xml（或对应sheet的XML）中
s2_xml = info['xml']

# 页边距水平/垂直居中 → printOptions 元素
if 'horizontalCentered="1"' in s2_xml: score += 2
if 'verticalCentered="1"' in s2_xml: score += 2

# 纸张横向 → pageSetup orientation
if 'orientation="landscape"' in s2_xml: score += 2

# 起始页码 → pageSetup firstPageNumber
if 'firstPageNumber="5"' in s2_xml: score += 1

# 页眉 → headerFooter → oddHeader (⚠️ & 编码为 &amp;)
# "&C&10销售情况表" → "&amp;C&amp;10销售情况表"
if '销售情况表' in s2_xml: score += 2  # 字符串搜足矣
# 页脚靠右页码 → oddFooter
# "&R&P" → "&amp;R&amp;P"
if '&amp;R&amp;P' in s2_xml: score += 2
# 页眉/页脚格式码：&C=居中 &R=靠右 &P=页码 &N=总页数 &D=日期 &10=10号字

# 自动列宽（单独的sheet）
if 'bestFit' in s2_xml and 'customWidth' in s2_xml: score += 1
```

### 13. 分类汇总（排序 + outlineLevel + SUBTOTAL）

```python
# ⚠️ 分类汇总检测三要素：排序 + 分级 + SUBTOTAL函数
s4_root = sheets.get('分类汇总', {}).get('root')

# 1. 排序 → sortState（日期降序A列 + 品名升序B列）
for ss in s4_root.iter(f'{{{S}}}sortState'):
    for sc in ss.iter(f'{{{S}}}sortCondition'):
        ref = sc.get('ref', ''); desc = sc.get('descending', '')
        if 'A' in ref and desc == '1': score += 2  # 日期降序
        if 'B' in ref and desc != '1': score += 2  # 品名升序

# 2. 分级显示 → 行的 outlineLevel 属性（不是 outlinePr！）
has_outline = any(row.get('outlineLevel', '') != ''
                  for row in s4_root.iter(f'{{{S}}}row'))
if has_outline: score += 2
# outlineLevel=1 → 汇总行（显示最大值/求和等）
# outlineLevel=2 → 明细行（原始数据）

# 3. SUBTOTAL函数 → 汇总行公式
if 'SUBTOTAL' in s4_root: score += 2
```

### 14. 套用表样式

```python
# 表样式在 xl/tables/table1.xml → tableStyleInfo
if 'xl/tables/table1.xml' in self.namelist:
    tbl = self.zf.read('xl/tables/table1.xml').decode('utf-8')
    # 中等深浅: TableStyleMedium, 浅色: TableStyleLight, 深色: TableStyleDark
    if 'TableStyleLight' in tbl: score += 4  # 浅色样式(如浅色11)
    elif 'TableStyleMedium' in tbl: score += 4  # 中等深浅
    # tableStyleInfo name="TableStyleLight11" → 浅色11
```

### 15. 数据透视表

```python
# ⚠️ 透视表定义在 xl/pivotTables/pivotTable1.xml
# 但字段名称在 xl/pivotCache/pivotCacheDefinition1.xml 的 cacheField 中！

if 'xl/pivotTables/pivotTable1.xml' in self.namelist:
    pt = self.zf.read('xl/pivotTables/pivotTable1.xml').decode('utf-8')
    pc = self.zf.read('xl/pivotCache/pivotCacheDefinition1.xml').decode('utf-8')
    full = pt + pc  # ⚠️ 合并搜索
    
    # 列标签=图书名称, 行标签=书店名称
    if '图书名称' in full: score += 4   # 列标签字段名
    if '书店名称' in full: score += 4   # 行标签字段名
    
    # 数据项=求和项:销量（本）
    if '求和项:销量' in full: score += 4

# ⚠️ 关键：pivotTable.xml 只用 field index (x="2") 引用字段
# 实际名称在 pivotCache → cacheField name="书店名称"
```

### 16. 迷你图 (Sparkline)

```python
# ⚠️ 迷你图在 sheet XML 中，搜索 sparkline 或 sparklineGroup
s7_xml = self.sheets.get('分析表', {}).get('xml', '')
if 'sparkline' in s7_xml.lower() or 'sparklineGroup' in s7_xml:
    score += 3  # 迷你图存在

# 高低点标记
if 'high' in s7_xml.lower() or 'markers' in s7_xml.lower():
    score += 2  # 标记了最高/最低点
```

### 17. VLOOKUP / SUMIF / SUMIFS

```python
# VLOOKUP跨表引用: VLOOKUP(E3,图书编目表!$A$1:$C$9,2)
# SUMIF条件求和: SUMIF(C2:C58,"新华书店",G2:G58)
# SUMIFS多条件: SUMIFS(源!$G$2:$G$58,源!$B$2:$B$58,">2021/5/31",源!$B$2:$B$58,"<2021/9/1")
for c in root.iter(f'{{{S}}}c'):
    f = c.find(f'{{{S}}}f')
    if f is not None and f.text:
        if 'VLOOKUP' in f.text.upper(): ...  # VLOOKUP函数
        if 'SUMIFS' in f.text.upper(): ...   # SUMIFS函数
        if 'SUMIF' in f.text.upper(): ...    # SUMIF函数
```

```python
s3 = self.sheets.get('筛选', {}).get('root')
if s3 is None: return  # 未找到筛选表
rows = list(s3.iter(f'{{{S}}}row'))
if len(rows) >= 5: score += 2  # 数据行充足

# 条件区域检测（用sv读取已解析的单元格值）
a1 = self.sv(s3, 'A1')  # 如 "高等代数"
b1 = self.sv(s3, 'B1')  # 如 "大学英语"
a2 = self.sv(s3, 'A2')  # 如 ">=95" (sv已解析>gt;)
b2 = self.sv(s3, 'B2')  # 如 ">=95"
if '高等代数' in a1: score += 1
if '大学英语' in b1: score += 1
if '95' in a2: score += 1   # sv 已解析 &gt;= 为 >=
if '95' in b2: score += 1

# 粘贴值 = 无公式（重要！）
has_f = any(f is not None and f.text for f in s3.iter(f'{{{S}}}f'))
if not has_f: score += 2  # 无公式 = 粘贴值

# 高级筛选 = 非匹配行隐藏 hidden="1"
hidden_count = 0
for row in s3.iter(f'{{{S}}}row'):
    if row.get('hidden') == '1':
        hidden_count += 1
if hidden_count > 0:
    score += 2   # 筛选已应用(XX行隐藏)
# ⚠️ 注意：hidden 行仍存在于 XML 中，只是 display 被隐藏

# ⚠️ 边框在 styles.xml，用 all_xml 搜
if 'double' in self.all_xml.lower() and 'border' in self.all_xml.lower():
    score += 3
```

## 一次性成功检查清单

### diff 阶段（必做）
- [ ] 对比文件列表：答案新增了哪些文件（chart、table、calcChain、drawings）？
- [ ] 对比单元格：哪些 cell 是答案独有的？哪些值不同？
- [ ] 确认实现方式：答案用 SUMIF 还是 AVERAGE？用 SUM 还是 +？
- [ ] 检查 sharedStrings：新增的文本（标题、列名、条件值）
- [ ] 检查 sheet 改名：workbook.xml → sheet name 属性
- [ ] 检查 styles：新增的字体/边框/填充/数字格式/对齐

### 加载阶段
- [ ] `ET.fromstring(xml_str)` — **不能 re-encode**
- [ ] `sharedStrings` 索引顺序不变
- [ ] 所有 XML 拼入 `all_xml`（含 styles、sheet1/2/3）

### 检测阶段
- [ ] 字体/字号/颜色/对齐/边框 → **搜 `all_xml`**（不是 sheet xml）
- [ ] 标签颜色 → 搜 `all_xml`（在 sheet1.xml 不是 workbook.xml）
- [ ] 跨列居中 → `centerContinuous`（不是 mergeCell）
- [ ] 列宽 → `customWidth` + `bestFit`（自动调整）
- [ ] 紫色 → `7030A0` / 蓝色 → `0070C0` / 橙色 → `FFC000` / 红色 → `FF0000` / 深蓝字 → `002060`
- [ ] 图表字体 → chart1.xml 标题 rich text: 字体名 + `<a:b/>`(加粗) + `sz="1200"`(12×100)
- [ ] 折线图 → `lineChart` + `marker`(数据标记)
- [ ] hidden行计数 → **用显式 `for` 循环**，不用 `sum(generator)`
- [ ] `>=` → XML中编码为 `&gt;=`，用 `sv()` 读取已解析值
- [ ] `&` → 页眉/页脚中编码为 `&amp;`（如 `&amp;10`, `&amp;R&amp;P`）
- [ ] 图表 → 单独读 `chart1.xml`（不在 all_xml 中）
- [ ] 图例位置 → `legendPos val="r"` 靠右
- [ ] 排序 → `sortState` > `sortCondition` 的 `ref` 和 `descending`
- [ ] 高级筛选 → `hidden="1"` 行数
- [ ] 底纹 → `patternFill solid` + `theme` + `tint`
- [ ] 函数 → 不硬编码函数名，用字符串包含判断
- [ ] 加粗 → 查单元格字体 `bold`，不要用 `'<b' in all_xml`（会命中 `<border>`）
- [ ] 字号 → `<sz val="N"`，不要用 `val="N"`
- [ ] 垂直居中 → `vertical="center"`，不要用 `'vertical' in all_xml`
- [ ] 明确单元格的格式要求 → `format_in(root, ref)`，不用全局搜索
- [ ] 合并 → `A1:G1` replace 空格后匹配
- [ ] 粘贴值 → 该 sheet 无 `c/f` 元素
- [ ] 页面设置 → `horizontalCentered` / `verticalCentered` / `firstPageNumber` / `orientation`
- [ ] 分类汇总 → `outlineLevel` 在行上 + `SUBTOTAL` 函数
- [ ] 表样式 → `xl/tables/table1.xml` 中 `tableStyleInfo` name
- [ ] 条件格式可能未实现 → 以答案文件为准，不要假设一定存在

## 示例用法

**用户说：** "给这个Excel操作题做评分器。答案文件是1结果.xlsx。"

**操作步骤：**
1. 对比两个xlsx → 列表差异 + 单元格diff → 理解每道题的实现
2. 加载XML（sharedStrings + styles + sheets）→ `all_xml`
3. 按检测模式写代码（字体搜 all_xml，公式查 c/f，边框搜 all_xml）
4. 测试 → 100/100

## 评分器方法论总结

### 核心原则：XML ∝ 得分

```
题干要求 → 对应XML证据 → 得分
          ❌ 无XML证据 → 不得分（代码逻辑不变）
```

### 实战验证（文件夹1-3）

| 文件夹 | 首次得分 | 原因 | 解决 |
|--------|---------|------|------|
| 1 | 100 | 答案完整 | — |
| 2 | 100 | 答案完整 | — |
| 3 | 97→100 | 条件格式缺失 | 用户补全答案 |
| 4 | 100 | 答案完整(修正hidden循环) | 显式for替代sum(generator) |

**关键认知：** 4个文件夹中3个一次通过，1个因答案文件缺失功能。代码逻辑本身无缺陷。hidden行计数使用显式 `for` 循环比 `sum(generator)` 更可靠。

### 开发流程最佳实践

1. **第一步：全量diff** → 对比素材和答案的文件列表、sharedStrings、单元格值、styles
2. **第二步：字符串验证** → 对每个题干要求，在all_xml中确认对应XML片段存在
3. **第三步：写检测代码** → 每个检测项标注数据源（all_xml / styles_xml / chart1.xml）
4. **第四步：测试** → 对答案跑100分，对素材跑低分（验证区分度）
5. **第五步：复查** → 如有失分，全量搜索确认是"答案缺失"还是"代码漏检"

---

## 回归与部署

改完任何评分器，先在 `student_exam_app/` 下跑：

```bash
python tools/selftest.py     # 全部评分器对标准答案跑一遍，必须全 100
python tools/core_sweep.py   # 核心库能否无异常解析 30 份答案文件
python deploy.py --dry-run   # 发布前预检 + 打包
```

`tools/selftest.py` 会自动找 `<科目>_grader_<题号>.py` 和对应文件夹里的 `1结果.*`，输出得分表；不满分会打印 stderr 摘要。
判分服务在 `student_exam_app/app.py`，`deploy.py` 负责打包、备份、上传、重启、探活。

---

**技能创建：** 2026-07-03  
**版本：** 1.9（统一内核 + 误判修正 + 回归部署）  
**更新内容：**
- 1.9：新增 `_core` 统一内核用法；陷阱11-13（`<b` 命中 `<border`、`val=` 模糊、全局搜索串格）；`Report` 输出契约；回归与部署命令
- 1.8：新增 `scripts/excel_analyze.py`（素材vs答案差分器）+ `scripts/excel_grader_template.py`（评分器模板） 
- 1.0：初始化
- 1.1：致命陷阱4条 + 第一步diff
- 1.2：跨列居中、排序、高级筛选、图表图例、底纹、边框颜色
- 1.3：&amp;编码、outlineLevel、页面设置、分类汇总、表样式
- 1.4：方法论总结、陷阱10重写（代码逻辑正确、答案需补全）
- 1.5：RANK函数、lineChart+marker、图表标题字体、hidden行显式for循环
- 1.6：**⚠️重大修复：sheetId≠XML编号(rels映射)**、数据透视表(pivotTable+pivotCache)、VLOOKUP/SUMIF/SUMIFS、sparkline迷你图、表样式浅色11
- 1.7：**新增精确XML属性参考表**（字体/填充/边框/对齐/数字格式）- 基于ISO/IEC 29500:2016
- 1.8：**架构级参考** — CT_Cell/CT_Row类型定义、度量单位矩阵、容差配置、Transitional/Strict说明

---

## 精确 XML 属性参考（基于 ISO/IEC 29500:2016）

### 颜色值格式
| 格式 | 示例 | 说明 |
|------|------|------|
| ARGB 六字符 | `"FF0000"` | 常用于 border color、font color |
| ARGB 八字符 | `"FF7030A0"` | 带 alpha 通道 |
| 主题色 + tint | `theme="8" tint="0.79998"` | Excel 底纹常用 |
| 索引色 | `indexed="64"` | 系统调色板 |

### 字体 (Font) 属性
| XML | 属性 | 说明 | 常见值 |
|-----|------|------|--------|
| `<b/>` | val | 加粗 | `"true"` (可为自闭合) |
| `<i/>` | val | 斜体 | `"true"` |
| `<u/>` | val | 下划线 | `"single"`, `"double"`, `"none"` |
| `<strike/>` | val | 删除线 | `"true"` |
| `<sz/>` | val | 字号(磅) | `"11"`, `"14"`, `"16"` |
| `<name/>` | val | 字体名 | `"宋体"`, `"华文楷体"` |
| `<family/>` | val | 字体族 | `2`=Swiss, `3`=Modern |
| `<charset/>` | val | 字符集 | `134`=GB2312中文 |
| `<scheme/>` | val | 主题方案 | `"major"`, `"minor"` |
| `<color/>` | rgb | RGB颜色 | `"FF0000"` |
| `<vertAlign/>` | val | 上下标 | `"superscript"`, `"subscript"` |

### 填充 (Fill) 属性
| XML | 属性 | 说明 | 常见值 |
|-----|------|------|--------|
| `<patternFill/>` | patternType | 图案类型 | `"solid"`, `"none"`, `"gray125"`, `"lightGray"` |
| `<fgColor/>` | rgb/theme/tint | 前景色 | `theme="8" tint="0.79998"` |
| `<bgColor/>` | rgb/indexed | 背景色 | `indexed="64"` |
| `<gradientFill/>` | type | 渐变类型 | `"linear"`, `"path"` |
| `<gradientFill/>` | degree | 渐变角度 | `0` |

### 边框 (Border) 属性
| XML | 属性 | 说明 | 常见值 |
|-----|------|------|--------|
| `<left/>` `<right/>` `<top/>` `<bottom/>` | style | 线条样式 | `"thin"`, `"medium"`, `"thick"`, `"double"`, `"dotted"`, `"dashed"`, `"hair"`, `"mediumDashed"`, `"dashDot"`, `"slantDashDot"` |
| | color | 线条颜色 | `"0070C0"`(蓝), `"FFC000"`(橙), `"FF0000"`(红) |
| `<diagonal/>` | style | 对角线样式 | 同上 |

### 对齐 (Alignment) 属性
| 属性 | 值 | 说明 |
|------|-----|------|
| `horizontal` | `"general"`, `"left"`, `"center"`, `"right"`, `"fill"`, `"justify"`, `"centerContinuous"`, `"distributed"` | ⚠️ centerContinuous=跨列居中 |
| `vertical` | `"top"`, `"center"`, `"bottom"`, `"justify"`, `"distributed"` | 垂直对齐 |
| `textRotation` | 0-180, 255 | 255=竖排 |
| `wrapText` | `"1"`, `"true"` | 自动换行 |
| `shrinkToFit` | `"1"`, `"true"` | 缩小填充 |

### 数字格式 (NumberFormat) 常用 ID
| numFmtId | formatCode | 示例 |
|----------|------------|------|
| 0 | — | 常规 |
| 2 | `"0.00"` | 两位小数 |
| 10 | `"0%"` | 百分比 |
| 14 | `"m/d/yyyy"` | 日期 |
| 164+ | 自定义 | `"¥#,##0.0"` |

### 单元格格式索引 (xf) 结构
```
styles.xml → cellXfs (count=N) → xf[N]:
  numFmtId="166" fontId="2" fillId="3" borderId="4" xfId="0"
  applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"
```

### 单位换算
| 单位 | 换算 |
|------|------|
| 1 英寸 = 2.54 cm | |
| 1 磅 (pt) = 1/72 英寸 | |
| 行高 (ht) | 磅 |
| 列宽 (width) | 字符平均宽度 |

### 评论检查清单（新增）
- [ ] 底纹颜色: patternFill + fgColor theme + tint（theme=5蓝/6红/7绿/8水绿/9橙）
- [ ] 边框双线: style="double" + 颜色 ARGB
- [ ] 条件格式 cfRule: type(如cellIs) + operator(greaterThanOrEqual) + formula + dxf引用
- [ ] 数据条: extLst → dataBar + cfvo type="min"/"max" + color
- [ ] 图表标题字号: chart1.xml → sz="1200"(12pt), sz="1400"(14pt) — 实际值×100
- [ ] 页面设置起始页码: pageSetup firstPageNumber 属性
- [ ] 打印区域: pageSetup → definedName → printArea

---

## 精确枚举参考（ISO/IEC 29500，PDF 深度分析提取）

### 内置数字格式 ID 完整表
| ID | formatCode | 说明 |
|----|------------|------|
| 0 | `General` | 常规 |
| 1 | `0` | 整数 |
| 2 | `0.00` | 两位小数 |
| 3 | `#,##0` | 千分位 |
| 4 | `#,##0.00` | 千分位两位小数 |
| 9 | `0%` | 百分比 |
| 10 | `0.00%` | 百分比两位小数 |
| 14 | `m/d/yyyy` | 日期 |
| 20 | `h:mm` | 时间 |
| 22 | `m/d/yyyy h:mm` | 日期时间 |
| 37 | `#,##0 ;(#,##0)` | 会计 |
| 38 | `#,##0 ;[Red](#,##0)` | 会计负红 |
| 49 | `@` | 文本 |
| 164+ | 自定义 | `¥#,##0.0` |

### CellValues 枚举
```python
# Excel 单元格数据类型 (c 元素 t 属性)
Boolean = "b"       # 布尔
Number = "n"        # 数值
SharedString = "s"  # 共享字符串(最常见)
String = "str"      # 公式结果字符串
Error = "e"         # 错误值
InlineString = "inlineStr"  # 内联字符串
Date = "d"          # 日期(已弃用)
```

### 图表元素参考
| 元素 | 属性 | 可能值 |
|------|------|--------|
| `<legendPos/>` | val | `"b"`(下), `"t"`(上), `"l"`(左), `"r"`(右), `"tr"`(右上) |
| `<axPos/>` | val | `"b"`(底部), `"l"`(左), `"r"`(右), `"t"`(顶部) |
| `<crosses/>` | val | `"autoZero"`, `"max"`, `"min"` |
| `<crossBetween/>` | val | `"between"`, `"midCat"` |
| `<orientation/>` | val | `"minMax"` |
| `<tickLblPos/>` | val | `"high"`, `"low"`, `"nextTo"`, `"none"` |

### 条件格式 cfRule 运算符
| operator 值 | 含义 |
|-------------|------|
| `greaterThan` | > |
| `greaterThanOrEqual` | ≥ |
| `lessThan` | < |
| `lessThanOrEqual` | ≤ |
| `equal` | = |
| `notEqual` | ≠ |
| `between` | 介于 |
| `containsText` | 包含文本 |

### 条件格式 cfvo 类型
| type 值 | 含义 |
|---------|------|
| `num` | 数值 |
| `percent` | 百分比 |
| `formula` | 公式 |
| `percentile` | 百分位 |
| `autoMin` / `autoMax` | 自动(数据条) |
| `min` / `max` | 最小/最大 |

### PatternFill patternType 完整值
`"none"`, `"solid"`, `"mediumGray"`, `"darkGray"`, `"lightGray"`,
`"darkHorizontal"`, `"darkVertical"`, `"darkDown"`, `"darkUp"`,
`"darkGrid"`, `"darkTrellis"`, `"lightHorizontal"`, `"lightVertical"`,
`"lightDown"`, `"lightUp"`, `"lightGrid"`, `"lightTrellis"`,
`"gray125"`, `"gray0625"`

### 表样式属性
| 属性 | 说明 |
|------|------|
| `name` | 如 `"TableStyleLight11"`, `"TableStyleMedium2"` |
| `showFirstColumn` | 首列特殊格式 |
| `showLastColumn` | 末列特殊格式 |
| `showRowStripes` | 行条纹 |
| `showColumnStripes` | 列条纹 |

---

## 架构级参考（ISO/IEC 29500:2016 字符级）

### CT_Cell 属性
| 属性 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `r` (cell reference) | ST_CellRef | — | 如 `"A1"`, `"B3"` |
| `s` (style index) | ST_UnsignedInt | `0` | 指向 cellXfs 索引 |
| `t` (cell type) | ST_CellType | `"n"` | b/n/s/str/e/inlineStr |
| `cm` (cell metadata) | ST_UnsignedInt | — | 元数据索引 |
| `vm` (value metadata) | ST_UnsignedInt | — | 值元数据 |
| `ph` (show phonetic) | ST_Boolean | `false` | 显示拼音 |

### CT_Row 属性
| 属性 | 类型 | 默认 | 说明 |
|------|------|------|------|
| `r` (row index) | ST_UnsignedInt | — | 如 `"1"`, `"10"` |
| `spans` | ST_CellSpans | — | 如 `"1:5"` |
| `s` (style index) | ST_UnsignedInt | `0` | 行默认样式 |
| `customFormat` | ST_Boolean | `false` | 自定义格式 |
| `ht` (height) | ST_Double | — | 行高(磅) |
| `hidden` | ST_Boolean | `false` | 隐藏行 |
| `customHeight` | ST_Boolean | `false` | 手动行高 |
| `outlineLevel` | ST_UnsignedByte | `0` | 分级显示 |
| `collapsed` | ST_Boolean | `false` | 折叠 |

### Excel 度量单位
| 属性 | 单位 | 说明 |
|------|------|------|
| `ht` (行高) | 磅 (pt) | 直接磅值 |
| `width` (列宽) | 字符宽度 | ≈7像素/字符 |
| `customWidth` | — | 存在=自定义 |
| `defaultRowHeight` | 磅 | 默认 15pt |
| `defaultColWidth` | 字符 | 默认 8.43 |

### 主题色 tint 值映射
| 描述 | theme | tint | 实际颜色 |
|------|-------|------|---------|
| "白色，背景1，深色5%" | 1 | -0.05 | 略灰白 |
| "蓝色，强调文字颜色1，淡色80%" | 5 | 0.79998 | 浅蓝 |
| "水绿色，强调文字颜色5，淡色80%" | 8 | 0.79998 | 浅水绿 |

### Transitional vs Strict
- **Transitional**: `http://schemas.openxmlformats.org/`（Office 默认）
- **Strict**: `https://purl.oclc.org/ooxml/`（ISO 标准）
- 评分时: 用 `re.search(r'xmlns[^"]*spreadsheetml')` 兜底，不硬编码命名空间前缀

## 判分点全表（含题库未考但可能出现的）

> 完整版（三科合并）在 `student_exam_app/tools/reference/判分点全表.md`；这里只放本科目。新题先对表，再写代码。

### 工作簿与工作表

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `workbook.xml/sheets/sheet@name/sheetId/state` | 表名与隐藏 | 改名判 `name`；`state="hidden"/"veryHidden"` 是隐藏表 |
| `xl/_rels/workbook.xml.rels` | 表→文件映射 | sheetId ≠ sheetN.xml，必须走 rels（老陷阱） |
| `sheetPr/sheetPr/tabColor@rgb/theme` | 工作表标签颜色 | "蓝色"可能是主题色 acc1，不是 rgb |
| `sheetPr/outlinePr`、`pageSetUpPr` | 分级/打印设置 | `summaryBelow`、`fitToPage` |
| `sheetViews/pane` | 冻结窗格 | `state="frozen"` + `xSplit/ySplit/topLeftCell` |
| `cols/col@width/customWidth/bestFit` | 列宽 | "自动调整列宽"= `bestFit="1"`；宽度单位是字符数 |
| `row@ht/customHeight/hidden/outlineLevel` | 行高/隐藏/分级 | 分类汇总的分级在行上；隐藏行用于筛选判定 |
| `mergeCells/mergeCell@ref` | 合并单元格 | 与"跨列居中"（centerContinuous）不同 |
| `autoFilter@ref` + `filterColumn` | 自动筛选 | 结果行 `hidden="1"`；高级筛选另有条件区 + `_FilterDatabase` 名称 |
| `sortState/sortCondition@ref/descending` | 排序 | 直接子节点；两列排序看 sortCondition 顺序 |
| `dataValidations/dataValidation@type/sqref` | 数据有效性 | 序列看 `formula1`；圈释无效数据在 `x14` |
| `conditionalFormatting@sqref` + `cfRule` | 条件格式 | 范围看 sqref；`dxfId` 指向样式表 dxfs |
| `hyperlinks` / `hyperlinks/hyperlink@ref` | 单元格超链接 | `r:id` 指向外部链接 rels |
| `printOptions@horizontalCentered/verticalCentered` | 打印居中 | 页面设置类考点 |
| `pageMargins@left/right/top/bottom/header/footer` | 页边距 | 单位英寸！1 厘米≈0.3937，别按 twips 判 |
| `pageSetup@paperSize/orientation/scale/fitToWidth/firstPageNumber` | 纸张缩放 | A4=9；横向 `orientation="landscape"` |
| `headerFooter/oddHeader/oddFooter` | 页眉页脚 | `&amp;L &amp;C &amp;R` 分段，`&amp;P` 页码，`&amp;N` 总页 |
| `definedNames/definedName@name` | 定义名称 | 打印标题 `_xlnm.Print_Titles`，打印区域 `_xlnm.Print_Area` |
| `sheetProtection@sheet/password` | 保护工作表 | 密码是哈希，判"是否保护"只看属性 |

### 样式与数字格式

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `numFmts/numFmt@numFmtId/formatCode` | 自定义数字格式 | 内置格式用 id（14=日期，4=#,##0.00） |
| `fonts/font@sz/name/b/i/color/u` | 字体 | 主题字体 `theme="1"` 要查 theme1.xml |
| `fills/fill@patternFill@fgColor/bgColor` | 填充 | 纯色在 fgColor；`patternType="solid"` 时 bgColor 无效 |
| `borders/border@top/left/bottom/right/diagonal` | 边框 | style 值 thin/medium/thick/double；颜色看 color@rgb |
| `cellXfs/xf@numFmtId/fontId/fillId/borderId/alignment` | 单元格格式索引 | `applyNumberFormat/applyFont…` 决定是否生效 |
| `alignment@horizontal/vertical/wrapText/textRotation/indent` | 对齐 | centerContinuous=跨列居中；shrinkToFit 缩排填充 |
| `dxfs/dxf` | 条件格式样式 | 条件格式的颜色在这里，不在 cellXfs |
| `tableStyles` | 表格样式 | 表用 `tableStyleInfo@name`（中文名） |
| 单元格 `t="s"/t="inlineStr"/t="str"` | 单元格类型 | 文本走 sharedStrings；`t="str"` 是公式结果字符串 |
| 单元格 `f` / `v` / `is` | 公式/值/富文本 | 判公式看 `f`，判结果看 `v`（缓存值可能过期，重算后为准） |

### 表、透视、图表、扩展

| 元素/属性 | 语义 | 判定要点 / 陷阱 |
| --- | --- | --- |
| `tableParts` + `tables/table1.xml` | 智能表 | `tableColumns` 是列名；`tableStyleInfo@showRowStripes` |
| `pivotTableDefinition` + `pivotCacheDefinition` | 数据透视表 | 字段名在 cacheField；行/列区看 `pivotField@axis`，值区看 `dataField@name/subtotal` |
| `xl/charts/chart1.xml` | 图表 | `c:barChart`/`c:bar3DChart`/`c:lineChart`/`c:pieChart`；标题在 `c:title/c:tx/c:rich/a:t`；图例 `c:legendPos` |
| `xdr:twoCellAnchor`（drawing1.xml） | 图表/图形位置 | from/to 行列反推区域，如 A18:G30 |
| `x14:dataBar`（extLst） | Excel 2010 数据条 | 渐变/实心、边框都在 x14，基础 dataBar 只是壳 |
| `x14:sparklineGroups` | 迷你图 | `x14:sparkline@f` 数据范围，`sqref` 放置位置 |
| `x14:conditionalFormattings` | 扩展条件格式 | 图标集、按值着色等 2010 新特性 |
| `legacyDrawing` + `vmlDrawing` | 批注/表单控件 | 批注形状在 vmlDrawing1.vml |
| `comments1.xml` | 批注文本 | 与 vml 形状配合 |
| `x14:slicerList` | 切片器 | 数据透视/表筛选器 |
| `calcPr@fullCalcOnLoad` | 打开时重算 | 判"公式是否有缓存值"时相关 |

