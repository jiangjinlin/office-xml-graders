# OOXML 结构速查（写评分器前先看这里）

规范本体是 ECMA-376 / ISO-IEC 29500。两个可查的地方：

- Microsoft Learn 的 Open XML SDK 文档 https://learn.microsoft.com/zh-cn/office/open-xml/open-xml-sdk
  —— SDK 的类名就是元素名（`SectionProperties` = `w:sectPr`），讲包结构、部件关系最清楚。
- ooxml.org https://ooxml.org/ —— 元素级参考，给了父/子元素、属性、ECMA 条款号
  （WordprocessingML 210 个、DrawingML 142 个；SpreadsheetML 和 PresentationML 目前写着
  Coming soon，Excel/PPT 走 MS 文档 + ECMA-376 Part 1 §18/§19）。
  注意它是前端渲染的，`urllib` 抓下来是空壳，`tools/fetch_ooxml_reference.py` 只是批量入口，
  单页请用能渲染的方式取。

## 我们判分时实际依赖的条款

| 检查点 | 元素/属性 | 条款 | 容易踩的坑 |
| --- | --- | --- | --- |
| 纸张/页边距 | `w:pgSz` `w:pgMar` | §17.6.11 / §17.6.10 | 单位 twips，1cm≈567；A4=11906×16838 |
| 页眉页脚 | `w:headerReference@w:type` | §17.6.5 | type 有 default/first/even 三态；奇偶页要先有 `w:evenAndOddHeaders` |
| 页码域 | `w:instrText` / `w:fldSimple@w:instr` | §17.16.5 / §17.16.12 | 直接搜 "PAGE" 会命中 NUMPAGES；ET.tostring 会加 ns0 前缀，正则要容忍 |
| 段落缩进 | `w:ind@w:firstLineChars` 等 | §17.3.1.12 | 字符单位存 1/100，200 = 2 字符 |
| 行距 | `w:spacing@w:line/@w:lineRule` | §17.3.1.29 | 固定值用 `lineRule="exact"`，倍数用 `auto` |
| 底纹 | `w:shd@w:fill/@w:themeFill` | §17.3.5 | 主题色带 tint/shade，不能只比 fill |
| 编号 | `w:numPr`→`w:num`→`w:abstractNumId` | §17.9 | 只是"编号定义存在"不算，要段落真的引用 |
| 表格边框 | `w:tblBorders` / `w:tcBorders` | §17.4.38 | 单元格级会覆盖表级 |
| 数字格式 | `numFmt@formatCode` + `cellXfs@applyNumberFormat` | §18.8.30 | 定义了但没 apply 不算生效 |
| 条件格式 | `conditionalFormatting@sqref` + `cfRule` | §18.3.1.18 | 判定范围看 sqref，不是 dxf |
| 数据条 | `cfRule@type="dataBar"` | §18.3.1.10 | Excel 2010 的渐变数据条在 `x14:dataBar`（extLst 里），基础 dataBar 只是兼容壳 |
| 排序 | `sortState` / `sortCondition@descending` | §18.3.1.91 | sortState 是 worksheet 直接子节点；descending 缺省=升序 |
| 放映/切换 | `p:transition`（2010 用 `p14:*`） | §19.3.1.50 | 百叶窗=`p:blinds`，溶解=`p:dissolve` |
| 字号 | DrawingML `a:rPr@sz` | §21.1.2.3.7 | 单位 1/100 磅，sz="4000" 是 40 磅 |
| 字体 | `a:latin`（西文）/`a:ea`（东亚） | §21.1.2.3.6 | 中文字体只在 `a:ea` 上，查 latin 会漏 |

## 本地配套

- `_core/checks.py::has_field(xml, 'PAGE')` —— 按域代码判断，别再用字符串搜。
- `_core/htmlcheck.py` —— HTML 不走 XML 解析（老页面标签不闭合），单独一套。
