# HTML-XML-评分器 → 名称定为 HTML-评分器（HTML 不是 XML，见下）

---
name: HTML-评分器
description: 通过解析网页 HTML 结构对网页制作（Dreamweaver）操作题进行精准评分。当用户需要自动评分或检查 .html/.htm 网页文件时使用此技能（框架 frameset/frame、标题、CSS 样式、图片、超链接/锚点、表格、背景色、水平线、下拉菜单/表单等）。HTML 不是 XML，不能用 xml.etree 严格解析，本技能用 GB2312 容错解码 + 宽松标签扫描实现答案 100%、素材 0 分。
path: C:\Users\23000\.workbuddy\skills\html-grader\SKILL.md
---

# HTML-评分器（网页制作操作题）

## 定位

- 评分对象：Dreamweaver 制作的老式 HTML（GB2312、标签不闭合、框架页）。
- **与 Office 三科的本质区别：HTML 不是 XML。** `xml.etree` 会因未闭合标签、
  裸 `&`、单双引号混用直接抛错，所以不用 XML 解析器，改用
  `_core/htmlcheck.py` 的容错扫描（正则找标签、属性值不区分大小写和引号）。
- 引擎：`student_exam_app/graders/dw_engine.py` + 规范表 `dw_specs.py`
  + 每题一个薄入口 `dw_grader_<N>.py`。加新题只加 spec，不写新评分器。

## 统一内核用法

```
_core/htmlcheck.py
  read_html(path)          gb2312→gbk→utf-8 依次试，返回 (text, enc)
  title_of / visible_text  标题、去标签文本（&nbsp; &gt; 已还原）
  css_text                 <style> 内容，压缩空白、小写
  find_tags(html, tag, **attrs)   属性宽松匹配（值忽略大小写，True=只要求存在）
  body_attr / frameset_attrs / frame_list / frame_srcs
  select_info              (options 列表, selected 项)
  anchor_names             <a name=...> 列表
  collect_html_files(root) {文件名小写: 路径}（按文件名定位页面，不依赖目录层级）
  locate_site_root(root)   找包含 index.html 的最浅目录
```

引擎检查元组（dw_specs 里用）：
`('title',f,v) ('text',f,v) ('css',f,*frags) ('body',f,attr,v)
('tag',f,tag,**attrs) ('attr',f,tag,attr,v) ('frame',f,src) ('noframe',f,src)
('select',f,options,selected) ('anchor',f,name) ('file',name)`

每个检查项的 `labels` 列表与 checks 一一对应，给出学生能看懂的失败原因；
引擎按 checks 数量把该项分值均分。

## 陷阱（实际踩过）

1. **编码**：素材全是 GB2312，read_html 必须多编码尝试；title/text 比较前
   各自 `re.sub(r'\s+','')`，Dreamweaver 会把文字拆进 span。
2. **CSS 匹配**：答案里 `font-family: "幼圆"` 带引号和空格，比较前压缩空白
   并整体小写；片段要连冒号写（如 `color:#ff9900`），避免匹配到别的属性。
3. **frameset 嵌套**：一个 index.html 里有多个 frameset，`('attr',...,'rows')`
   遍历所有 frameset 找任一满足即可；`<frame src>` 大小写不一（TOP1.html），
   src 比较统一小写。
4. **setup_stdout 不要重复调用**：`run()` 内部已调用；在 main 里再调一次，
   第二个 TextIOWrapper 被 GC 时会关闭底层 buffer，print 直接
   `ValueError: I/O operation on closed file`。
5. **双标准**：素材 zip（未作答）必须 0 分，结果 zip 必须 100 分。素材里
   本来就有的内容（原表格、原文字）不设检查点，只对答案真实新增/修改的
   差异设点。
6. **上传形态**：学生交整站 zip；引擎解压后用 locate_site_root 定位，
   兼容 HTMLMaterial/HTMLResult 等任意层目录名。
7. **本机沙箱**：开发服务器在判分后 rmtree 50+ 文件的工作目录会触发
   WorkBuddy safe-delete 钩子把进程硬退（生产 Linux 无此问题）——本地验证
   用 CLI（python graders/dw_grader_N.py 答案.zip）为准。

## 题库形态（高三计算机技能题库/DW-*）

```
DW-16/213.doc            题干（一个 doc 覆盖整组小题；DW-16 的 213.doc 缺失）
DW-16/213/HMD_001_1/
  考试素材/考试素材.zip   学生素材（=materials 里的 2素材.zip）
  结果素材/结果素材.zip   标准答案（=1结果.zip）
  图片/样张.jpg           样张
```
每小题 6 个操作点；素材 zip 直接复制为 materials/dw/<qid>/2素材.zip，
答案 zip 为 1结果.zip，题干用 3题干.txt（app 支持读取 txt 题干）。

## 回归

```
python graders/dw_grader_1.py materials/dw/1/1结果.zip   # 必须 100
python graders/dw_grader_1.py materials/dw/1/2素材.zip   # 必须 0
```
五套全过后再部署；线上用 curl -F file=@ASCII路径.zip POST /api/grade/dw/<qid>
验证（中文路径 curl 在 Windows 下会静默打不开文件，先复制成 ASCII 名）。
