# Office 操作题评分（Open XML 判分系统 + 评分技能）

解析 Office/HTML 文件底层结构做自动判分的完整工程：判分服务、40 套评分器、
配套技能文档。

- 线上：https://web.xiaoshizi.asia （学生上传作业实时判分）
- 覆盖：Excel / Word / PPT 各 10 套（2022 准易题库）、网页制作 10 套（高三技能题库）
- 硬标准：**标准答案 100 分，未作答素材 0 分**，40 套全部满足

## 目录

```
Excel/Word/PPT 2010 操作题题库/   题库源（1结果=答案、2素材=原始、3题干=要求）
高三计算机技能题库/                            网页(DW)/Word/Excel 技能题库（考试素材+结果素材）
student_exam_app/                             Flask 判分服务（部署这个目录）
  graders/                                    40 个评分器 + _core 公共内核
  materials/                                  发给学生的题目包（题干+素材，不含答案）
  templates/index.html                        前端（四科页签、上传规定、小程序码）
  tools/                                      回归脚本 + OOXML 速查表
skills/                                       评分技能
  word-xml-grader/  excel-xml-grader/  ppt-xml-grader/  html-grader/
```

## 常用命令

```
cd student_exam_app
python tools/selftest.py          # 40 个评分器跑标准答案，必须全 100
python tools/discriminate.py      # 跑未作答素材，必须全 0
python tools/why_score.py excel 3 # 看单套的每个得分点
python app.py                     # 本地起服务
python deploy.py --dry-run        # 发布预检
python deploy.py                  # 发上线
```

## 判分要点

- 只看 XML 证据：素材默认就满足的（A4、默认页边距、预置样式、已有文字）不计分，
  只对答案真实改动设检查点；字体字号按目标单元格取样，不做全文搜索。
- Office 走 `xml.etree` + 名称空间；HTML 不是 XML（标签不闭合、GB2312），
  用容错扫描 `_core/htmlcheck.py`。
- Word 页码按域代码判断（`w:instrText` / `w:fldSimple`），直接搜 "PAGE" 会误中 NUMPAGES。
- 条款号与坑位见 `student_exam_app/tools/reference/OOXML-速查.md`。

## 技能用法

`skills/` 下四份 SKILL.md 可直接放进 AI 助手的技能目录（如
`~/.workbuddy/skills/`）：Word/Excel/PPT 靠 Open XML，网页靠容错 HTML 扫描。
每份都带「权威参考」「致命陷阱」「回归验证」三节。

## 提交前检查

```
python tools/selftest.py && python tools/discriminate.py
```

两条都绿再 `python deploy.py`。改题目素材时同样跑一遍，防止素材被污染。
