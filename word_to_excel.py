#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word题库提取脚本 v2.1 (健壮版)
功能：从Word/txt文档中提取判断题、单选题、多选题，转换为标准Excel格式

依赖安装：
    pip install python-docx pandas openpyxl

使用方法：
    python word_to_excel_v2.py <文件路径> [输出excel路径] [--debug]

    示例：
    python word_to_excel_v2.py 考试题库.docx 输出题库.xlsx
    python word_to_excel_v2.py 考试题库.txt 输出题库.xlsx --debug
"""

import sys
import re
import os
import argparse

try:
    import pandas as pd
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError:
    print("请先安装依赖: pip install pandas openpyxl")
    sys.exit(1)


def extract_text(file_path):
    """从文件提取文本，支持docx和txt"""
    ext = os.path.splitext(file_path)[1].lower()

    if ext == '.docx':
        try:
            from docx import Document
            doc = Document(file_path)
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text)
            return "\n".join(full_text)
        except ImportError:
            print("❌ 读取docx需要安装 python-docx: pip install python-docx")
            sys.exit(1)
    elif ext == '.txt':
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    else:
        # 尝试按txt读取
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.read()
        except:
            print(f"❌ 不支持的文件格式: {ext}")
            sys.exit(1)


def split_sections(text):
    """按题型分割文本，支持多种格式"""
    # 定义各种可能的题型标记
    judgment_markers = ["一、判断", "判断题", "一、判断题", "【判断题】", "判断："]
    single_markers = ["二、单选", "单选题", "二、单选题", "【单选题】", "单选："]
    multi_markers = ["三、多选", "多选题", "三、多选题", "【多选题】", "多选："]

    def find_marker(text, markers):
        """找到第一个出现的标记位置"""
        for marker in markers:
            pos = text.find(marker)
            if pos != -1:
                return pos
        return -1

    judgment_start = find_marker(text, judgment_markers)
    single_start = find_marker(text, single_markers)
    multi_start = find_marker(text, multi_markers)

    # 调试输出
    print(f"  题型标记位置: 判断={judgment_start}, 单选={single_start}, 多选={multi_start}")

    # 如果找不到标记，尝试从文本中推断
    if judgment_start == -1 and single_start == -1 and multi_start == -1:
        print("  ⚠️ 未找到标准题型标记，尝试自动识别...")
        return text, "", ""

    # 按位置排序，确保分割正确
    positions = []
    if judgment_start != -1:
        positions.append((judgment_start, "judgment"))
    if single_start != -1:
        positions.append((single_start, "single"))
    if multi_start != -1:
        positions.append((multi_start, "multi"))

    positions.sort()

    # 根据排序后的位置分割文本
    judgment_text = ""
    single_text = ""
    multi_text = ""

    for i, (pos, ptype) in enumerate(positions):
        end_pos = positions[i + 1][0] if i + 1 < len(positions) else len(text)
        segment = text[pos:end_pos]

        if ptype == "judgment":
            judgment_text = segment
        elif ptype == "single":
            single_text = segment
        elif ptype == "multi":
            multi_text = segment

    return judgment_text, single_text, multi_text


def clean_text(text):
    """清理文本"""
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def validate_question(q, qtype, debug=False):
    """验证题目是否合理，过滤异常匹配"""
    # 题号应该是数字
    if not q["题号"].isdigit():
        if debug:
            print(f"    ⚠️ 跳过: 题号非数字 '{q['题号']}'")
        return False

    # 题目长度检查
    if len(q["题目"]) > 500:
        if debug:
            print(f"    ⚠️ 跳过: 题目过长 ({len(q['题目'])}字符)")
        return False

    if len(q["题目"]) < 5:
        if debug:
            print(f"    ⚠️ 跳过: 题目过短 ({len(q['题目'])}字符)")
        return False

    # 答案验证
    ans = q["答案"]
    if qtype == "判断题":
        if ans not in ["正确", "错误"]:
            if debug:
                print(f"    ⚠️ 跳过: 判断题答案异常 '{ans}'")
            return False
    elif qtype == "单选题":
        if ans not in ["A", "B", "C", "D"]:
            if debug:
                print(f"    ⚠️ 跳过: 单选题答案异常 '{ans}'")
            return False
    elif qtype == "多选题":
        # 多选题答案只能是2-4个大写字母，且必须在A-D范围内
        if not re.match(r"^[A-D]{2,4}$", ans):
            if debug:
                print(f"    ⚠️ 跳过: 多选题答案异常 '{ans}'")
            return False

    return True


def parse_judgments(text, debug=False):
    """解析判断题 - 答案严格限定为Y/N"""
    pattern = r"第(\d+)题.*?题号[:：](\d+).*?第(\d+)章\s+(.*?)\s+答案[:：]([YN])"
    matches = re.findall(pattern, text, re.DOTALL)

    results = []
    for m in matches:
        q = {
            "题型": "判断题",
            "题号": m[1],
            "题目": clean_text(m[3]),
            "选项A": "",
            "选项B": "",
            "选项C": "",
            "选项D": "",
            "答案": "正确" if m[4] == "Y" else "错误"
        }
        if validate_question(q, "判断题", debug):
            results.append(q)

    return results


def parse_single_choice(text, debug=False):
    """解析单选题 - 答案严格限定为单个A-D"""
    pattern = r"第(\d+)题.*?题号[:：](\d+).*?第(\d+)章\s+(.*?)\s+A[:：](.*?)\s+B[:：](.*?)\s+C[:：](.*?)\s+D[:：](.*?)\s+答案[:：]([A-D])"
    matches = re.findall(pattern, text, re.DOTALL)

    results = []
    for m in matches:
        q = {
            "题型": "单选题",
            "题号": m[1],
            "题目": clean_text(m[3]),
            "选项A": clean_text(m[4]),
            "选项B": clean_text(m[5]),
            "选项C": clean_text(m[6]),
            "选项D": clean_text(m[7]),
            "答案": m[8].strip()
        }
        if validate_question(q, "单选题", debug):
            results.append(q)

    return results


def parse_multi_choice(text, debug=False):
    """解析多选题 - 答案严格限定为2-4个大写字母A-D"""
    pattern = r"第(\d+)题.*?题号[:：](\d+).*?第(\d+)章\s+(.*?)\s+A[:：](.*?)\s+B[:：](.*?)\s+C[:：](.*?)\s+D[:：](.*?)\s+答案[:：]([A-D]{2,4})"
    matches = re.findall(pattern, text, re.DOTALL)

    results = []
    for m in matches:
        q = {
            "题型": "多选题",
            "题号": m[1],
            "题目": clean_text(m[3]),
            "选项A": clean_text(m[4]),
            "选项B": clean_text(m[5]),
            "选项C": clean_text(m[6]),
            "选项D": clean_text(m[7]),
            "答案": m[8].strip()
        }
        if validate_question(q, "多选题", debug):
            results.append(q)

    return results


def parse_all_questions(text, debug=False):
    """解析所有题型"""
    print("🔍 正在分割题型区块...")
    judgment_text, single_text, multi_text = split_sections(text)

    print(f"  判断题区块: {len(judgment_text)} 字符")
    print(f"  单选题区块: {len(single_text)} 字符")
    print(f"  多选题区块: {len(multi_text)} 字符")

    all_data = []

    # 解析判断题
    if judgment_text:
        print("\n📋 解析判断题...")
        j = parse_judgments(judgment_text, debug)
        print(f"  ✓ 判断题解析完成: {len(j)} 题")
        all_data.extend(j)

    # 解析单选题
    if single_text:
        print("\n📋 解析单选题...")
        s = parse_single_choice(single_text, debug)
        print(f"  ✓ 单选题解析完成: {len(s)} 题")
        all_data.extend(s)

    # 解析多选题
    if multi_text:
        print("\n📋 解析多选题...")
        m = parse_multi_choice(multi_text, debug)
        print(f"  ✓ 多选题解析完成: {len(m)} 题")
        all_data.extend(m)

    return all_data


def create_excel(data, output_path):
    """创建格式化的Excel文件"""
    df = pd.DataFrame(data)
    df = df[["题型", "题号", "题目", "选项A", "选项B", "选项C", "选项D", "答案"]]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="题库")

        workbook = writer.book
        worksheet = writer.sheets["题库"]

        # 设置列宽
        worksheet.column_dimensions["A"].width = 10
        worksheet.column_dimensions["B"].width = 10
        worksheet.column_dimensions["C"].width = 60
        worksheet.column_dimensions["D"].width = 25
        worksheet.column_dimensions["E"].width = 25
        worksheet.column_dimensions["F"].width = 25
        worksheet.column_dimensions["G"].width = 25
        worksheet.column_dimensions["H"].width = 12

        # 表头样式
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=12)
        header_align = Alignment(horizontal="center", vertical="center")

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_align

        # 边框和对齐
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if cell.column == 3:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                elif cell.column in [1, 2, 8]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in range(2, worksheet.max_row + 1):
            worksheet.row_dimensions[row].height = 30

    print(f"\n✅ Excel文件已保存: {output_path}")
    print(f"📊 总计: {len(data)} 道题目")

    # 输出各题型统计
    if data:
        types = {}
        for q in data:
            t = q["题型"]
            types[t] = types.get(t, 0) + 1
        print("\n📈 题型分布:")
        for t, c in types.items():
            print(f"   {t}: {c} 题")


def main():
    parser = argparse.ArgumentParser(description="Word题库提取工具 v2.1")
    parser.add_argument("input_file", help="输入文件路径(.docx 或 .txt)")
    parser.add_argument("output_file", nargs="?", help="输出Excel路径(可选)")
    parser.add_argument("--debug", action="store_true", help="启用调试模式，显示详细解析信息")

    args = parser.parse_args()

    input_path = args.input_file

    if args.output_file:
        output_path = args.output_file
    else:
        base_name = os.path.splitext(input_path)[0]
        output_path = base_name + "_提取结果.xlsx"

    if not os.path.exists(input_path):
        print(f"❌ 文件不存在: {input_path}")
        sys.exit(1)

    print(f"📄 正在读取: {input_path}")
    text = extract_text(input_path)
    print(f"  文本长度: {len(text)} 字符")

    data = parse_all_questions(text, debug=args.debug)

    if not data:
        print("\n⚠️ 未解析到任何题目，请检查文件格式")
        print("\n提示:")
        print("  1. 确保文件包含'一、判断'/'二、单选'/'三、多选'等题型标记")
        print("  2. 确保每道题包含: 题号、题目、选项(A/B/C/D)、答案")
        print("  3. 使用 --debug 参数查看详细解析过程")
        sys.exit(1)

    create_excel(data, output_path)
    print(f"\n🎉 全部完成！")


if __name__ == "__main__":
    main()