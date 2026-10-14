# -*- coding: utf-8 -*-
"""
影视记录管理脚本 movie_list.py
负责生成/更新「影视记录.xlsx」，用于记录电影、纪录片等影视内容。

功能：
  1. 首次运行          ：生成带样式和下拉菜单的空表格（无文件时）
  2. 已有文件直接运行  ：只读取并打印当前记录数，不动任何数据
  3. --rebuild         ：读取现有数据后按标准样式重建表格（保留数据、刷新格式）

用法：
  python movie_list.py            # 生成新表或查看现有记录数
  python movie_list.py --rebuild  # 保留数据重建格式

依赖：openpyxl（pip install openpyxl）
"""

import os
import sys
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===================== 核心配置 =====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_PATH = os.path.join(BASE_DIR, "影视记录.xlsx")   # Excel 文件路径
SHEET_TITLE = "我的影视清单"        # 工作表名称

# 类型 / 地区 / 状态选项（下拉菜单用）
type_options = [
    "电影", "纪录片", "剧集", "动画", "综艺", "其他",
]
region_options = [
    "中国大陆", "中国香港", "中国台湾", "美国", "日本", "韩国",
    "英国", "法国", "印度", "其他国家",
]
status_options = ["已看", "想看", "在看", "弃看", "养肥", "待看"]

# 表头与列宽
excel_headers = ["影视名称", "类型", "导演", "地区", "状态", "开始日期", "完成日期", "评分", "备注"]
column_widths = [26, 10, 18, 12, 10, 12, 12, 8, 30]


# ===================== 构建标准表格 =====================
def build_standard_sheet(wb, data_rows=None):
    """创建带标准样式和下拉菜单的工作表，可选写入已有数据行"""
    # 如果工作簿里已有同名表则删除重建（避免重复）
    if SHEET_TITLE in wb.sheetnames:
        del wb[SHEET_TITLE]
    ws = wb.create_sheet(SHEET_TITLE)

    # 表头样式
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    # 写入表头并设置列宽
    for col_idx, (header, width) in enumerate(zip(excel_headers, column_widths), 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # 下拉菜单（数据验证）
    dv_type = DataValidation(type="list", formula1=f'"{",".join(type_options)}"', showDropDown=True)
    dv_region = DataValidation(type="list", formula1=f'"{",".join(region_options)}"', showDropDown=True)
    dv_status = DataValidation(type="list", formula1=f'"{",".join(status_options)}"', showDropDown=True)
    ws.add_data_validation(dv_type)
    ws.add_data_validation(dv_region)
    ws.add_data_validation(dv_status)
    dv_type.add("B2:B1000")
    dv_region.add("D2:D1000")
    dv_status.add("E2:E1000")

    # 内容区默认样式（居中 + 边框 + 日期格式）
    content_alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2, max_row=1000, min_col=1, max_col=len(excel_headers)):
        for cell in row:
            cell.alignment = content_alignment
            cell.border = thin_border
            if cell.column in (6, 7):  # 开始日期 / 完成日期列
                cell.number_format = "YYYY-MM-DD"

    # 备注列左对齐，更适合长文本
    for row in range(2, 1001):
        ws.cell(row=row, column=9).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # 写入已有数据行（跳过表头）
    if data_rows:
        for row_idx, row_data in enumerate(data_rows, 2):
            for col_idx, value in enumerate(row_data, 1):
                ws.cell(row=row_idx, column=col_idx, value=value)

    return ws


# ===================== 第一步：生成新表 =====================
def generate_excel():
    """生成全新的影视记录表（空表，仅表头 + 样式 + 下拉菜单）"""
    wb = openpyxl.Workbook()
    default_ws = wb.active
    wb.remove(default_ws)
    build_standard_sheet(wb)
    wb.save(EXCEL_FILE_PATH)
    wb.close()
    print(f"✅ 已生成新表格：{EXCEL_FILE_PATH}")
    return EXCEL_FILE_PATH


# ===================== 第二步：读取现有数据 =====================
def read_existing_data():
    """读取现有表格的数据行（返回二维列表，不含表头）"""
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    ws = wb[SHEET_TITLE] if SHEET_TITLE in wb.sheetnames else wb.active

    data_rows = []
    row_idx = 2
    while True:
        movie_name = ws.cell(row=row_idx, column=1).value
        if movie_name is None or str(movie_name).strip() == "":
            break
        row_data = [ws.cell(row=row_idx, column=col).value for col in range(1, len(excel_headers) + 1)]
        data_rows.append(row_data)
        row_idx += 1

    wb.close()
    return data_rows


# ===================== 第三步：保留数据重建格式 =====================
def rebuild_excel():
    """读取现有数据后按标准样式重建（保留数据、刷新格式与下拉菜单）"""
    data_rows = read_existing_data()
    wb = openpyxl.Workbook()
    default_ws = wb.active
    wb.remove(default_ws)
    build_standard_sheet(wb, data_rows)
    wb.save(EXCEL_FILE_PATH)
    wb.close()
    print(f"✅ 已重建表格，保留 {len(data_rows)} 条记录：{EXCEL_FILE_PATH}")
    return EXCEL_FILE_PATH


# ===================== 主执行逻辑 =====================
if __name__ == "__main__":
    rebuild_mode = "--rebuild" in sys.argv

    if not os.path.exists(EXCEL_FILE_PATH):
        generate_excel()
    elif rebuild_mode:
        rebuild_excel()
    else:
        count = len(read_existing_data())
        print(f"📌 {EXCEL_FILE_PATH} 已存在，当前共 {count} 条记录")
        print("   如需保留数据重建格式，请运行：python movie_list.py --rebuild")

    print("📌 使用说明：")
    print("  1. 直接在 Excel 中填写/修改影视信息（类型、地区、状态有下拉菜单）")
    print("  2. 重新运行本代码可查看记录数；加 --rebuild 可刷新格式")
    print("  3. 开始日期/完成日期请按 YYYY-MM-DD 格式填写，如 2026-09-19")
    print("  4. 评分建议填 1-10 的整数或一位小数，如 9 或 8.5")
