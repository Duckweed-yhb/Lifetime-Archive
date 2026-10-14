# -*- coding: utf-8 -*-
"""
网文记录管理脚本 webnovel_list.py
负责生成/更新「网文记录.xlsx」，用于记录追更的网络小说。

功能：
  1. 首次运行          ：生成带样式和下拉菜单的空表格（无文件时）
  2. 已有文件直接运行  ：只读取并打印当前记录数，不动任何数据
  3. --rebuild         ：读取现有数据后按标准样式重建表格（保留数据、刷新格式）

用法：
  python webnovel_list.py            # 生成新表或查看现有记录数
  python webnovel_list.py --rebuild  # 保留数据重建格式

依赖：openpyxl（pip install openpyxl）
"""

import os
import sys
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ===================== 核心配置 =====================
EXCEL_FILE_PATH = "webnovels/网文记录.xlsx"  # Excel 文件路径
SHEET_TITLE = "我的网文书单"       # 工作表名称

# 平台 / 类型 / 状态选项（下拉菜单用）
platform_options = [
    "起点中文网", "番茄小说", "晋江文学城", "纵横中文网", "七猫小说",
    "飞卢小说", "QQ阅读", "刺猬猫", "微信读书", "其他",
]
type_options = [
    "玄幻", "仙侠", "都市", "历史", "科幻", "悬疑",
    "言情", "武侠", "奇幻", "游戏", "体育", "轻小说", "其他",
]
status_options = ["追更", "在看", "已看", "完结", "弃书", "养肥", "待看"]

# 表头与列宽
excel_headers = ["网文名称", "作者", "平台", "类型", "状态", "开始日期", "完结日期", "当前进度", "备注"]
column_widths = [25, 18, 14, 10, 10, 12, 12, 16, 30]


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
    dv_platform = DataValidation(type="list", formula1=f'"{",".join(platform_options)}"', showDropDown=True)
    dv_type = DataValidation(type="list", formula1=f'"{",".join(type_options)}"', showDropDown=True)
    dv_status = DataValidation(type="list", formula1=f'"{",".join(status_options)}"', showDropDown=True)
    ws.add_data_validation(dv_platform)
    ws.add_data_validation(dv_type)
    ws.add_data_validation(dv_status)
    dv_platform.add("C2:C1000")
    dv_type.add("D2:D1000")
    dv_status.add("E2:E1000")

    # 内容区默认样式（居中 + 边框 + 日期格式）
    content_alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2, max_row=1000, min_col=1, max_col=len(excel_headers)):
        for cell in row:
            cell.alignment = content_alignment
            cell.border = thin_border
            if cell.column in (6, 7):  # 开始日期 / 完结日期列
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
    """生成全新的网文记录表（空表，仅表头 + 样式 + 下拉菜单）"""
    wb = openpyxl.Workbook()
    # 删除默认的 Sheet，改用标准表
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
        book_name = ws.cell(row=row_idx, column=1).value
        if book_name is None or str(book_name).strip() == "":
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
        print("   如需保留数据重建格式，请运行：python webnovel_list.py --rebuild")

    print("📌 使用说明：")
    print("  1. 直接在 Excel 中填写/修改网文信息（平台、类型、状态有下拉菜单）")
    print("  2. 重新运行本代码可查看记录数；加 --rebuild 可刷新格式")
    print("  3. 开始日期/完结日期请按 YYYY-MM-DD 格式填写，如 2026-09-12")
