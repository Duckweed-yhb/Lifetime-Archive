import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, NamedStyle

# 创建工作簿和工作表
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "我的阅读书单"

# 定义表头和列宽
headers = ["书籍名字", "作者", "分类", "类型", "状态", "开始日期", "完成日期"]
column_widths = [20, 15, 15, 10, 10, 12, 12]

# 写入表头并设置样式
header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
header_alignment = Alignment(horizontal="center", vertical="center")
thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)

for col_idx, (header, width) in enumerate(zip(headers, column_widths), 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    # 设置表头样式
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = header_alignment
    cell.border = thin_border
    # 设置列宽
    ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

# 定义下拉菜单的选项
category_options = ["重塑自我", "灵魂与理想", "学习成长", "静心修身", "技术", "历史", "其他"]
type_options = ["故事", "干货", "散文", "哲学", "技术", "历史", "小说"]
status_options = ["想读", "在读", "已读", "搁置"]

# 创建数据验证规则（下拉菜单）
dv_category = DataValidation(type="list", formula1=f'"{",".join(category_options)}"', showDropDown=True)
dv_type = DataValidation(type="list", formula1=f'"{",".join(type_options)}"', showDropDown=True)
dv_status = DataValidation(type="list", formula1=f'"{",".join(status_options)}"', showDropDown=True)

# 添加数据验证到对应列（分类=第3列，类型=第4列，状态=第5列）
ws.add_data_validation(dv_category)
ws.add_data_validation(dv_type)
ws.add_data_validation(dv_status)

# 应用下拉菜单到整列（从第2行开始，避免覆盖表头）
dv_category.add(f"C2:C1000")
dv_type.add(f"D2:D1000")
dv_status.add(f"E2:E1000")

# 给日期列设置格式
date_format = NamedStyle(name="date_style")
date_format.number_format = "YYYY-MM-DD"
wb.add_named_style(date_format)  # 先将样式添加到工作簿

# 遍历F列和G列的单元格，应用日期样式
for row in range(2, 1001):
    ws[f"F{row}"].style = date_format
    ws[f"G{row}"].style = date_format

# 给内容行设置默认样式
content_alignment = Alignment(horizontal="center", vertical="center")
for row in ws.iter_rows(min_row=2, max_row=1000, min_col=1, max_col=7):
    for cell in row:
        cell.alignment = content_alignment
        cell.border = thin_border

# 保存Excel文件
file_path = "我的阅读书单.xlsx"
wb.save(file_path)

print(f"✅ Excel表格已生成！文件路径：{file_path}")
print("📌 表格特性：")
print("  1. 分类/类型/状态列自带下拉菜单，避免输入不一致")
print("  2. 日期列自动适配YYYY-MM-DD格式")
print("  3. 表头和内容都有边框、居中对齐，样式整洁")
print("  4. 列宽已优化，内容不挤压")