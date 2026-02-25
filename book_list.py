import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, NamedStyle
import os

# ===================== 核心配置 =====================
EXCEL_FILE_PATH = "一生的阅读计划.xlsx"  # Excel文件路径
# 分类/类型/状态选项（下拉菜单用）
category_options = ["自我成长", "人文哲思", "学习提升", "静心疗愈", "技术深耕", "历史长河", "闲趣杂谈"]
type_options = ["叙事故事", "干货知识", "散文随笔", "哲学思辨", "技术手册", "历史纪实", "虚构小说"]
status_options = ["想读", "在读", "已读", "搁置"]

# ===================== 第一步：读取Excel中的最新书籍数据 =====================
def read_excel_books_data():
    """读取Excel中的书籍数据，返回整理后的books_data（适配README格式）"""
    # 检查Excel文件是否存在
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"⚠️ 未找到{EXCEL_FILE_PATH}，将生成初始版本")
        # 返回初始默认数据（首次运行用）
        return [
            ["书籍名字", "作者", "分类", "类型", "状态"],
            ["局外人", "阿尔贝·加缪", "人文哲思", "叙事故事", "想读"],
            ["刻意练习", "安德斯·艾利克森", "学习提升", "干货知识", "想读"],
            ["认知觉醒", "周岭", "自我成长", "干货知识", "想读"],
            ["学习之道", "芭芭拉·奥克利", "学习提升", "干货知识", "想读"],
            ["原则", "瑞·达利欧", "人文哲思", "干货知识", "想读"],
            ["沉思录", "马可·奥勒留", "人文哲思", "哲学思辨", "搁置"],
            ["经济学原理", "N. 格里高利·曼昆", "人文哲思", "干货知识", "想读"],
            ["理想国", "柏拉图", "人文哲思", "哲学思辨", "搁置"],
            ["刘少奇纪事", "佚名/编者", "历史长河", "历史纪实", "搁置"],
            ["毛泽东选集", "毛泽东", "历史长河", "历史纪实", "想读"],
            ["杀死一只知更鸟", "哈珀·李", "人文哲思", "叙事故事", "搁置"],
            ["资本论", "卡尔·马克思", "人文哲思", "干货知识", "搁置"],
            ["自卑与超越", "阿尔弗雷德·阿德勒", "人文哲思", "哲学思辨", "搁置"],
            ["巴黎圣母院", "维克多·雨果", "人文哲思", "虚构小说", "已读"],
            ["白鸟与蝙蝠", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["被讨厌的勇气", "岸见一郎、古贺史健", "自我成长", "哲学思辨", "已读"],
            ["传习录", "王阳明", "人文哲思", "哲学思辨", "已读"],
            ["恶意", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["放学后", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["飞鸟集", "泰戈尔", "静心疗愈", "散文随笔", "已读"],
            ["伽利略的苦恼", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["海子的诗", "海子", "静心疗愈", "散文随笔", "已读"],
            ["花开半季 情暖三生", "白落梅", "静心疗愈", "散文随笔", "已读"],
            ["假面饭店", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["假面前夜", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["假面山庄", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["解忧杂货店", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["你当像李白长风破浪", "随园散人", "人文哲思", "散文随笔", "已读"],
            ["你当像鸟飞往你的山", "塔拉·韦斯特弗", "自我成长", "叙事故事", "已读"],
            ["你的孤独 虽败犹荣", "刘同", "自我成长", "散文随笔", "已读"],
            ["人间词话", "王国维", "人文哲思", "散文随笔", "已读"],
            ["人生如逆旅，幸好还有苏轼", "倾蓝紫", "人文哲思", "散文随笔", "已读"],
            ["少年时鲜衣怒马", "佚名/编者", "静心疗愈", "散文随笔", "已读"],
            ["谁杀了她", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["圣女的救济", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["文化苦旅", "余秋雨", "人文哲思", "散文随笔", "已读"],
            ["我们生活在巨大的差距里", "余华", "人文哲思", "散文随笔", "已读"],
            ["嫌疑人X的献身", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["自卑与救赎", "斯蒂芬·金", "人文哲思", "虚构小说", "已读"],
            ["虚无的十字架", "东野圭吾", "闲趣杂谈", "虚构小说", "已读"],
            ["一卷大唐的风华", "白落梅", "静心疗愈", "散文随笔", "已读"],
            ["以诗之名", "席慕蓉", "静心疗愈", "散文随笔", "已读"],
            ["月亮与六便士", "毛姆", "人文哲思", "虚构小说", "已读"],
            ["资治通鉴", "司马光", "历史长河", "历史纪实", "已读"],
            ["纵马踏花向自由", "佚名/编者", "静心疗愈", "散文随笔", "已读"]
        ]
    
    # 读取Excel文件
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    ws = wb.active
    
    # 存储读取到的数据（适配README格式：只取前5列，去掉日期列）
    books_data = [["书籍名字", "作者", "分类", "类型", "状态"]]
    
    # 从第2行开始读取（跳过表头），直到空行结束
    row_idx = 2
    while True:
        # 读取当前行的前5列（书籍名字、作者、分类、类型、状态）
        book_name = ws.cell(row=row_idx, column=1).value
        author = ws.cell(row=row_idx, column=2).value
        category = ws.cell(row=row_idx, column=3).value
        book_type = ws.cell(row=row_idx, column=4).value
        status = ws.cell(row=row_idx, column=5).value
        
        # 终止条件：书籍名字为空，说明到了最后一行
        if book_name is None:
            break
        
        # 处理空值（避免读取到空单元格导致表格错位）
        author = author if author else ""
        category = category if category else ""
        book_type = book_type if book_type else ""
        status = status if status else ""
        
        # 添加到数据列表
        books_data.append([book_name, author, category, book_type, status])
        row_idx += 1
    
    wb.close()
    print(f"✅ 已从Excel读取到 {len(books_data)-1} 本书籍数据")
    return books_data

# ===================== 第二步：生成/更新Excel文件（保留样式和下拉菜单） =====================
def generate_excel():
    """生成/更新Excel文件，保留样式和下拉菜单（首次生成用，后续修改直接改Excel）"""
    # 创建工作簿和工作表
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "我的阅读书单"

    # 定义表头和列宽
    excel_headers = ["书籍名字", "作者", "分类", "类型", "状态", "开始日期", "完成日期"]
    column_widths = [25, 18, 15, 12, 10, 12, 12]

    # 写入表头并设置样式
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    for col_idx, (header, width) in enumerate(zip(excel_headers, column_widths), 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = thin_border
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    # 创建数据验证规则（下拉菜单）
    dv_category = DataValidation(type="list", formula1=f'"{",".join(category_options)}"', showDropDown=True)
    dv_type = DataValidation(type="list", formula1=f'"{",".join(type_options)}"', showDropDown=True)
    dv_status = DataValidation(type="list", formula1=f'"{",".join(status_options)}"', showDropDown=True)

    # 添加并应用下拉菜单到对应列
    ws.add_data_validation(dv_category)
    ws.add_data_validation(dv_type)
    ws.add_data_validation(dv_status)
    dv_category.add(f"C2:C1000")
    dv_type.add(f"D2:D1000")
    dv_status.add(f"E2:E1000")

    # 日期列格式设置
    date_format = NamedStyle(name="date_style")
    date_format.number_format = "YYYY-MM-DD"
    wb.add_named_style(date_format)
    for row in range(2, 1001):
        ws[f"F{row}"].style = date_format
        ws[f"G{row}"].style = date_format

    # 内容行默认样式
    content_alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows(min_row=2, max_row=1000, min_col=1, max_col=7):
        for cell in row:
            cell.alignment = content_alignment
            cell.border = thin_border

    # 写入初始书籍数据（仅首次生成时用）
    initial_books_data = read_excel_books_data()[1:]  # 跳过README表头
    for row_idx, book in enumerate(initial_books_data, 2):
        # 补充日期列空值，匹配Excel的7列结构
        book_with_date = book + ["", ""]
        for col_idx, value in enumerate(book_with_date, 1):
            ws.cell(row=row_idx, column=col_idx, value=value)

    # 保存Excel文件
    wb.save(EXCEL_FILE_PATH)
    wb.close()
    return EXCEL_FILE_PATH

# ===================== 第三步：生成/更新README.md =====================
def update_readme(books_data):
    """根据读取到的最新书籍数据，更新README.md"""
    # 1. 构建Markdown格式的书单表格
    md_table = []
    # 添加表格表头分隔线
    md_table.append("| " + " | ".join(books_data[0]) + " |")
    md_table.append("| " + " | ".join(["---"]*len(books_data[0])) + " |")
    # 添加书籍数据行
    for row in books_data[1:]:
        md_table.append("| " + " | ".join(row) + " |")
    md_table_str = "\n".join(md_table)

    # 2. 完整的README内容模板
    readme_content = f"""# 一生的阅读计划

这里是我个人的阅读书单管理仓库，记录着漫漫书途中的每一次驻足与收获。

### 📖 关于本仓库
本仓库用于存放我的个人阅读计划表「一生的阅读计划.xlsx」，通过结构化的方式管理阅读清单、进度与分类。

### 📚 我的书库
以下是当前整理的阅读书单，会随Excel文件同步更新：

{md_table_str}

### 📋 分类/类型说明
| 分类       | 适用范围                     | 类型       | 适用范围                     |
|------------|------------------------------|------------|------------------------------|
| 自我成长   | 个人认知、心态、成长类书籍   | 叙事故事   | 非虚构类故事型书籍           |
| 人文哲思   | 哲学、文学经典、社科类书籍   | 干货知识   | 方法论、原理类书籍           |
| 学习提升   | 学习方法、技能提升类书籍     | 散文随笔   | 诗歌、随笔、散文类书籍       |
| 静心疗愈   | 诗歌、散文、疗愈类书籍       | 哲学思辨   | 哲学论著、心理学类书籍       |
| 技术深耕   | 技术手册、编程类书籍（预留） | 技术手册   | 技术类工具书（预留）         |
| 历史长河   | 历史典籍、人物纪事类书籍     | 历史纪实   | 历史典籍、纪事类书籍         |
| 闲趣杂谈   | 悬疑、小说等休闲类书籍       | 虚构小说   | 小说、悬疑等虚构类书籍       |

### 🤝 交流与荐书
阅读是一场孤独的旅行，也是一次思想的共鸣。
非常欢迎各位朋友通过以下方式，向我推荐好书、交流阅读心得，或是提出宝贵的改进建议。

- **GitHub**: [Duckweed-yhb](https://github.com/Duckweed-yhb)
- **联系邮箱**: 3071974740@qq.com

期待与你在文字的世界里相遇！
"""

    # 3. 写入/更新README.md文件
    with open("README.md", "w", encoding="utf-8") as f:
        f.write(readme_content)
    print("✅ README.md已同步更新！")

# ===================== 主执行逻辑 =====================
if __name__ == "__main__":
    # 第一步：检查Excel是否存在，不存在则生成初始版本；存在则读取最新数据
    if not os.path.exists(EXCEL_FILE_PATH):
        print(f"📌 首次运行，生成初始Excel文件...")
        generate_excel()
    
    # 第二步：读取Excel中的最新书籍数据
    books_data = read_excel_books_data()
    
    # 第三步：更新README.md（基于最新数据）
    update_readme(books_data)
    
    # 打印完成信息
    print(f"🎉 同步完成！当前书单共 {len(books_data)-1} 本书")
    print("📌 使用说明：")
    print("  1. 直接修改Excel文件中的书籍信息（新增/删除/改分类/改状态等）")
    print("  2. 重新运行本代码，自动读取Excel最新内容并更新README")
    print("  3. Excel的下拉菜单、样式均保留，可正常手动编辑")