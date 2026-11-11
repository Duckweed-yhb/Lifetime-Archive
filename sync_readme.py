#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""sync_readme.py —— 从根目录总表重新生成 README 的统计与书籍快照。

唯一数据源：Lifetime-Archive_总表.xlsx（工作表「总记录」）
用法：python sync_readme.py
只改 README 里两段被标记包住的内容，其余文字不动。
"""
import os
import re
import sys

import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(BASE_DIR, "Lifetime-Archive_总表.xlsx")
README = os.path.join(BASE_DIR, "README.md")

STAT_KEYS = ["书籍", "网文", "影视", "学习资料"]
BOOK_COLS = ["名称", "作者", "分类", "类型", "状态"]


def load_rows():
    wb = openpyxl.load_workbook(EXCEL, data_only=True)
    ws = wb["总记录"]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    data = []
    for r in rows[1:]:
        if not any(r):
            continue
        data.append(dict(zip(header, r)))
    return data


def build_stats(rows):
    counts = {k: 0 for k in STAT_KEYS}
    other = 0
    for r in rows:
        cat = (r.get("类别") or "").strip()
        if cat in counts:
            counts[cat] += 1
        else:
            other += 1
    parts = [f"{k} {v}" for k, v in counts.items() if v]
    if other:
        parts.append(f"其他 {other}")
    return "｜".join(parts) + f"（合计 {len(rows)} 条）"


def build_books_table(rows):
    books = [r for r in rows if (r.get("类别") or "").strip() == "书籍"]
    lines = ["| " + " | ".join(BOOK_COLS) + " |",
             "| " + " | ".join(["---"] * len(BOOK_COLS)) + " |"]
    for r in books:
        cells = []
        for col in BOOK_COLS:
            v = r.get(col)
            cells.append("" if v is None else str(v).strip())
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


def replace_block(text, name, body):
    start = f"<!-- {name}:START -->"
    end = f"<!-- {name}:END -->"
    if start not in text or end not in text:
        raise SystemExit(f"README 缺少标记 {start} / {end}")
    pattern = re.compile(re.escape(start) + r".*?" + re.escape(end), re.S)
    return pattern.sub(start + "\n" + body + "\n" + end, text)


def main():
    if not os.path.exists(EXCEL):
        sys.exit(f"找不到总表：{EXCEL}")
    rows = load_rows()
    with open(README, "r", encoding="utf-8") as f:
        text = f.read()
    text = replace_block(text, "STATS", build_stats(rows))
    text = replace_block(text, "BOOKS", build_books_table(rows))
    with open(README, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    print("已同步：" + build_stats(rows))


if __name__ == "__main__":
    main()
