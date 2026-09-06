# -*- coding: utf-8 -*-
"""
把 Excel 资源表转换成 data.txt（搜索站的数据文件）。

用法：
    双击「转换.bat」，或命令行运行  python convert.py

原理：
    读取下面 EXCEL_PATH 指定的 Excel，取「第 1 列=标题、第 2 列=链接」，
    生成 data.txt（每行一条：标题|链接|标签），写到本脚本所在文件夹。
"""

import os
import sys
import re
import openpyxl

# ====== 你的 Excel 文件路径（改这里） ======
EXCEL_PATH = r'C:\Users\谢鸿扬\Desktop\(电影查询)表格视图.xlsx'
# ==========================================

# 输出到本脚本所在的文件夹（就是克隆下来的仓库文件夹）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_FILE = os.path.join(BASE_DIR, 'data.txt')

# 默认标签：你的表里都是电影，统一填"电影"。
# 以后有别的类型，可以改成"电影,游戏"之类。
DEFAULT_TAG = '电影'


def convert():
    if not os.path.exists(EXCEL_PATH):
        print('[错误] 找不到 Excel 文件：', EXCEL_PATH)
        print('请用记事本打开 convert.py，把 EXCEL_PATH 改成你的 Excel 真实路径。')
        sys.exit(1)

    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
    ws = wb.active

    lines = []
    count = 0
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:  # 第一行是表头，跳过
            continue
        # 标题里的换行、多余空格清掉
        title = re.sub(r'\s+', ' ', row[0] or '').strip()
        url = (row[1] or '').strip()
        if not title or not url:  # 空行跳过
            continue
        if '|' in title or '|' in url:  # 含竖线会破坏格式，跳过
            continue
        lines.append(title + '|' + url + '|' + DEFAULT_TAG)
        count += 1

    with open(OUT_FILE, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')

    print('[OK] 转换完成，共 %d 条资源' % count)
    print('已写入：', OUT_FILE)


if __name__ == '__main__':
    convert()
