# coding: utf-8
# @Time    : 2020/9/21 11:19 下午
# @Author  : 蟹蟹 ！！
# @FileName: test_2.py.py
# @Software: PyCharm

import xlrd
from datetime import datetime
from xlrd import xldate_as_tuple

a = xlrd.open_workbook("考勤报表.xls")
sheet = a.sheet_by_index(2)

row_num = 42
row_values = sheet.row_values(row_num)
first_row_values = sheet.row_values(42)
num = 1
if row_values:
    str_obj = {}
for i in range(len(first_row_values)):
    ctype = sheet.cell(num, i).ctype
    cell = sheet.cell_value(num, i)
    if ctype == 2 and cell % 1 == 0.0:  # ctype为2且为浮点
        cell = int(cell)  # 浮点转成整型
        cell = str(cell)  # 转成整型后再转成字符串，如果想要整型就去掉该行
    elif ctype == 3:
        date = datetime(*xldate_as_tuple(cell, 0))
        cell = date.strftime('%Y/%m/%d %H:%M:%S')
    elif ctype == 4:
        cell = True if cell == 1 else False
    str_obj[first_row_values[i]] = cell
list.append(str_obj)
# 0.72430419921875