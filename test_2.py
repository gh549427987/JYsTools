# coding: utf-8
# @Time    : 2020/9/21 11:19 下午
# @Author  : 蟹蟹 ！！
# @FileName: test_2.py.py
# @Software: PyCharm


# import datetime
# print(datetime.datetime.now())
#
# from time import strftime, localtime
#
# print(strftime("%Y年%m月员工刷卡记录表", localtime()))

import xlrd

wb = xlrd.open_workbook("考勤报表.xls")
sheet = wb.sheet_by_index(2)
print(sheet.cell_value(41,1))
# print(sheet.cell_value(42,1))
a = "考勤报表.xls"
print(a[-2:])

a = "制表时间：2020-10-10 10:53:34"
print(a[5:12])
print(a[5:9])
print(a[10:12])