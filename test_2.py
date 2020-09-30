# coding: utf-8
# @Time    : 2020/9/21 11:19 下午
# @Author  : 蟹蟹 ！！
# @FileName: test_2.py.py
# @Software: PyCharm


import datetime
print(datetime.datetime.now())

from time import strftime, localtime

print(strftime("%Y年%m月员工刷卡记录表", localtime()))