# coding: utf-8
# @Time    : 2020/9/20 10:51 下午
# @Author  : 蟹蟹 ！！
# @FileName: test.py
# @Software: PyCharm

import xlwt
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('员工刷卡记录')

class color:
    LIGHT_BLUE = 0xEF7D31
    LIGHT_PINK = 45

# 设置列宽
a = worksheet.col(0)
b = worksheet.col(1)
c = worksheet.col(2)
d = worksheet.col(3)
e = worksheet.col(4)
f = worksheet.col(5)
g = worksheet.col(6)
h = worksheet.col(7)
i = worksheet.col(8)
k = worksheet.col(9)
l = worksheet.col(10)
m = worksheet.col(11)
n = worksheet.col(12)
o = worksheet.col(13)
p = worksheet.col(14)
q = worksheet.col(15)
r = worksheet.col(16)
s = worksheet.col(17)
t = worksheet.col(18)
u = worksheet.col(19)
v = worksheet.col(20)
w = worksheet.col(21)
x = worksheet.col(22)
y = worksheet.col(23)
z = worksheet.col(24)
aa = worksheet.col(25)
ab = worksheet.col(26)
ac = worksheet.col(27)
ad = worksheet.col(28)
ae = worksheet.col(29)
af = worksheet.col(30)
ag = worksheet.col(31)
ah = worksheet.col(32)
ai = worksheet.col(33)
ak = worksheet.col(34)
al = worksheet.col(35)

a.width = 256*5
b.width = 256*5
c.width = 256*5
d.width = 256*5
e.width = 256*5
f.width = 256*5
g.width = 256*5
h.width = 256*5
i.width = 256*5
k.width = 256*5
l.width = 256*5
m.width = 256*5
n.width = 256*5
o.width = 256*5
p.width = 256*5
q.width = 256*5
r.width = 256*5
s.width = 256*5
t.width = 256*5
u.width = 256*5
v.width = 256*5
w.width = 256*5
x.width = 256*5
y.width = 256*5
z.width = 256*5
aa.width = 256*5
ab.width = 256*5
ac.width = 256*5
ad.width = 256*5
ae.width = 256*5
af.width = 256*5
ag.width = 256*5
ah.width = 256*5
ai.width = 256*5
ak.width = 256*5
al.width = 256*5



# 标题字体
font_title = xlwt.Font()
font_title.name = 'name Times New Roman'
font_title.height = 20*24
font_title.bold = True

# 设置单元格对齐方式
alignment_title = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
alignment_title.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
alignment_title.vert = 0x01

# 初始化样式
style_title = xlwt.XFStyle()
style_title.font = font_title
style_title.alignment = alignment_title

worksheet.write_merge(0, 4, 8, 23, '员 工 刷 卡 记 录 表', style_title)
worksheet.write_merge(2, 2 , 25,31, '考勤日期：2020-8-1～2020-8-31')
worksheet.write_merge(3, 3, 25,31, '制表时间：2020-7-19 14:43:43')

items = 20
row_first_index = 0
for item in range(0, items):

    row_first_index += 5

    # 设置行高
    row_first = worksheet.row(row_first_index)
    row_first_height_style = xlwt.easyxf('font:height 500;')
    row_first.set_style(row_first_height_style)

    # 第一格子，工号style
    font_workNum = xlwt.Font()
    font_workNum.height = 20*11
    font_workNum.bold = True
    alignment_workNum = xlwt.Alignment()
    alignment_workNum.horz = 0x02
    alignment_workNum.vert = 0x01
    borders_workNum = xlwt.Borders()
    borders_workNum.left = 5
    borders_workNum.top = 5
    borders_workNum.bottom = 5
    borders_workNum.left_colour = color.LIGHT_BLUE
    borders_workNum.top_colour = color.LIGHT_BLUE
    borders_workNum.bottom_colour = color.LIGHT_BLUE
    style_workNum = xlwt.XFStyle()
    style_workNum.font = font_workNum
    style_workNum.alignment = alignment_workNum
    style_workNum.borders = borders_workNum

    # 第一个姓名
    font_name = xlwt.Font()
    font_name.height = 20*11
    font_name.bold = True
    alignment_name = xlwt.Alignment()
    alignment_name.horz = 0x02
    alignment_name.vert = 0x01
    borders_name = xlwt.Borders()
    borders_name.top = 5
    borders_name.bottom = 5
    borders_name.top_colour = color.LIGHT_BLUE
    borders_name.bottom_colour = color.LIGHT_BLUE
    style_name = xlwt.XFStyle()
    style_name.font = font_name
    style_name.alignment = alignment_name
    style_name.borders = borders_name

    # 第一个部门
    font_comp = xlwt.Font()
    font_comp.height = 20*11
    font_comp.bold = True
    alignment_comp = xlwt.Alignment()
    alignment_comp.horz = 0x02
    alignment_comp.vert = 0x01
    borders_comp = xlwt.Borders()
    borders_comp.top = 5
    borders_comp.bottom = 5
    borders_comp.top_colour = color.LIGHT_BLUE
    borders_comp.bottom_colour = color.LIGHT_BLUE
    style_comp = xlwt.XFStyle()
    style_comp.font = font_comp
    style_comp.alignment = alignment_comp
    style_comp.borders = borders_comp

    # 第一个公司名字
    font_compName = xlwt.Font()
    font_compName.height = 20*11
    font_compName.bold = True
    alignment_compName = xlwt.Alignment()
    alignment_compName.horz = 0x02
    alignment_compName.vert = 0x01
    borders_compName = xlwt.Borders()
    borders_compName.top = 5
    borders_compName.bottom = 5
    borders_compName.top_colour = color.LIGHT_BLUE
    borders_compName.bottom_colour = color.LIGHT_BLUE
    style_compName = xlwt.XFStyle()
    style_compName.font = font_compName
    style_compName.alignment = alignment_compName
    style_compName.borders = borders_compName


    worksheet.write_merge(row_first_index, row_first_index, 1,2, '工号：', style_workNum)
    for i in range(3, 9):
        borders_first_row = xlwt.Borders()
        borders_first_row.top = 5
        borders_first_row.bottom = 5
        borders_first_row.top_colour = color.LIGHT_BLUE
        borders_first_row.bottom_colour = color.LIGHT_BLUE
        style_first_row = xlwt.XFStyle()
        style_first_row.borders = borders_first_row
        worksheet.write(row_first_index, i, '', style_first_row)
    worksheet.write_merge(row_first_index, row_first_index, 9,10, '姓名：', style_name)
    for i in range(11, 16):
        borders_first_row = xlwt.Borders()
        borders_first_row.top = 5
        borders_first_row.bottom = 5
        borders_first_row.top_colour = color.LIGHT_BLUE
        borders_first_row.bottom_colour = color.LIGHT_BLUE
        style_first_row = xlwt.XFStyle()
        style_first_row.borders = borders_first_row
        worksheet.write(row_first_index, i, '', style_first_row)
    worksheet.write_merge(row_first_index, row_first_index, 16,17, '部门：', style_comp)
    worksheet.write_merge(row_first_index, row_first_index, 18,19, '建利灯配', style_compName)
    for i in range(20, 33):
        borders_first_row = xlwt.Borders()
        borders_first_row.top = 5
        borders_first_row.bottom = 5
        borders_first_row.top_colour = color.LIGHT_BLUE
        borders_first_row.bottom_colour = color.LIGHT_BLUE
        style_first_row = xlwt.XFStyle()
        style_first_row.borders = borders_first_row
        worksheet.write(row_first_index, i, '', style_first_row)
    # 收尾
    borders_first_row = xlwt.Borders()
    borders_first_row.top = 5
    borders_first_row.right = 5
    borders_first_row.bottom = 5
    borders_first_row.top_colour = color.LIGHT_BLUE
    borders_first_row.right_colour = color.LIGHT_BLUE
    borders_first_row.bottom_colour = color.LIGHT_BLUE
    style_first_row = xlwt.XFStyle()
    style_first_row.borders = borders_first_row
    worksheet.write(row_first_index, 33, '', style_first_row)

    second_row = worksheet.row(row_first_index+1)
    second_height = xlwt.easyxf('font:height 500;')
    second_row.set_style(second_height)
    for i in range(1,34):
        # 第一个的话
        if i == 1:
            borders_second_row = xlwt.Borders()
            borders_second_row.left = 5
            borders_second_row.bottom = 2
            borders_second_row.right = 2
            borders_second_row.left_colour = color.LIGHT_BLUE
            borders_second_row.bottom_colour = color.LIGHT_PINK
            borders_second_row.right_colour = color.LIGHT_PINK
            style_second_row = xlwt.XFStyle()
            style_second_row.borders = borders_second_row

        # 最后一个
        elif i == 33:
            borders_second_row = xlwt.Borders()
            borders_second_row.bottom = 2
            borders_second_row.right = 5
            borders_second_row.bottom_colour = color.LIGHT_PINK
            borders_second_row.right_colour = color.LIGHT_BLUE
            style_second_row = xlwt.XFStyle()
            style_second_row.borders = borders_second_row
        else:
            # 第二个以及以后
            borders_second_row = xlwt.Borders()
            borders_second_row.bottom = 2
            borders_second_row.right = 2
            borders_second_row.bottom_colour = color.LIGHT_PINK
            borders_second_row.right_colour = color.LIGHT_PINK
            style_second_row = xlwt.XFStyle()
            style_second_row.borders = borders_second_row

        if i == 32 :
            style_second_row.alignment.wrap = 1
            worksheet.write(row_first_index+1, i, '上班时间', style_second_row)
        elif i == 33:
            style_second_row.alignment.wrap = 1
            worksheet.write(row_first_index+1, i, '加班小时', style_second_row)
        else:
            worksheet.write(row_first_index+1, i, f'{i}', style_second_row)

    third_row = worksheet.row(row_first_index+2)
    third_height = xlwt.easyxf('font:height 1000;')
    third_row.set_style(third_height)
    for i in range(1,34):
        # 第一个的话
        if i == 1:
            borders_second_row = xlwt.Borders()
            borders_second_row.left = 5
            borders_second_row.bottom = 2
            borders_second_row.right = 2
            borders_second_row.left_colour = color.LIGHT_BLUE
            borders_second_row.bottom_colour = color.LIGHT_PINK
            borders_second_row.right_colour = color.LIGHT_PINK
            style_second_row = xlwt.XFStyle()
            style_second_row.borders = borders_second_row

        # 最后一个
        elif i == 33:
            borders_second_row = xlwt.Borders()
            borders_second_row.bottom = 2
            borders_second_row.right = 5
            borders_second_row.bottom_colour = color.LIGHT_PINK
            borders_second_row.right_colour = color.LIGHT_BLUE
            style_second_row = xlwt.XFStyle()
            style_second_row.borders = borders_second_row
        else:
            # 第二个以及以后
            borders_second_row = xlwt.Borders()
            borders_second_row.bottom = 2
            borders_second_row.right = 2
            borders_second_row.bottom_colour = color.LIGHT_PINK
            borders_second_row.right_colour = color.LIGHT_PINK
            style_second_row = xlwt.XFStyle()
            style_second_row.borders = borders_second_row

        worksheet.write(row_first_index+2, i, '', style_second_row)



workbook.save('Merge_cell.xls')
import os
os.system("open Merge_cell.xls")