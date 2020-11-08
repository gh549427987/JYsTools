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

a.width = 0
b.width = 256*6
c.width = 256*6
d.width = 256*6
e.width = 256*6
f.width = 256*6
g.width = 256*6
h.width = 256*6
i.width = 256*6
k.width = 256*6
l.width = 256*6
m.width = 256*6
n.width = 256*6
o.width = 256*6
p.width = 256*6
q.width = 256*6
r.width = 256*6
s.width = 256*6
t.width = 256*6
u.width = 256*6
v.width = 256*6
w.width = 256*6
x.width = 256*6
y.width = 256*6
z.width = 256*6
aa.width = 256*6
ab.width = 256*6
ac.width = 256*6
ad.width = 256*6
ae.width = 256*6
af.width = 256*6
ag.width = 256*6
ah.width = 256*6
ai.width = 256*6
ak.width = 256*6
al.width = 256*6



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

# 获取数据
from Clockin import ClockIn
data = ClockIn().employee()
print(data)
worksheet.write_merge(0, 4, 8, 23, '员 工 刷 卡 记 录 表', style_title)

worksheet.write_merge(2, 2 , 25,31, data['kqrq'])
worksheet.write_merge(3, 3, 25,31, data['zbsj'])

row_first_index = 1
input_data = ''
third_row_index = 3

import traceback
try:
    # 遍历每个员工
    for j in data.keys():  # 第一个员工
        dayEachMember = 0
        if j == 'kqrq' or j == 'zbsj' or j == 'monthdays':
            continue
        print(f"第{j}个员工打卡时间录入。。。")

        # region 确认最大行高

        # endregion

        row_first_index += 4
        third_row_index += 4

        # region 行高设置
        # 设置行高
        row_first = worksheet.row(row_first_index)
        row_first_height_style = xlwt.easyxf('font:height 500;')
        row_first.set_style(row_first_height_style)
        #endregion

        # region 第一行style
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

        #endregion

        # region 第一行数据填入
        worksheet.write_merge(row_first_index, row_first_index, 1,2, '工号：', style_workNum)
        worksheet.write(row_first_index, 3,  data[j]["workNum"], style_workNum)
        for i in range(4, 9):
            borders_first_row = xlwt.Borders()
            borders_first_row.top = 5
            borders_first_row.bottom = 5
            borders_first_row.top_colour = color.LIGHT_BLUE
            borders_first_row.bottom_colour = color.LIGHT_BLUE
            style_first_row = xlwt.XFStyle()
            style_first_row.borders = borders_first_row
            worksheet.write(row_first_index, i, '', style_first_row)
        worksheet.write_merge(row_first_index, row_first_index, 9,10, '姓名：', style_name)
        worksheet.write_merge(row_first_index, row_first_index, 11,12, data[j]["name"], style_name)


        for i in range(13, 16):
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
        #endregion

        # region 第二行数据
        second_row = worksheet.row(row_first_index+1)
        second_height = xlwt.easyxf('font:height 250;')
        second_row.set_style(second_height)
        for i in range(1,34):
            # 第一个的话
            if i == 1:
                font_second_row = xlwt.Font()
                font_second_row.height = 20*9
                borders_second_row = xlwt.Borders()
                borders_second_row.left = 5
                borders_second_row.bottom = 2
                borders_second_row.right = 2
                borders_second_row.left_colour = color.LIGHT_BLUE
                borders_second_row.bottom_colour = color.LIGHT_PINK
                borders_second_row.right_colour = color.LIGHT_PINK
                alignment_second_row = xlwt.Alignment()
                alignment_second_row.vert = 0x02 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
                style_second_row = xlwt.XFStyle()
                style_second_row.borders = borders_second_row
                style_second_row.alignment = alignment_second_row
                style_second_row.font = font_second_row


        # 最后一个
            elif i == 33:
                font_second_row = xlwt.Font()
                font_second_row.height = 20*9
                borders_second_row = xlwt.Borders()
                borders_second_row.bottom = 2
                borders_second_row.right = 5
                borders_second_row.bottom_colour = color.LIGHT_PINK
                borders_second_row.right_colour = color.LIGHT_BLUE
                alignment_second_row = xlwt.Alignment()
                alignment_second_row.vert = 0x02 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
                style_second_row = xlwt.XFStyle()
                style_second_row.borders = borders_second_row
                style_second_row.alignment = alignment_second_row
                style_second_row.font = font_second_row


            else:
                # 第二个以及以后
                font_second_row = xlwt.Font()
                font_second_row.height = 20*9
                borders_second_row = xlwt.Borders()
                borders_second_row.bottom = 2
                borders_second_row.right = 2
                borders_second_row.bottom_colour = color.LIGHT_PINK
                borders_second_row.right_colour = color.LIGHT_PINK
                alignment_second_row = xlwt.Alignment()
                alignment_second_row.vert = 0x02 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
                style_second_row = xlwt.XFStyle()
                style_second_row.borders = borders_second_row
                style_second_row.alignment = alignment_second_row
                style_second_row.font = font_second_row

            if i == 32 :
                style_second_row.alignment.wrap = 1
                worksheet.write(row_first_index+1, i, '上班时间', style_second_row)
            elif i == 33:
                style_second_row.alignment.wrap = 1
                worksheet.write(row_first_index+1, i, '加班小时', style_second_row)
            else:
                worksheet.write(row_first_index+1, i, f'{i}', style_second_row)

        #endregion

        # region 第三行数据

        while True:
            dayEachMember+=1
            max_height = 0
            if dayEachMember <= data["monthdays"]:
                ct_list = data[j][f'day_{dayEachMember}'] # 获取第1天数据
                ct_list_height = len(ct_list) #判断最大行高
                if ct_list_height > max_height:
                    max_height = ct_list_height
                input_data = "\n".join(ct_list)
            elif 32 > dayEachMember > data["monthdays"]:
                input_data=""
            elif dayEachMember in [32, 33]:
                pass
            elif dayEachMember > 33:
                break
            else:
                break


            # 填入打卡时间/上班天数/加班时间
            third_row = worksheet.row(row_first_index+2)
            third_height = xlwt.easyxf(f'font:height {max_height*167};')
            third_row.set_style(third_height)
            #   筛选应该选用什么样的style
            # 第一个的话
            if dayEachMember == 1:
                font_third_row = xlwt.Font()
                font_third_row.height = 20*8
                borders_third_row = xlwt.Borders()
                borders_third_row.left = 5
                borders_third_row.bottom = 2
                borders_third_row.right = 2
                borders_third_row.left_colour = color.LIGHT_BLUE
                borders_third_row.bottom_colour = color.LIGHT_PINK
                borders_third_row.right_colour = color.LIGHT_PINK
                alignment_third_row = xlwt.Alignment()
                alignment_third_row.wrap = 1#设置自动换行
                alignment_third_row.vert = 0x00 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
                style_third_row = xlwt.XFStyle()
                style_third_row.borders = borders_third_row
                style_third_row.alignment = alignment_third_row
                style_third_row.font = font_third_row

            elif dayEachMember == 33:
                font_third_row = xlwt.Font()
                font_third_row.height = 20*8
                borders_third_row = xlwt.Borders()
                borders_third_row.bottom = 2
                borders_third_row.right = 5
                borders_third_row.bottom_colour = color.LIGHT_PINK
                borders_third_row.right_colour = color.LIGHT_BLUE
                alignment_third_row = xlwt.Alignment()
                alignment_third_row.vert = 0x00 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
                alignment_third_row.horz = 0x01
                alignment_third_row.wrap = 1#设置自动换行
                style_third_row = xlwt.XFStyle()
                style_third_row.borders = borders_third_row
                style_third_row.alignment = alignment_third_row
                style_third_row.font = font_third_row

            elif dayEachMember == 32:
                # 倒数第二个
                font_third_row = xlwt.Font()
                font_third_row.height = 20*8
                borders_third_row = xlwt.Borders()
                borders_third_row.bottom = 2
                borders_third_row.right = 2
                borders_third_row.bottom_colour = color.LIGHT_PINK
                borders_third_row.right_colour = color.LIGHT_PINK
                alignment_third_row = xlwt.Alignment()
                alignment_third_row.wrap = 1#设置自动换行
                alignment_third_row.vert = 0x00 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
                alignment_third_row.horz = 0x01
                style_third_row = xlwt.XFStyle()
                style_third_row.borders = borders_third_row
                style_third_row.alignment = alignment_third_row
                style_third_row.font = font_third_row

            else:
                # 第二个以及以后
                font_third_row = xlwt.Font()
                font_third_row.height = 20*8
                borders_third_row = xlwt.Borders()
                borders_third_row.bottom = 2
                borders_third_row.right = 2
                borders_third_row.bottom_colour = color.LIGHT_PINK
                borders_third_row.right_colour = color.LIGHT_PINK
                alignment_third_row = xlwt.Alignment()
                alignment_third_row.wrap = 1#设置自动换行
                alignment_third_row.vert = 0x00 # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
                style_third_row = xlwt.XFStyle()
                style_third_row.borders = borders_third_row
                style_third_row.alignment = alignment_third_row
                style_third_row.font = font_third_row


            # 录入所有的打卡时间
            if input_data is not '' and dayEachMember != 33 and dayEachMember != 32:
                print(f"第{dayEachMember}天录入")
                worksheet.write(third_row_index, dayEachMember, input_data, style_third_row)
            elif dayEachMember == 32:
                alignment_workday = xlwt.Alignment()
                worksheet.write(third_row_index, dayEachMember, data[j]['workday'], style_third_row)
            elif dayEachMember == 33 :
                worksheet.write(third_row_index, dayEachMember, '', style_third_row)
            else:
                worksheet.write(third_row_index, dayEachMember, '', style_third_row)
                print(f"第{dayEachMember}天没有打卡时间")
                continue
        #endregion

    from time import strftime, localtime

    # filename = strftime("%Y年%m月员工刷卡记录表.xls", localtime())
    import os
    import random
    import time
    ddd = time.time()
    filename = f"{ddd}{data['kqrq'][5:9]}年{data['kqrq'][10:12]}月员工刷卡记录表.xls"
    workbook.save(filename)
    os.system(f"open {filename}")

except:
    workbook.save('Merge_cell.xls')
    traceback.print_exc()

