import xlrd


kqbb = xlrd.open_workbook("考勤报表.xls")

sheet_1 = kqbb.sheet_by_index(2)

print(sheet_1.cell_value(1, 33))