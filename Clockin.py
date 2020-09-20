import xlrd
import xlwt
import json

with open("file.json", 'rb') as f:
    js = json.load(f)
    new_xlsx = js["刷卡记录表"]
    kqbb_xlsxpath = js["考勤报表"]

class ClockIn:

    def __self__(self):

        self.kqbb = xlrd.open_workbook(kqbb_xlsxpath)

        pass

    def sheetdate(self):
        """
        填入考勤日期和制表时间
        :return:
        """

        sheet = self.kqbb.sheet_by_index(2)
        kq_date = sheet.cell_value(1, 33)
        zb_date = sheet.cell_value(1, 34)

        pass

    def employee(self):
        """
        名字和工号填写,填入所有的打卡时间
        :return:
        """

        # 从左向右分别读取三位员工的基本信息
        for sheetIndex in range(2, 21):
            sheet = self.kqbb.sheet_by_index(sheetIndex)

            # 第一位员工
            workNum = sheet.cell_value(4, 9)

            # 第二位员工

            # 第三位员工

        pass

    def run(self):
        """
        整理所有的方法，写入所有的数据
        :return:
        """
        self.sheetdate()
        self.base_data_employee()
        self.clockin_data_employee()
    pass
kqbb = xlrd.open_workbook("考勤报表.xls")

sheet_1 = kqbb.sheet_by_index(2)

print(sheet_1.cell_value(1, 33))