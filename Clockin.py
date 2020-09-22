import xlrd
import xlwt
import json

with open("file.json", 'rb') as f:
    js = json.load(f)
    new_xlsx = js["刷卡记录表"]
    kqbb_xlsxpath = js["考勤报表"]

class ClockIn:

    employee_data = {}


    def __init__(self):

        self.kqbb = xlrd.open_workbook(kqbb_xlsxpath)
        self.jlb = xlrd.open_workbook(new_xlsx)

        self.employee_count = 0

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

        # 检验员工是否还在职，如果一次打卡时间都没有，就判定为不在职
        def onjob(sheet, which_employee):
            if which_employee is 1:
                start_index = (12,1)
                end_index = (42, 12)
                # 遍历所有的时间点
                for row in range(start_index[0], end_index[0]+1):
                    morning_sb = sheet.cell_value(row, 1)
                    morning_xb = sheet.cell_value(row, 3)
                    noon_sb = sheet.cell_value(row, 6)
                    noon_xb = sheet.cell_value(row, 8)
                    jiaban_sb = sheet.cell_value(row, 10)
                    jiaban_xb = sheet.cell_value(row, 12)


                    # 如果任意一个时间不是空的，那就是在职的
                    if morning_sb is not "" or morning_xb is not "" or noon_sb is not "" or noon_xb is not "" or \
                            jiaban_sb is not "" or jiaban_xb is not "":
                        print(f"ddddd {morning_sb} {morning_xb} {noon_sb} {noon_xb} {jiaban_sb} {jiaban_xb}")
                        return True

            elif which_employee is 2:
                start_index = (12,16)
                end_index = (42, 27)
                # 遍历所有的时间点
                for row in range(start_index[0], end_index[0]+1):
                    morning_sb = sheet.cell_value(row, 16)
                    morning_xb = sheet.cell_value(row, 18)
                    noon_sb = sheet.cell_value(row, 21)
                    noon_xb = sheet.cell_value(row, 23)
                    jiaban_sb = sheet.cell_value(row, 25)
                    jiaban_xb = sheet.cell_value(row, 27)


                    # 如果任意一个时间不是空的，那就是在职的
                    if morning_sb is not "" or morning_xb is not "" or noon_sb is not "" or noon_xb is not "" or \
                            jiaban_sb is not "" or jiaban_xb is not "":
                        print(f"ddddd {morning_sb} {morning_xb} {noon_sb} {noon_xb} {jiaban_sb} {jiaban_xb}")
                        return True
            elif which_employee is 3:
                start_index = (12,20)
                end_index = (42, 31)
                # 遍历所有的时间点
                for row in range(start_index[0], end_index[0]+1):
                    morning_sb = sheet.cell_value(row, 20)
                    morning_xb = sheet.cell_value(row, 22)
                    noon_sb = sheet.cell_value(row, 25)
                    noon_xb = sheet.cell_value(row, 27)
                    jiaban_sb = sheet.cell_value(row, 29)
                    jiaban_xb = sheet.cell_value(row, 31)


                    # 如果任意一个时间不是空的，那就是在职的
                    if morning_sb is not "" or morning_xb is not "" or noon_sb is not "" or noon_xb is not "" or \
                            jiaban_sb is not "" or jiaban_xb is not "":
                        print(f"ddddd {morning_sb} {morning_xb} {noon_sb} {noon_xb} {jiaban_sb} {jiaban_xb}")
                        return True
            else:
                return None



            print(f"{sheet} {which_employee} is not on job")
            return False

        def EmployeeData(sheet, which_employee):
            '''
            employee_single_data = {
                "name" : "xxx"
                "workNum" : "01"
                "day_1" : [8:00]
                "day_2" : [9:00, 10:00]
                ...
            }
            :param sheet:
            :param which_employee:
            :return:
            '''

            employee_single_data = {}

            if which_employee is 1:
                start_index = (12,1)
                end_index = (42, 12)
                employeeName = sheet.cell_value(3, 9)
                workNum = sheet.cell_value(4, 9)
            elif which_employee is 2:
                start_index = (12,16)
                end_index = (42, 27)
                employeeName = sheet.cell_value(3, 9)
                workNum = sheet.cell_value(4, 24)
            elif which_employee is 3:
                start_index = (12,20)
                end_index = (42, 31)
                employeeName = sheet.cell_value(3, 9)
                workNum = sheet.cell_value(4, 39)
            else:
                return None

            day = 0
            for row in range(start_index[0], end_index[0]+1):

                # 前面已经获得了工号以及名字
                # 读取所有的打卡时间
                clock_time = []
                morning_sb = sheet.cell_value(row, 1)
                morning_xb = sheet.cell_value(row, 3)
                noon_sb = sheet.cell_value(row, 6)
                noon_xb = sheet.cell_value(row, 8)
                jiaban_sb = sheet.cell_value(row, 10)
                jiaban_xb = sheet.cell_value(row, 12)

                if morning_sb is not '': clock_time.append(morning_sb)
                if morning_xb is not '': clock_time.append(morning_xb)
                if noon_sb is not '': clock_time.append(noon_sb)
                if noon_xb is not '': clock_time.append(noon_xb)
                if jiaban_sb is not '': clock_time.append(jiaban_sb)
                if jiaban_xb is not '': clock_time.append(jiaban_xb)

                # 收集数据
                employee_single_data["name"] = employeeName
                employee_single_data["workNum"] = workNum
                employee_single_data[f"day_{str(day+1)}"] = clock_time

                return employee_single_data


            pass


        # 从左向右分别读取三位员工的基本信息
        index = 0
        for sheetIndex in range(2, 21):
            sheet = self.kqbb.sheet_by_index(sheetIndex)
            index+=1
            print(f"这是第{index}个表")
            # 第一位员工
            print(f"这是该表第一位员工")
            if onjob(sheet, 1):
                self.employee_data[self.employee_count+1] = EmployeeData(sheet, 1)

            # 第二位员工
            print(f"这是该表第二位员工")
            if onjob(sheet, 2):
                self.employee_data[self.employee_count+1] = EmployeeData(sheet, 2)

            # 第三位员工
            print(f"这是该表第三位员工")
            if onjob(sheet, 3):
                self.employee_data[self.employee_count+1] = EmployeeData(sheet, 3)
        return self.employee_data

    def run(self):
        """
        整理所有的方法，写入所有的数据
        :return:
        """
        print(self.employee())
    pass


# ClockIn().run()
a = ClockIn()
a.run()