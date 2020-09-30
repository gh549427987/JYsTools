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
            workday = 0
            notworkday = 0
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
                        workday+=1
                        continue
                    else:
                        notworkday+=1

                    if notworkday==31:
                        return False, workday
                return True, workday
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
                        workday+=1
                        continue
                    else:
                        notworkday+=1

                    if notworkday==31:
                        return False, workday
                return True, workday
            elif which_employee is 3:
                start_index = (12,31)
                end_index = (42, 42)
                # 遍历所有的时间点
                for row in range(start_index[0], end_index[0]+1):
                    morning_sb = sheet.cell_value(row, 31)
                    morning_xb = sheet.cell_value(row, 33)
                    noon_sb = sheet.cell_value(row, 36)
                    noon_xb = sheet.cell_value(row, 38)
                    jiaban_sb = sheet.cell_value(row, 40)
                    jiaban_xb = sheet.cell_value(row, 42)


                    # 如果任意一个时间不是空的，那就是在职的
                    if morning_sb is not "" or morning_xb is not "" or noon_sb is not "" or noon_xb is not "" or \
                            jiaban_sb is not "" or jiaban_xb is not "":
                        workday+=1
                        continue
                    else:
                        notworkday+=1

                    if notworkday==31:
                        return False, workday
                return True, workday
            else:
                return None, None



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

            if which_employee == 1:
                start_index = (12,1)
                end_index = (42, 12)
                employeeName = sheet.cell_value(3, 9)
                workNum = sheet.cell_value(4, 9)
                timeCol = [1, 3, 6, 8, 10, 12]
            elif which_employee == 2:
                start_index = (12,16)
                end_index = (42, 27)
                employeeName = sheet.cell_value(3, 24)
                workNum = sheet.cell_value(4, 24)
                timeCol = [16, 18, 21, 23, 25, 27]
            elif which_employee == 3:
                start_index = (12,20)
                end_index = (42, 31)
                employeeName = sheet.cell_value(3, 39)
                workNum = sheet.cell_value(4, 39)
                timeCol = [31, 33, 36, 38, 40, 42]
            else:
                return None

            day = 0
            employee_single_data["name"] = employeeName
            employee_single_data["workNum"] = int(workNum)
            for row in range(start_index[0], end_index[0]+1):

                # 前面已经获得了工号以及名字
                # 读取所有的打卡时间
                clock_time = []
                morning_sb = sheet.cell_value(row, timeCol[0])
                morning_xb = sheet.cell_value(row, timeCol[1])
                noon_sb = sheet.cell_value(row, timeCol[2])
                noon_xb = sheet.cell_value(row, timeCol[3])
                jiaban_sb = sheet.cell_value(row, timeCol[4])
                jiaban_xb = sheet.cell_value(row, timeCol[5])

                # if morning_sb is not '': print(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[0]).value, 0))
                # if morning_xb is not '': print(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[1]).value, 0))
                # if noon_sb is not '': print(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[2]).value, 0))
                # if noon_xb is not '': print(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[3]).value, 0))
                # if jiaban_sb is not '': print(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[4]).value, 0))
                # if jiaban_xb is not '': print(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[5]).value, 0))

                if morning_sb is not '':
                    morning_sb = str(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[0]).value, 0))[11:16]
                    clock_time.append(morning_sb)
                if morning_xb is not '':
                    morning_xb = str(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[1]).value, 0))[11:16]
                    clock_time.append(morning_xb)
                if noon_sb is not '':
                    noon_sb = str(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[2]).value, 0))[11:16]
                    clock_time.append(noon_sb)
                if noon_xb is not '':
                    noon_xb = str(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[3]).value, 0))[11:16]
                    clock_time.append(noon_xb)
                if jiaban_sb is not '':
                    jiaban_sb = str(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[4]).value, 0))[11:16]
                    clock_time.append(jiaban_sb)
                if jiaban_xb is not '':

                    if type(jiaban_xb) is str:
                        jiaban_xb = sheet.cell_value(row, timeCol[5])
                    else:
                        jiaban_xb = str(xlrd.xldate_as_datetime(sheet.cell(row, timeCol[5]).value, 0))[11:16]
                    clock_time.append(jiaban_xb)

                # 收集数据
                day += 1
                employee_single_data[f"day_{day}"] = clock_time

            return employee_single_data


        # 从左向右分别读取三位员工的基本信息
        index = 0
        for sheetIndex in range(2, 21):
            sheet = self.kqbb.sheet_by_index(sheetIndex)
            index+=1
            # 第一位员工
            isonjob_1, workday = onjob(sheet, 1)
            if isonjob_1:
                self.employee_data[f"{index}"] = EmployeeData(sheet, 1)
                self.employee_data[f"{index}"]["workday"] = workday
            # 第二位员工
            isonjob_2, workday = onjob(sheet, 2)
            if isonjob_2:
                self.employee_data[f"{index}"] = EmployeeData(sheet, 2)
                self.employee_data[f"{index}"]["workday"] = workday
            # 第三位员工
            isonjob_3, workday = onjob(sheet, 3)
            print(f"{isonjob_3}")
            print(workday)
            if isonjob_3:
                self.employee_data[f"{index}"] = EmployeeData(sheet, 3)
                self.employee_data[f"{index}"]["workday"] = workday
            self.employee_data['kqrq'] = sheet.cell_value(1, 33)
            self.employee_data['zbsj'] = sheet.cell_value(2, 33)
        return self.employee_data

    def run(self):
        """
        整理所有的方法，写入所有的数据
        :return:
        """
        print(self.employee())
    pass


# ClockIn().run()
if __name__ == '__main__':

    a = ClockIn()
    a.run()