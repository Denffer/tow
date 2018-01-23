# coding:utf-8 # for Chinese character
import os, json, re, sys, xlrd, datetime, xlsxwriter, linecache
from operator import itemgetter
import Tkinter as tk
import tkFileDialog as filedialog
import tkMessageBox
reload(sys)  # Reload does the trick!
sys.setdefaultencoding('utf8')

class Check:
    """ This program aims to
        1. take (1) data/106年保管費清單.xls (2) data/106年移置費清單.xls (3) data/期間查詢1060101-1061031.xls as input
        2. compare receipt_no and plate_no
        3. render result in result.xls
    """

    def __init__(self):
        """ initialize """
        self.src_tow, self.src_keep, self.src_duration = "", "", ""
        self.tow_data, self.keep_data, self.duration_data = [], [], []
        self.undo_tow_data, self.undo_keep_data = [], []
        self.tow_receipt_errors, self.keep_receipt_errors = [], []
        self.date_errors, self.time_errors = [], []

        self.workbook = xlsxwriter.Workbook(u"比對結果.xlsx")
        self.worksheet1 = self.workbook.add_worksheet("移置費錯誤")
        self.worksheet2 = self.workbook.add_worksheet("保管費錯誤")
        self.worksheet3 = self.workbook.add_worksheet("日期錯誤")
        self.worksheet4 = self.workbook.add_worksheet("時間錯誤")
        self.center_format = self.workbook.add_format({'align': 'center'})

    def prompt_file(self):
        """ prompt user to choose file """

        root = tk.Tk()
        root.withdraw()
        # set python to the top among the windows
        os.system('''/usr/bin/osascript -e 'tell app "Finder" to set frontmost of process "Python" to true' ''')

        # tow
        tkMessageBox.showinfo("Prompt Message", "請選取'移置費清單'", parent = root)
        self.src_tow = filedialog.askopenfilename(title="請選取'移置費清單'")

        if not self.src_tow:
            tkMessageBox.showerror("Error Message", "'移置費清單'選取錯誤，請重新執行程式。")
            exit()

        # keep
        tkMessageBox.showinfo("Prompt Message", "請選取'保管費清單'")
        self.src_keep = filedialog.askopenfilename(title="請選取'保管費清單'")
        #p.save_book_as(file_name=self.src_keep, dest_file_name=self.src_keep2)

        if not self.src_keep:
            tkMessageBox.showerror("Error Message", "'保管費清單'選取錯誤，請重新執行程式。")
            exit()

        # duration
        tkMessageBox.showinfo("Prompt Message", "請選取'期間查詢清單'")
        self.src_duration = filedialog.askopenfilename(title="請選取'期間查詢清單'")

        #p.save_book_as(file_name=self.src_duration, dest_file_name=self.src_duration2)
        if not self.src_duration:
            tkMessageBox.showerror("Error Message", "'期間查詢清單'選取錯誤，請重新執行程式。")
            exit()

        root.update()

    def get_tow_source(self):
        """ load data from data/results/ """

        #print "Collecting tow source ..."
        book = xlrd.open_workbook(self.src_tow)
        sheet = book.sheets()[0]
        raw_data = []
        for i in xrange(sheet.nrows):
            raw_data.append(sheet.row_values(i))

        data = []
        for i in raw_data:
            row_data = []
            try:
                if int(i[4]) >= 600:  # 金額 has value
                    for j in i:
                        if j: # element has value
                            row_data.append(j.encode('utf8'))
                data.append(row_data)
            except:
                pass
        #print str(data).decode('string_escape')
        # Ex: ['10601 車輛移置費', '08632', '600', '蕭伯偵(8155-YE)', '汽車違停', '1061002', '13858360015126', '已銷帳(1061003)']

        plate_no = ""
        data_len = len(data)
        cnt = 0
        for row in data:
            cnt+=1
            try:
                #print str(row).decode('string_escape')
                if int(row[1]) > 0: # if 單據編號 has value
                    if int(row[1]) == 0:
                        receipt_no = 0
                    elif len(str(int(row[1]))) == 1:
                        receipt_no = "0000" + str(int(row[1]))
                    elif len(str(int(row[1]))) == 2:
                        receipt_no = "000" + str(int(row[1]))
                    elif len(str(int(row[1]))) == 3:
                        receipt_no = "00" + str(int(row[1]))
                    elif len(str(int(row[1]))) == 4:
                        receipt_no = "0" + str(int(row[1]))
                    else:
                        receipt_no = str(int(row[1]))

                    plate_no = re.sub(r'[^a-zA-Z0-9\-]', '', row[3])

                    year = str(int(row[5][:3]) + 1911)
                    month = str(int(row[5][3:5]))
                    day = str(int(row[5][5:7]))

                    # redeem date
                    date = year + "/" + month + "/" + day

                else: # 單據編號 has no value
                    pass

                if len(row) > 7: # if 撤銷
                    self.tow_data.append([date, plate_no, receipt_no])
                else:
                    pass

            except:
                #self.PrintException()
                pass

            sys.stdout.write("\rStatus: %s / %s"%(cnt, data_len))
            sys.stdout.flush()
        #print self.tow_data

    def get_keep_source(self):
        """ load data from data/results/ """

        print "\nCollecting keep source ..."
        book = xlrd.open_workbook(self.src_keep)
        sheet = book.sheets()[0]
        raw_data = []
        for i in xrange(sheet.nrows):
            raw_data.append(sheet.row_values(i))

        data = []
        for i in raw_data:
            row_data = []
            try:
                if int(str(i[4])) >= 100:  # 金額 has value
                    for j in i:
                        if j: # element has value
                            row_data.append(j.encode('utf8'))
                data.append(row_data)
            except:
                pass
        #print str(data).decode('string_escape')
        # Ex: ['10602 車輛保管費', '09445', '100', '林憲男(3633-S2)', '汽車一日', '1061031', '13858960016617', '已銷帳(1061101)']

        plate_no = ""
        data_len = len(data)
        cnt = 0
        for row in data:
            #print str(row).decode('string_escape')
            cnt+=1
            try:
                if int(row[1]) > 0: # if 單據編號 has value
                    if int(row[1]) == 0:
                        receipt_no = 0
                    elif len(str(int(row[1]))) == 1:
                        receipt_no = "0000" + str(int(row[1]))
                    elif len(str(int(row[1]))) == 2:
                        receipt_no = "000" + str(int(row[1]))
                    elif len(str(int(row[1]))) == 3:
                        receipt_no = "00" + str(int(row[1]))
                    elif len(str(int(row[1]))) == 4:
                        receipt_no = "0" + str(int(row[1]))
                    else:
                        receipt_no = str(int(row[1]))

                    plate_no = re.sub(r'[^a-zA-Z0-9\-]', '', row[3])

                    year = str(int(row[5][:3]) + 1911)
                    month = str(int(row[5][3:5]))
                    day = str(int(row[5][5:7]))

                    # redeem date
                    date = year + "/" + month + "/" + day
                    #self.keep_data.append([date, plate_no, receipt_no])

                else: # 單據編號 has no value
                    pass

                if len(row) > 7: # if 撤銷
                    self.keep_data.append([date, plate_no, receipt_no])
                else:
                    pass

            except:
                #self.PrintException()
                pass

            sys.stdout.write("\rStatus: %s / %s"%(cnt, data_len))
            sys.stdout.flush()


        #print self.keep_data

    def get_duration_source(self):
        """ load data from data/results/ """

        print "\nCollecting duration source ..."
        book = xlrd.open_workbook(self.src_duration)
        sheet = book.sheets()[0]
        raw_data = []
        for i in xrange(sheet.nrows):
            raw_data.append(sheet.row_values(i))

        data = []
        for i in raw_data[1:]:
            row_data = []
            try:
                for j in i:
                    row_data.append(j.encode('utf8'))
                data.append(row_data)
            except:
                pass
        #print str(data).decode('string_escape')
        #['RB0565148', '違停拖吊', '0187-L2', '自小客車', '基隆市仁愛區', '愛一路68號', '第一分局', '56條1項1款', '紅線停車', '2017年5月19日', '', '呂朝舜016', '20:00', '2017年5月19日', '廖本溪0319', '', '04400', '04401', '2017年5月19日', '20:35']

        plate_no = ""
        data_len = len(data)
        receipt_no1, receipt_no2 = "", ""
        cnt = 0
        for row in data:
            cnt += 1
            try:
                plate_no = re.sub(r'[^a-zA-Z0-9\-]', '', row[2])

                # 移置單據編號
                if str(row[16]) == "":
                    receipt_no = 0
                elif len(str(int(row[16]))) == 1:
                    receipt_no1 = "0000" + str(int(row[16]))
                elif len(str(int(row[16]))) == 2:
                    receipt_no1 = "000" + str(int(row[16]))
                elif len(str(int(row[16]))) == 3:
                    receipt_no1 = "00" + str(int(row[16]))
                elif len(str(int(row[16]))) == 4:
                    receipt_no1 = "0" + str(int(row[16]))
                else:
                    receipt_no1 = str(int(row[16]))
                #print receipt_no1
                # 保管單據編號
                row[17] = row[17].replace(".","")
                if str(row[17]) == "":
                    receipt_no2 = 0
                elif len(str(int(row[17]))) == 1:
                    receipt_no2 = "0000" + str(int(row[17]))
                elif len(str(int(row[17]))) == 2:
                    receipt_no2 = "000" + str(int(row[17]))
                elif len(str(int(row[17]))) == 3:
                    receipt_no2 = "00" + str(int(row[17]))
                elif len(str(int(row[17]))) == 4:
                    receipt_no2 = "0" + str(int(row[17]))
                else:
                    receipt_no2 = str(int(row[17]))
                # print receipt_no2

                # 查扣日期
                detain_date = re.sub(r'[^0-9]+', '/', row[13])[:-1]
                #print detain_date

                # 拖吊日期
                tow_date = re.sub(r'[^0-9]+', '/', row[9])[:-1]
                #print tow_date

                # 拖吊時間
                tow_time = row[10]
                #print tow_time

                # 進場時間
                in_time = row[12]
                #print in_time

                # 領車時間
                redeem_date = re.sub(r'[^0-9]+', '/', row[18])[:-1]
                # print redeem

                self.duration_data.append([cnt, redeem_date, plate_no, receipt_no1, receipt_no2, detain_date, tow_date, tow_time, in_time])

            except:
                self.PrintException()
                pass

            sys.stdout.write("\rStatus: %s / %s"%(cnt, data_len))
            sys.stdout.flush()
        #print self.duration_data

    def check_date(self):
        """ check if 1) detain_date > tow_date 2) tow_time > in_time """

        print "\nChecking if detain_date > tow_date or tow_time > in_time ..."
        #print "[Index, redeem_date,  plate_no, receipt_no1, receipt_no2, detain_date, tow_date, tow_time, in_time]"
        cnt = 0
        for i in self.duration_data:
            # 1)
            detain_year = int(i[5].split("/")[0])
            detain_month = int(i[5].split("/")[1])
            detain_day = int(i[5].split("/")[2])
            detain_date = datetime.date(detain_year, detain_month, detain_day)

            tow_year = int(i[6].split("/")[0])
            tow_month = int(i[6].split("/")[1])
            tow_day = int(i[6].split("/")[2])
            tow_date = datetime.date(tow_year, tow_month, tow_day)

            if detain_date > tow_date:
                cnt+=1
                #print cnt, ": detain_date > tow_date in", i
                self.date_errors.append([str(i[1]), str(i[2]), str(i[5]), str(i[6])])
            # 2)
            try:
                tow_hour = int(str(i[7]).split(":")[0].replace("0",""))
                tow_minute = int(str(i[7]).split(":")[1].replace("0",""))
                tow_time = datetime.datetime(tow_year, tow_month, tow_day, hour=tow_hour, minute=tow_minute)
                in_hour = int(i[8].split(":")[0])
                in_minute = int(i[8].split(":")[1])
                in_time = datetime.datetime(tow_year, tow_month, tow_day, hour=in_hour, minute=in_minute)

                if tow_time > in_time:
                    cnt+=1
                    #print cnt, ": tow_time > in_time in", i
                    self.time_errors.append([str(i[1]), str(i[2]), str(i[7]), str(i[8])])
            except:
                #tow_time is nan
                pass
        #print self.time_errors

    def check_receipts(self):
        """ check if receipt is correct """

        print "Running cross-reference on tow receipts ..."
        cnt = 0
        for i in self.duration_data:
            for j in self.tow_data:
                try:
                    # i[1] is plate_no # check date and plate_no
                    if i[2] == j[1] and i[1] == j[0]:
                        # if receipt number doesn't match
                        if i[3] != j[2]:
                            cnt += 1
                            #print cnt, "拖吊單據不合 ->", "場內:", i[3], "部隊:", j[2]
                            # tow_data, redeem_date, plate_no, 場內單據, 部隊單據
                            # ['2017/8/7', '2017/8/7', 'ATN-5091', '06937', '06936']
                            self.tow_receipt_errors.append([i[6],i[1],i[2],i[3],j[2]])
                        else:
                            # everything correct
                            pass
                    else:
                        pass
                except:
                    self.PrintException()
                    #pass

        print "Running cross-reference on keep receipts ..."
        cnt = 0
        for i in self.duration_data:
            for j in self.keep_data:
                try:
                    # i[1] is plate_no # check date and plate_no
                    if i[1] == j[0] and i[2] == j[1]:
                        # if receipt number doesn't match
                        if i[4] != j[2]:
                            cnt += 1
                            #print cnt, "保管單據不合 ->", "場內:", i[4], "部隊:", j[2]
                            self.keep_receipt_errors.append([i[6],i[1],i[2],i[4],j[2]])
                        else:
                            # everything correct
                            pass
                    else:
                        pass
                except:
                    self.PrintException()
                    #pass
        #print self.keep_receipt_errors

    def write_excel(self):
        """ render everything into three sheets """

        print "Rendering excel file ..."
        # sheet1 # write headers
        self.worksheet1.write("A1", '錯誤類型', self.center_format)
        self.worksheet1.write("B1", '拖吊時間', self.center_format)
        self.worksheet1.write("C1", '填用時間', self.center_format)
        self.worksheet1.write("D1", '車號', self.center_format)
        self.worksheet1.write("E1", '隊部單號', self.center_format)
        self.worksheet1.write("F1", '拖吊場單號', self.center_format)

        self.worksheet1.set_column(0, 0, 16)
        self.worksheet1.set_column(1, 1, 11)
        self.worksheet1.set_column(2, 4, 10)
        self.worksheet1.set_column(5, 5, 16)

        row = 1
        for row_data in self.tow_receipt_errors:
            column = 0
            self.worksheet1.write(row, column, '拖吊單據不合', self.center_format)
            for element in row_data:
                column+=1
                self.worksheet1.write(row, column, element, self.center_format)
            row+=1

        # sheet2 # write headers
        self.worksheet2.write("A1", '錯誤類型', self.center_format)
        self.worksheet2.write("B1", '拖吊類型', self.center_format)
        self.worksheet2.write("C1", '填用時間', self.center_format)
        self.worksheet2.write("D1", '車號', self.center_format)
        self.worksheet2.write("E1", '隊部單號', self.center_format)
        self.worksheet2.write("F1", '拖吊場單號', self.center_format)

        self.worksheet2.set_column(0, 0, 16)
        self.worksheet2.set_column(1, 1, 11)
        self.worksheet2.set_column(2, 4, 10)
        self.worksheet2.set_column(5, 5, 16)

        row = 1
        for row_data in self.keep_receipt_errors:
            column = 0
            self.worksheet2.write(row, column, '保管單據不合', self.center_format)
            for element in row_data:
                column+=1
                self.worksheet2.write(row, column, element, self.center_format)
            row+=1

        # sheet3 # write headers
        self.worksheet3.write("A1", '錯誤類型', self.center_format)
        self.worksheet3.write("B1", '領車日期', self.center_format)
        self.worksheet3.write("C1", '車號', self.center_format)
        self.worksheet3.write("D1", '查扣日期', self.center_format)
        self.worksheet3.write("E1", '拖吊日期', self.center_format)

        self.worksheet3.set_column(0, 0, 20)
        self.worksheet3.set_column(1, 1, 11)
        self.worksheet3.set_column(2, 4, 10)

        row = 1
        for row_data in self.date_errors:
            column = 0
            self.worksheet3.write(row, column, '查扣日期>拖吊日期')
            for element in row_data:
                column+=1
                self.worksheet3.write(row, column, element, self.center_format)
            row+=1

        # sheet4 # write headers
        self.worksheet4.write("A1", '錯誤類型', self.center_format)
        self.worksheet4.write("B1", '領車日期', self.center_format)
        self.worksheet4.write("C1", '車號', self.center_format)
        self.worksheet4.write("D1", '拖吊時間', self.center_format)
        self.worksheet4.write("E1", '進場時間', self.center_format)

        self.worksheet4.set_column(0, 0, 20)
        self.worksheet4.set_column(1, 1, 11)
        self.worksheet4.set_column(2, 4, 10)
        for row_data in self.time_errors:
            column = 0
            self.worksheet4.write(row, column, '拖吊時間>進場時間', self.center_format)
            for element in row_data:
                column+=1
                self.worksheet4.write(row, column, element, self.center_format)
            row+=1

        self.workbook.close()

    def run(self):
        self.prompt_file()
        self.get_tow_source()
        self.get_keep_source()
        self.get_duration_source()
        self.check_date()
        self.check_receipts()
        self.write_excel()

    def PrintException(self):
        exc_type, exc_obj, tb = sys.exc_info()
        f = tb.tb_frame
        lineno = tb.tb_lineno
        filename = f.f_code.co_filename
        linecache.checkcache(filename)
        line = linecache.getline(filename, lineno, f.f_globals)
        print '    Exception in ({}, LINE {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj)

if __name__ == '__main__':
    check = Check()
    check.run()

