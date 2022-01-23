import asyncio
import configparser
import datetime
import random
import socket
import sys
import time
import os
import aiomysql.sa as aio_sa
import pandas as pd
import pymysql
import xlwt
import numpy
import pyodbc
import pyperclip
import win32api
import win32con
import win32gui
import matplotlib
import threading
import queue
import wx
import wx.aui
from threading import Thread
import seaborn as sns
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_template import FigureCanvas
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg

from matplotlib.figure import Figure
from matplotlib.pyplot import MultipleLocator
from pubsub import pub

from wx.lib.mixins.listctrl import CheckListCtrlMixin, ListCtrlAutoWidthMixin

from GUI import MyFrame1

class TestThread(Thread):
    def __init__(self):
        Thread.__init__(self)
        self.start()
    def run(self):
        try:
            global StopFlag
            StopFlag = 1
            time.sleep(5)
            global cnxn, crsr,Output_path
            global DataID, DataCheck, DataRec,DataRec_yesterday
            # cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=./DataSend03.accdb')
            crsr = cnxn.cursor()
            now = datetime.datetime.now().strftime("%Y-%m-%d")
            timenow = datetime.datetime.now().strftime("%Y-%m-%d")
            timelast = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d")

            # 创建excel
            workbook = xlwt.Workbook(encoding='ascii')
            # 创建新的sheet表
            worksheet = workbook.add_sheet("设备稼动总表")
            # 往表格写入标题
            style_title = xlwt.XFStyle()
            font = xlwt.Font()
            font.bold = True
            style_title.font = font
            borders = xlwt.Borders()  # Create Borders
            borders.left = xlwt.Borders.THIN
            borders.right = xlwt.Borders.THIN
            borders.top = xlwt.Borders.THIN
            borders.bottom = xlwt.Borders.THIN
            borders.left_colour = 0x40
            borders.right_colour = 0x40
            borders.top_colour = 0x40
            borders.bottom_colour = 0x40
            style_title.borders = borders

            worksheet.write(0, 0, "日期", style_title)
            worksheet.write(0, 1, "设备编号", style_title)
            worksheet.write(0, 2, "机种", style_title)
            worksheet.write(0, 3, "投入数", style_title)
            worksheet.write(0, 4, "产数数", style_title)
            worksheet.write(0, 5, "运行时间/h", style_title)
            worksheet.write(0, 6, "停机时间/h", style_title)
            worksheet.write(0, 7, "平均C/T", style_title)
            worksheet.write(0, 8, "稼动率", style_title)

            worksheet.col(0).width = 256 * 15
            worksheet.col(5).width = 256 * 15
            worksheet.col(6).width = 256 * 15
            worksheet.col(7).width = 256 * 15

            style_Contes = xlwt.XFStyle()
            style_Contes.borders = borders
            xls_Rows = 0

            for D in range(len(DataID)):
                if DataCheck[D] == 2:
                    xls2_Rows = 0
                    worksheet2 = workbook.add_sheet(DataID[D] + "故障明细记录表")
                    worksheet2.write(0, 0, "日期", style_title)
                    worksheet2.write(0, 1, "设备编号", style_title)
                    worksheet2.write(0, 2, "机种", style_title)
                    worksheet2.write(0, 3, "停机开始时间", style_title)
                    worksheet2.write(0, 4, "停机结束时间", style_title)
                    worksheet2.write(0, 5, "停机时长", style_title)
                    worksheet2.write(0, 6, "故障编号", style_title)
                    worksheet2.write(0, 7, "停机备注（产线补充）", style_title)
                    worksheet2.col(0).width = 256 * 15
                    worksheet2.col(3).width = 256 * 30
                    worksheet2.col(4).width = 256 * 30
                    worksheet2.col(7).width = 256 * 30
                    try:
                        # print(DataID[D])
                        # 每日产量原始数据
                        crsr.execute(
                            "SELECT * FROM 产量2 WHERE 时间>=#" + timelast + " 8:00:00" + "# and 时间<=#" + timenow + " 8:00:00" + "# and 机种 = '" +
                            DataID[D] + "'")
                        list = crsr.fetchall()
                        list0 = numpy.array(list)
                        LastDay_Data = list0[:, 3]
                        CL = 0
                        for Data_Rank in range(len(LastDay_Data)):
                            if Data_Rank > 0 and int(LastDay_Data[Data_Rank]) > int(LastDay_Data[Data_Rank - 1]):
                                CL = CL + (int(LastDay_Data[Data_Rank]) - int(LastDay_Data[Data_Rank - 1]))

                        DataRec_yesterday[D] = CL
                        worksheet.write(xls_Rows + 1, 0, timelast, style_Contes)
                        worksheet.write(xls_Rows + 1, 1, DataID[D], style_Contes)
                        JZ_Type = DataID[D].split("-")
                        worksheet.write(xls_Rows + 1, 2, JZ_Type[0], style_Contes)
                        worksheet.write(xls_Rows + 1, 3, str(CL), style_Contes)
                        worksheet.write(xls_Rows + 1, 4, str(CL), style_Contes)

                        TJ_AllTime = 0
                        TJ_flag = 0
                        TJ_Startflag = 0
                        TJ_Stopflag = 0
                        for i in range(1, len(list)):
                            if list[i - 1][3] == list[i][3]:
                                if TJ_flag == 0:
                                    TJ_Startflag = i
                                TJ_flag = 1
                            else:
                                if TJ_flag == 1:
                                    TJ_flag = 0
                                    delta = list[i][2] - list[TJ_Startflag][2]
                                    if int(delta.total_seconds() / 60) > 5:
                                        worksheet2.write(xls2_Rows + 1, 0, timelast, style_Contes)
                                        worksheet2.write(xls2_Rows + 1, 1, DataID[D], style_Contes)
                                        worksheet2.write(xls2_Rows + 1, 2, JZ_Type[0], style_Contes)
                                        worksheet2.write(xls2_Rows + 1, 3, str(list[TJ_Startflag][2]), style_Contes)
                                        worksheet2.write(xls2_Rows + 1, 4, str(list[i][2]), style_Contes)
                                        worksheet2.write(xls2_Rows + 1, 5, str(round(delta.total_seconds() / 60, 2)),
                                                         style_Contes)
                                        TJ_AllTime = TJ_AllTime + (round((delta.total_seconds() / 60), 2))
                                        xls2_Rows = xls2_Rows + 1
                        delta_kj = list[-1][2] - list[0][2]
                        KJ_time = delta_kj.total_seconds() / 60 - TJ_AllTime

                        if CL > 0:
                            kjsc =  str(round(KJ_time / 60, 2))
                            tjsc = str(round(24 - (KJ_time / 60), 2))
                            CTsj = str(round((KJ_time) * 60 / CL, 2))
                            jdl = str(round(KJ_time / 60 / 24, 2))

                            worksheet.write(xls_Rows + 1, 5, str(round(KJ_time / 60, 2)), style_Contes)
                            worksheet.write(xls_Rows + 1, 6, str(round(24 - (KJ_time / 60), 2)), style_Contes)
                            worksheet.write(xls_Rows + 1, 7, str(round((KJ_time) * 60 / CL, 2)) + " S/Pcs",
                                            style_Contes)
                            worksheet.write(xls_Rows + 1, 8, str(round(KJ_time / 60 / 24, 2)), style_Contes)
                            try:
                                date = "'" + datetime.datetime.now().strftime("%Y-%m-%d") + "'"
                                crsr.execute(
                                    "INSERT INTO 每日汇总(机种,日期,产量,开机时长,停机时长,CT时间,稼动率)VALUES(" + "'" + DataID[D] + "'" + "," + date + "," + str(CL) + "," + kjsc + "," + tjsc + "," + CTsj +"," +jdl+")")
                                crsr.commit()

                            except Exception as e:
                                print(e)
                        else:
                            worksheet.write(xls_Rows + 1, 5, "0", style_Contes)
                            worksheet.write(xls_Rows + 1, 6, "24", style_Contes)
                            worksheet.write(xls_Rows + 1, 7, "0 S/Pcs", style_Contes)
                            worksheet.write(xls_Rows + 1, 8, "0", style_Contes)
                            try:
                                date = "'" + datetime.datetime.now().strftime("%Y-%m-%d") + "'"
                                crsr.execute(
                                    "INSERT INTO 每日汇总(机种,日期,产量,开机时长,停机时长,CT时间,稼动率)VALUES(" + "'" + DataID[
                                        D] + "'" + "," + date + "," + str(
                                        CL) + "," + "0" + "," + "24" + "," + "0" + "," + "0" + ")")
                                crsr.commit()

                            except Exception as e:
                                print(e)
                        xls_Rows = xls_Rows + 1

                    except Exception as e:
                        print(e)

            # 保存
            workbook.save(Output_path + "\\" + "流水线生产数据报表 " + now + ".xls")
            global Data_15Day
            for D in range(len(DataID)):
                if DataCheck[D] == 2:
                    timenow = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d")
                    timelast = (datetime.datetime.now() + datetime.timedelta(days=-16)).strftime("%Y-%m-%d")
                    crsr.execute(
                        "SELECT 产量 FROM 每日汇总 WHERE 日期>=#" + timelast + "# and 日期<=#" + timenow + "# and 机种 = '" +
                        DataID[D] + "'")
                    list0 = crsr.fetchall()
                    if list0 != []:
                        for _list0 in range(15):
                            try:
                                Data_15Day[D][_list0] = list0[_list0][0]
                            except:
                                Data_15Day[D][_list0] = 0
                    else:
                        for _list0 in range(15):
                            Data_15Day[D][_list0] = 0

            time.sleep(5)
            StopFlag = 0
        except Exception as e:
            StopFlag = 0
            print(e)

class TestThread2(Thread):
    def __init__(self):
        Thread.__init__(self)
        self.start()

    def run(self):
        global StopFlag_1, StopFlag, cnxn, crsr
        global DataID, DataCheck, DataRec, D_Rank
        try:
            if StopFlag == 0:
                StopFlag_1 = 1
                time.sleep(1)
                draw_datas = []

                for D in range(len(DataID)):
                    if DataCheck[D] == 2:

                        timenow = (datetime.datetime.now() + datetime.timedelta(days=-0)).strftime("%Y-%m-%d %H:%M:%S")
                        timelast = (datetime.datetime.now() + datetime.timedelta(hours=-1)).strftime("%Y-%m-%d %H:%M:%S")
                        print( "SELECT * FROM 产量 WHERE 时间>=#" + timelast + "# and 时间<=#" + timenow + "# and 机种 = '" +
                            DataID[D] + "'")


                        crsr = cnxn.cursor()
                        crsr.execute(
                            "SELECT * FROM 产量2 WHERE 时间>=#" + timelast + "# and 时间<=#" + timenow + "# and 机种 = '" +
                            DataID[D] + "'")
                        list0 = crsr.fetchall()


                        # print(list0)
                        if list0 != [] :
                            list_Data = []
                            list_Data_x = []
                            list_Data_x1 = []

                            for i in range(len(list0)):
                                list_Data.append(list0[i][3])
                                list_Data_x.append(list0[i][2])
                                list_Data_x1.append(i)

                            d = {'a':list_Data_x,
                                 'b':list_Data}
                            df = pd.DataFrame(d)
                            df.drop_duplicates(subset=['a'], keep='first', inplace=True)
                            ts = pd.Series(df['b'].tolist(), index=df['a'])
                            ts_10T = ts.resample('1T').bfill()
                            draw_datas.append(ts_10T)
                            print(D_Rank[D])
                        else:
                            draw_datas.append([])
                            print(draw_datas[D_Rank[D]])
                wx.CallAfter(pub.sendMessage, "plt_Data", msg=draw_datas)
                time.sleep(1)
                StopFlag_1 = 0
                print("12123")
            else:

                pass
        except:
            StopFlag = 0
            pass

class TestThread3(Thread):
    def __init__(self):
        Thread.__init__(self)
        self.start()
    def run(self):
        try:
            global StopFlag
            time.sleep(1)
            global cnxn, crsr,Output_path
            global DataID, DataCheck, DataRec,DataRec_yesterday
            # cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=./DataSend03.accdb')
            crsr = cnxn.cursor()
            now = datetime.datetime.now().strftime("%Y-%m-%d")
            timenow = datetime.datetime.now().strftime("%Y-%m-%d")
            timelast = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d")

            for D in range(len(DataID)):
                if DataCheck[D] == 2:
                    try:
                        # print(DataID[D])
                        # 每日产量原始数据
                        crsr.execute(
                            "SELECT * FROM 产量2 WHERE 时间>=#" + timelast + " 8:00:00" + "# and 时间<=#" + timenow + " 8:00:00" + "# and 机种 = '" +
                            DataID[D] + "'")
                        list = crsr.fetchall()
                        list0 = numpy.array(list)
                        LastDay_Data = list0[:, 3]
                        CL = 0
                        for Data_Rank in range(len(LastDay_Data)):
                            if Data_Rank > 0 and int(LastDay_Data[Data_Rank]) > int(LastDay_Data[Data_Rank - 1]):
                                CL = CL + (int(LastDay_Data[Data_Rank]) - int(LastDay_Data[Data_Rank - 1]))

                        DataRec_yesterday[D] = CL
                    except Exception as e:
                        print(e)

            # 保存
            time.sleep(1)
        except Exception as e:
            print(e)

class TestThread4(Thread):
    def __init__(self):
        Thread.__init__(self)
        self.start()
    def run(self):
        global StopFlag_1, StopFlag, cnxn, crsr
        global DataID, DataCheck, DataRec, D_Rank, z_choice,TestThread4_flag
        try:
            if StopFlag == 0 :
                TestThread4_flag = 1
                time.sleep(0.1)
                draw_datas = []
                timenow = (datetime.datetime.now() + datetime.timedelta(days=-0)).strftime("%Y-%m-%d %H:%M:%S")
                timelast = (datetime.datetime.now() + datetime.timedelta(hours=-1)).strftime("%Y-%m-%d %H:%M:%S")

                crsr = cnxn.cursor()
                crsr.execute(
                    "SELECT * FROM 产量2 WHERE 时间>=#" + timelast + "# and 时间<=#" + timenow + "# and 机种 = '" +
                    z_choice + "'")
                list0 = crsr.fetchall()

                # print(list0)
                if list0 != []:
                    list_Data = []
                    list_Data_x = []
                    list_Data_x1 = []

                    for i in range(len(list0)):
                        list_Data.append(list0[i][3])
                        list_Data_x.append(list0[i][2])
                        list_Data_x1.append(i)

                    d = {'a': list_Data_x,
                         'b': list_Data}
                    df = pd.DataFrame(d)
                    df.drop_duplicates(subset=['a'], keep='first', inplace=True)
                    ts = pd.Series(df['b'].tolist(), index=df['a'])
                    ts_10T = ts.resample('1T').bfill()
                    wx.CallAfter(pub.sendMessage, "z_plt_Data", msg=ts_10T)
                else:
                    wx.CallAfter(pub.sendMessage, "z_plt_Data", msg=[])
                print("456")
                TestThread4_flag = 0
                time.sleep(0.1)
            else:
                TestThread4_flag = 0
                pass
        except:
            TestThread4_flag = 0
            pass

class TestThread5(Thread):
    def __init__(self):
        Thread.__init__(self)
        self.start()
    def run(self):
        global StopFlag_1, StopFlag, cnxn, crsr,z_CL
        global DataID, DataCheck, DataRec, D_Rank, z_choice,TestThread4_flag
        while TestThread4_flag == 1:
            time.sleep(1)
        try:
            if StopFlag == 0 :
                time.sleep(0.1)
                draw_datas = []
                timenow = (datetime.datetime.now() + datetime.timedelta(days=0)).strftime("%Y-%m-%d %H:%M:%S")
                timelast = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d %H:%M:%S")
                crsr = cnxn.cursor()
                print("SELECT * FROM 产量2 WHERE 时间>=#" + timelast + "# and 时间<=#" + timenow + "# and 机种 = '" +
                    z_choice + "'")
                crsr.execute(
                    "SELECT * FROM 产量2 WHERE 时间>=#" + timelast + "# and 时间<=#" + timenow + "# and 机种 = '" +
                    z_choice + "'")
                list0 = crsr.fetchall()

                if list0 != []:
                    list_Data = []
                    list_Data_x = []
                    list_Data_x1 = []
                    draw_datas = []
                    for i in range(len(list0)):
                        list_Data.append(list0[i][3])
                        list_Data_x.append(list0[i][2])
                        list_Data_x1.append(i)

                    d = {'a': list_Data_x,
                         'b': list_Data}
                    df = pd.DataFrame(d)

                    df.drop_duplicates(subset=['a'], keep='first', inplace=True)
                    ts = pd.Series(df['b'].tolist(), index=df['a'])
                    print(ts)
                    ts_10T = ts.resample('1T').bfill()
                    z_timenow_H = (datetime.datetime.now() + datetime.timedelta(days=-29)).strftime("%H")
                    z_timenow = (datetime.datetime.now() + datetime.timedelta(days=-30)).strftime("%Y-%m-%d")
                    z_timelast = (datetime.datetime.now() + datetime.timedelta(days=-30)).strftime("%Y-%m-%d")

                    if int(z_timenow_H) >= 8 and int(z_timenow_H) <= 24:
                        ts_bai = ts_10T[z_timenow + " 8:00:00":z_timenow + " 21:00:00"]
                        ts_ye = ts_10T[z_timenow + " 21:00:00":z_timenow + " 23:59:59"]
                    else:
                        ts_bai = ts_10T[z_timelast + " 8:00:00":z_timelast + " 21:00:00"]
                        ts_ye = ts_10T[z_timelast + " 21:00:00":z_timenow + " 8:00:00"]
                    draw_datas.append(ts_bai)
                    draw_datas.append(ts_ye)
                    draw_datas.append(ts_10T)
                    wx.CallAfter(pub.sendMessage, "z_plt_Data2", msg=draw_datas)

        except:
            pass

        print("132")

class CheckListCtrl(wx.ListCtrl, CheckListCtrlMixin, ListCtrlAutoWidthMixin):
    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent, -1, style= wx.LC_EDIT_LABELS|wx.LC_HRULES|wx.LC_REPORT |wx.LC_VRULES)
        CheckListCtrlMixin.__init__(self)
        ListCtrlAutoWidthMixin.__init__(self)

class mainWin(MyFrame1):
    def __init__(self, parent):
        MyFrame1.__init__(self, parent)
        # self.listbox01.Append("123")
        # self.listbox01.SetString(0, "456")

        self.CSH()
        Refresh = threading.Thread(target=self.refresh)
        Refresh.start()
        #
        Send = threading.Thread(target=self.send)
        Send.start()
        #
        Hour = threading.Thread(target=self.hour)
        Hour.start()
        #
        Database = threading.Thread(target=self.InserDatabase)
        Database.start()

        self.bn1.Bind(wx.EVT_BUTTON, self.Exit)
        self.bn2.Bind(wx.EVT_BUTTON, self.clear)
        self.bn3.Bind(wx.EVT_BUTTON, self.Path)
        self.bn_scbb.Bind(wx.EVT_BUTTON, self.outputxls)
        self.tb_shuaxin.Bind(wx.EVT_BUTTON, self.refreshpic)

        global DataID, DataCheck, DataRec, D_Rank,z_choice,z_CL
        z_CL = 0

        self.Bind(wx.aui.EVT_AUINOTEBOOK_PAGE_CHANGED, self.OnSelChange)
        self.z_comboBox1.Bind(wx.EVT_COMBOBOX, self.OnCombo)
        pub.subscribe(self.draw, "plt_Data")
        pub.subscribe(self.z_draw, "z_plt_Data")
        pub.subscribe(self.z_draw2, "z_plt_Data2")
        global DataRec_test, DataRec_yesterday
        item = 0
        DataRec_test = [item]*50
        DataRec_yesterday = [item] * 50

        for D in range(len(DataID)):
            if DataCheck[D] == 2:
                DataRec_test[D] = (str(random.randint(2000,10000)))
                DataRec_yesterday[D] = 0

        global Data_15Day
        Data_15Day = [([0]*50) for i in range(50)]

        for D in range(len(DataID)):
            if DataCheck[D] == 2 and D_Rank[D] < 5:
                for E in range(15):
                    Data_15Day[D][E] = 0
            elif 10>= D_Rank[D] >= 5 :
                for E in range(15):
                    Data_15Day[D][E] = 0
            else:
                for E in range(15):
                    Data_15Day[D][E] = 0
        self.z_comboBox1.SetSelection(24)
        z_choice = self.z_comboBox1.GetValue()
        # self.z_draw1_start()
        # self.z_draw2_start()
        # self.draw1_start()
        # self.draw2_start()
        # self.datasta_yellow()

        time_start = time.time()  # 开始计时
        # TestThread5()
        print(1)
        # cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=./新建 Microsoft Access 数据库.accdb')
        # crsr = cnxn.cursor()
        # crsr = cnxn.cursor()
        # timenow = (datetime.datetime.now() + datetime.timedelta(days=-0)).strftime("%Y-%m-%d %H:%M:%S")
        # timelast = (datetime.datetime.now() + datetime.timedelta(days=-1)).strftime("%Y-%m-%d %H:%M:%S")
        # crsr.execute(
        #     "SELECT * FROM 产量2")
        # print(2)
        # list0 = crsr.fetchall()
        #
        # print(list0)
        # _db = pymysql.connect(host='localhost',
        #                       user='user1',
        #                       password='ruanjianjishu',
        #                       database='test')
        # # 使用 cursor() 方法创建一个游标对象 cursor
        # _cursor = _db.cursor()
        #
        # _cursor.execute(
        #     "SELECT * FROM yield_all WHERE name = '" +
        #                     "7302-2-FML" + "'")
        # list0 = _cursor.fetchall()
        # print(list0)
        # # 关闭数据库连接
        # _db.close()
        # loop = asyncio.get_event_loop()
        # loop.run_until_complete(self.test())
        # loop.close()
        # time_end = time.time()
        # time_c = time_end - time_start  # 运行所花时间
        #
        # print('time cost', time_c, 's')


    async def test(self):
        async with aio_sa.create_engine(host="localhost",
                                        port=3306,
                                        user="user1",
                                        password="ruanjianjishu",
                                        db="test",
                                        connect_timeout=10) as engine:
            async with engine.acquire() as conn:
                sql = "SELECT * FROM yield_all WHERE name = '" + "7302-2-FML" + "'"
                result = await conn.execute(sql)
                data = await result.fetchall()
                print(list(map(dict, data)))



    def handle_event(self, event):
        pass

    def CSH(self):
        self.tb_listbox.InsertColumn(0, '机种', width=150)
        self.tb_listbox.InsertColumn(1, '状态', wx.LIST_FORMAT_RIGHT, 50)
        self.tb_listbox.InsertColumn(2, '产量', wx.LIST_FORMAT_RIGHT, 50)
        # 入库队列
        global DatabaseQ
        DatabaseQ = queue.Queue()
        # 配置初始化
        global cf, sections, Output_path
        cf = configparser.ConfigParser()
        cf.read("Config.ini")
        sections = cf.sections()
        sections.remove("Setting")
        print(sections)
        global DataID, DataCheck, DataRec, D_Rank, last_Data,TestThread4_flag
        TestThread4_flag = 0
        DataID = sections
        item = 0
        DataCheck = [item] * 50
        item2 = 0
        last_Data = [item2] * 50
        DataRec =[(["0"] * 4) for i in range(50)]
        item3 = 0
        D_Rank = [item3] * 50
        D_flag = 0
        for D in range(len(sections)):
            DataCheck[D] = int(cf.get(sections[D], "data_ischeck"))
            if DataCheck[D] == 2:

                self.listbox01.Append(DataID[D])
                self.listbox02.Append(cf.get(sections[D], "name"))
                index = self.tb_listbox.InsertItem(D_flag, DataID[D])
                self.tb_listbox.SetItem(D_flag, 1, "正常")
                self.tb_listbox.SetItemBackgroundColour(D_flag, "Green")
                # self.tb_listbox.Set(1, 1, "Yellow");
                self.tb_listbox.SetItem(D_flag, 2, "1")
                D_Rank[D] = D_flag
                D_flag += 1
        self.tb_listbox.SetItemBackgroundColour(1, "Yellow")
        self.tb_listbox.SetItem(1, 1, "无信号")
        self.tb_listbox.SetItemBackgroundColour(2, "Red")
        self.tb_listbox.SetItem(2, 1, "停机")

        global Port, DelayTime, SunnyLink, SunnyLink_IsCheck
        Port = cf.get("Setting", "Port")
        DelayTime = cf.get("Setting", "Time")
        SunnyLink = cf.get("Setting", "SunnyLink")
        SunnyLink_IsCheck = cf.get("Setting", "SunnyLink_IsCheck")


        # 设置中文字体
        matplotlib.rc("font", family='FangSong')
        # 遍历句柄
        hwnd_title = dict()

        def get_all_hwnd(hwnd, mouse):
            if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd) and win32gui.IsWindowVisible(hwnd):
                hwnd_title.update({hwnd: win32gui.GetWindowText(hwnd)})

        win32gui.EnumWindows(get_all_hwnd, 0)
        win32gui.EnumWindows(get_all_hwnd, 0)
        # hd为存放句柄的数组，hh为计数符号
        hd = [0]
        hh = 0
        # 筛选出所有窗口，将其打印出来，并存放到数组中
        global hwnd_SunnyLink
        hwnd_SunnyLink = win32gui.FindWindow(0, "SunnyLink")
        for h, t in hwnd_title.items():
            hwnd_Mag = str(h) + " : " + str(t)
            print(hwnd_Mag)
            if t == "SunnyLink":
                hwnd_SunnyLink = h

        # 根据titlename信息查找窗口
        print("SunnyLink句柄：" + str(hwnd_SunnyLink))

        # 数据库连接
        global cnxn, crsr
        cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=./DataSend03.accdb')
        crsr = cnxn.cursor()

        # 绑定一个UDP端口
        global udpSocket
        udpSocket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        bindAdress = ('', 3007)
        udpSocket.bind(bindAdress)  # 绑定一个端口
        # 全局变量
        global recvMsg, Hour_Msg
        recvMsg = ""
        Hour_Msg = ""

        self.path.SetValue(cf.get("Setting", "path"))
        Output_path = cf.get("Setting", "path")

        # 绘图锁
        global StopFlag_1
        StopFlag_1 = 0

        # 选中全部线体
        num = self.tb_listbox.GetItemCount()
        for i in range(num):
            self.tb_listbox.CheckItem(i)

        # 子界面选项
        for D in range(len(sections)):
            DataCheck[D] = int(cf.get(sections[D], "data_ischeck"))
            if DataCheck[D] == 2:
                self.z_comboBox1.Append(DataID[D])
        self.z_comboBox1.SetSelection(0)
        self.z_listbox.InsertColumn(0, '项目', width=150)
        self.z_listbox.InsertColumn(1, '数据', 50)

        self.z_listbox.InsertItem(0, "线体名称：")
        self.z_listbox.InsertItem(1, "线体位置：")
        self.z_listbox.InsertItem(2, "线体状态：")
        self.z_listbox.InsertItem(3, "实时产量：")
        self.z_listbox.InsertItem(4, "实时稼动率：")
        self.z_listbox.InsertItem(5, "实时CT：")
        self.z_listbox.InsertItem(6, "停机时间：")
        self.z_listbox.InsertItem(7, "昨日产量：")
        self.z_listbox.InsertItem(8, "昨日稼动率：")
        self.z_listbox.InsertItem(9, "昨日CT：")
        self.z_listbox.InsertItem(10, "昨日停机时间：")

        self.z_TJlistbox.InsertColumn(0, '日期', width=100)
        self.z_TJlistbox.InsertColumn(1, '设备编号', width=100)
        self.z_TJlistbox.InsertColumn(2, '机种', width=100)
        self.z_TJlistbox.InsertColumn(3, '停机开始时间', width=200)
        self.z_TJlistbox.InsertColumn(4, '停机结束时间', width=200)
        self.z_TJlistbox.InsertColumn(5, '停机时长', width=50)


    def SendMsg_SunnyLink(self, hwnd):
            # 获取左上和右下的坐标
            left, top, right, bottom = win32gui.GetClientRect(hwnd)
            print(left, top, right, bottom)
            # 激活选中SunnyLink窗口

            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            long_position = win32api.MAKELONG(int(300), int(bottom - 50))

            long_position2 = win32api.MAKELONG(int(828), int(bottom - 24))
            long_position3 = win32api.MAKELONG(int(310), int(bottom - 50))
            # 点击
            win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position)
            win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position)
            time.sleep(0.5)
            win32api.PostMessage(hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, long_position)
            win32api.PostMessage(hwnd, win32con.WM_RBUTTONUP, win32con.MK_RBUTTON, long_position)
            win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position3)
            win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position3)
            time.sleep(0.5)
            win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position2)
            win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position2)

    # 保持不进入休眠
    def NOSleep(self, hwnd):
            # 获取左上和右下的坐标
            # left, top, right, bottom = win32gui.GetClientRect(hwnd)
            # print(left, top, right, bottom)
            # 激活选中微信窗口
            # win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            # long_position = win32api.MAKELONG(int(300), int(bottom - 50))
            # 点击左键
            # win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position)
            # win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position)
            try:
                win32gui.SendMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
                time.sleep(5)
                win32api.keybd_event(145, win32api.MapVirtualKey(145, 0), 0, 0)
                win32api.keybd_event(145, win32api.MapVirtualKey(145, 0), win32con.KEYEVENTF_KEYUP, 0)
                win32api.keybd_event(145, win32api.MapVirtualKey(145, 0), 0, 0)
                win32api.keybd_event(145, win32api.MapVirtualKey(145, 0), win32con.KEYEVENTF_KEYUP, 0)
                time.sleep(0.5)

            except:
                pass

    # 存入数据库
    def SaveData(self, data01, data02, data03, data04):
        global cnxn, crsr
        # cnxn = pyodbc.connect(
        #     r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=./DataSend03.accdb')
        # crsr = cnxn.cursor()
        try:
            crsr.execute(
                "INSERT INTO 产量(机种,日期,时间,产量)VALUES(" + "'" + data01 + "'" + "," + data02 + "," + data03 + "," + data04 + ")")
            crsr.commit()
            # crsr.close()
            # cnxn.close()
        except Exception as e:
            print(e)

    def SaveData2(self, data01, data02, data03, data04):
        global cnxn, crsr
        # cnxn = pyodbc.connect(
        #     r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=./DataSend03.accdb')
        # crsr = cnxn.cursor()
        try:
            crsr.execute(
                "INSERT INTO 产量2(机种,日期,时间,产量)VALUES(" + "'" + data01 + "'" + "," + data02 + "," + data03 + "," + data04 + ")")
            crsr.commit()
            # crsr.close()
            # cnxn.close()
        except Exception as e:
            print(e)

    def refresh(self):
        global Send_Data, recvMsg, udpSocket, DatabaseQ, restartflag01
        global last_Data,sta_flag,Recflag
        restartflag01 = 1
        item = 0
        sta_flag = [item] * 50
        while True:
            if restartflag01 == 1:
                try:
                    recvDate,recvAddr = udpSocket.recvfrom(1024)#如果没有收到发往这个绑定端口的消息，会一直阻塞在这里
                    Send_Data = recvDate.decode('gbk')
                    # recvMsg = '【Receive from %s : %s】：%s'%(recvAddr[0],recvAddr[1],recvDate.decode('gbk'))

                    DatabaseQ.put(recvDate)
                    _msg = recvDate.decode('gbk').split("：")
                    global DataID, DataCheck, DataRec
                    for D in range(len(DataID)):

                        if DataCheck[D] == 2 and _msg[0] == DataID[D]:
                            Recflag[D] = 1
                            self.tb_listbox.SetItem(D_Rank[D], 2, str(_msg[2]))
                            if last_Data[D] != _msg[2]:
                                self.tb_listbox.SetItemBackgroundColour(D_Rank[D], "Green")
                                self.tb_listbox.SetItem(D_Rank[D], 1, "正常")
                                last_Data[D] = _msg[2]
                            else:
                                sta_flag[D] += 1
                                if sta_flag[D] >= 6:
                                    sta_flag[D] = 0
                                    self.tb_listbox.SetItemBackgroundColour(D_Rank[D], "Red")
                                    self.tb_listbox.SetItem(D_Rank[D], 1, "停线")


                    self.Qnum.SetValue("队列个数：" + str(DatabaseQ.qsize()))
                except Exception as e:
                    print(e)
            else:
                print("refresh重启")
                break

        time.sleep(5)
        Refresh = threading.Thread(target=self.refresh)
        Refresh.start()

    def InserDatabase(self):
        global Send_Data, recvMsg, udpSocket, DatabaseQ, restartflag02,DataRec,D_Rank,StopFlag, InserDatabaseflag,Recflag
        restartflag02 = 1
        item = 0
        InserDatabaseflag = [item] * 50
        item1 = 0
        Recflag = [item1] * 50
        while True:
            if restartflag02 == 1:
                try:
                    time.sleep(0.1)

                    Hour_now = datetime.datetime.now().strftime("%M")
                    if not DatabaseQ.empty() and Hour_now != "02" and Hour_now != "03" and Hour_now != "01" and   Hour_now != "04"and   Hour_now != "05" and StopFlag == 0 and StopFlag_1 == 0:
                        recvDate = DatabaseQ.get()
                        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        recvMsg = now + " " + recvDate.decode('gbk')
                        print(recvMsg)
                        date = "'" + datetime.datetime.now().strftime("%Y-%m-%d") + "'"
                        timenow = "'" + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "'"
                        LensNums = recvDate.decode('gbk').split("：")
                        if len(LensNums) == 3:
                            # print(LensNums)
                            self.SaveData(LensNums[0], date, timenow, str(int(LensNums[2])))

                            global DataID, DataCheck, DataRec
                            for D in range(len(DataID)):
                                if LensNums[0] == DataID[D]:
                                    InserDatabaseflag[D] += 1

                                    DataRec[D] = LensNums
                                    if DataCheck[D] == 2:
                                        if InserDatabaseflag[D] >= 10:
                                            InserDatabaseflag[D] = 0
                                            self.SaveData2(LensNums[0], date, timenow, str(int(LensNums[2])))
                                        self.listbox01.SetString(D_Rank[D], str(LensNums[0]) + ": " + now + ": " + str(LensNums[2]))

                    time.sleep(0.1)
                except Exception as e:
                    print(e)
            else:
                print("InserDatabase重启")
                break
        time.sleep(5)
        Database = threading.Thread(target=self.InserDatabase)
        Database.start()

    # 发送
    def send(self):
            delayTimes = 0
            global recvMsg, hwnd_SunnyLink, restartflag03
            global DataID, DataCheck, DataRec
            restartflag03 = 1
            while True:
                if restartflag03 == 1:
                    try:
                        Text = ""
                        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        # msg = DataBaseSelectTime("2020-11-5 20:00:00" , "2020-11-5 21:00:00")
                        time.sleep(60)
                        self.NOSleep(hwnd_SunnyLink)
                        Send_Msg = "时间: " + now  + ": "+  "\n"

                        for D in range(len(DataID)):
                            if DataCheck[D] == 2:
                                try:
                                    Send_Msg = Send_Msg + "     " + DataID[D] + " : 投入： " + DataRec[D][2] + "\n"
                                except Exception as e:
                                    print(e)

                            time.sleep(0.1)
                        delayTimes += 1
                        Hour_now = datetime.datetime.now().strftime("%M")

                        if (delayTimes >= 10 and  Hour_now != "59" and Hour_now != "00" and Hour_now != "01"and Hour_now != "02"and Hour_now != "03"and Hour_now != "04"and Hour_now != "05") :
                            print(Send_Msg)
                            Text = pyperclip.paste()
                            # Send_Msg = Text + Send_Msg
                            pyperclip.copy(Send_Msg)
                            self.label01.SetValue(Send_Msg)
                            self.SendMsg_SunnyLink(hwnd_SunnyLink)
                            pyperclip.copy(Text)
                            delayTimes = 0
                    except:
                        delayTimes = 0
                else:
                    print("Send重启")
                    break
            Send = threading.Thread(target=self.send)
            Send.start()

    # 整点报时
    def DataBaseSelectTime(self, Time01, Time02):
        global cnxn, crsr
        global DataID, DataCheck, DataRec
        # cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=./DataSend03.accdb')
        # crsr = cnxn.cursor()
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        Send_Msg = "整点报时：" + Time01  + "~" + Time02+ "\n"
        try:
            for D in range(len(DataID)):
                if DataCheck[D] == 2:
                    try:
                        crsr.execute("SELECT * FROM 产量2 WHERE 时间>=#" + Time01 + "# and 时间<=#" + Time02 + "# and 机种 = '" + DataID[D] +"'")
                        list = crsr.fetchall()
                        list0 = numpy.array(list)
                        LastDay_Data = list0[:, 3]
                        cl_data = 0
                        for Data_Rank in range(len(LastDay_Data)):
                            if Data_Rank > 0 and int(LastDay_Data[Data_Rank]) > int(LastDay_Data[Data_Rank - 1]):
                                cl_data = cl_data + (int(LastDay_Data[Data_Rank]) - int(LastDay_Data[Data_Rank - 1]))
                        Renum = 0
                        for i in range(len(list0)):
                            if i > 0:
                                if list0[i][3] == list0[i - 1][3]:
                                    Renum += 1
                        # # 交接班
                        # if int(list0[-1][3]) >= int(list0[0][3]):
                        #     cl_data = str(int(list0[-1][3]) - int(list0[0][3]))
                        # else:
                        #     list0 = numpy.array(list0)
                        #     # print(max(map(int,list0[:, 3])), int(list0[0][3]), int(list0[-1][3]))
                        #     cl_data = str(max(map(int,list0[:, 3])) - int(list0[0][3]) + int(list0[-1][3]) )
                        #
                        if cl_data > 0:
                            Send_Msg = Send_Msg + DataID[D] + " : 产量：" + str(cl_data) + "\n" + "           停机时长：" + str(Renum) + "分钟" + "\n" + "           CT时间：" + str(
                                round(3600 / (int(cl_data)), 1)) + " S/pcs\n"
                        else:
                            Send_Msg = Send_Msg + DataID[D] + " : 数据异常\n"
                    except:
                        Send_Msg = Send_Msg + DataID[D] +" : 数据异常\n"

            # crsr.close()
            # cnxn.close()
            return Send_Msg

        except Exception as e:
            print(e)
            return "整点报时异常"

    def DataBaseSelectTime2(self):
        TestThread()

    def fileSend(self, filename):
        try:
            global hwnd_SunnyLink
            hwnd_SunnyLink2 = hwnd_SunnyLink
            left, top, right, bottom = win32gui.GetClientRect(hwnd_SunnyLink2)
            # print(left, top, right, bottom)
            # print(hwnd_SunnyLink2)
            win32gui.ShowWindow(hwnd_SunnyLink2, win32con.SW_SHOW)
            long_position0 = win32api.MAKELONG(int(300), int(bottom - 50))
            long_position = win32api.MAKELONG(int(828), int(bottom - 24))
            long_position2 = win32api.MAKELONG(int(375), int(445))
            # 点击
            win32api.PostMessage(hwnd_SunnyLink2, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position0)
            win32api.PostMessage(hwnd_SunnyLink2, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position0)
            time.sleep(0.5)
            win32api.PostMessage(hwnd_SunnyLink2, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position2)
            win32api.PostMessage(hwnd_SunnyLink2, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position2)

            time.sleep(10)
            hwnd_Openfile = win32gui.FindWindow(0, "打开")
            # 获取文件名输入框

            a1 = win32gui.FindWindowEx(hwnd_Openfile, None, "ComboBoxEx32", None)

            a2 = win32gui.FindWindowEx(a1, None, "ComboBox", None)

            hwnd_filename = win32gui.FindWindowEx(a2, None, "Edit", None)

            win32gui.SendMessage(hwnd_filename, win32con.WM_SETTEXT, None, filename)

            # 在文件名输入框中输入文件名

            hwnd_save = win32gui.FindWindowEx(hwnd_Openfile, None, "Button", None)

            win32gui.PostMessage(hwnd_save, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)

            win32gui.PostMessage(hwnd_save, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

            time.sleep(1)
            win32api.PostMessage(hwnd_SunnyLink2, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position)
            win32api.PostMessage(hwnd_SunnyLink2, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position)
        except:
            pass

    # 整点报时
    def hour(self):
        global DataID, DataCheck, DataRec
        global Hour_Msg, restartflag04, StopFlag
        StopFlag = 0
        restartflag04 = 1
        while True:
            if restartflag04 == 1:
                Hour_now = datetime.datetime.now().strftime("%M%S")
                Hour_now2 = datetime.datetime.now().strftime("%H")
                if Hour_now == "0100":
                    try:
                        if Hour_now2 == "09":
                            try:
                                StopFlag = 1
                                self.DataBaseSelectTime2()
                                time.sleep(30)
                                now = datetime.datetime.now().strftime("%Y-%m-%d")
                                self.fileSend(self.path.GetValue() + "\\"+"流水线生产数据报表 " + now + ".xls")
                                StopFlag = 0
                            except Exception as e:
                                StopFlag = 0
                                print(e)
                        else:
                            StopFlag = 1
                            # print(int(Hour_now2))
                            if int(Hour_now2) > 9 and int(Hour_now2) < 24:
                                now = datetime.datetime.now().strftime("%Y-%m-%d")
                                self.fileSend(self.path.GetValue() + "\\"+"流水线生产数据报表 " + now + ".xls")
                                time.sleep(10)
                            timenow = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            timelast = (datetime.datetime.now() + datetime.timedelta(hours=-1)).strftime("%Y-%m-%d %H:%M:%S")
                            # print(timenow)
                            # print(timelast)
                            # msg = DataBaseSelectTime("2020-11-5 20:00:00", "2020-11-5 21:00:00")
                            msg = self.DataBaseSelectTime(str(timelast), str(timenow))
                            print(msg)
                            pyperclip.copy(msg)
                            self.label01.SetValue(msg)
                            self.SendMsg_SunnyLink(hwnd_SunnyLink)
                            time.sleep(10)
                            StopFlag = 0
                        time.sleep(30)
                    except Exception as e:
                        StopFlag = 0
                        print(e)
                # if Hour_now == "1200":
                #     try:
                #         msg = DataBaseSelectTime2()
                #
                #         time.sleep(10)
                #     except Exception as e:
                #         print(e)
            else:
                print("Hour重启")
                break

        Hour = threading.Thread(target= self.hour)
        Hour.start()

    def Exit(self,event):
        try:
            global cnxn, crsr
            crsr.close()
            cnxn.close()
        except:
            pass
        os._exit(0)

    def clear(self,event):
        # 配置初始化
        global cf, sections

        global DataID, DataCheck, DataRec, D_Rank

        for D in range(len(sections)):
            if DataCheck[D] == 2:
                self.listbox01.SetString(D_Rank[D],DataID[D])

    def Path(self,event):
        filesFilter = "输出文档 (*.txt)|*.txt|" "All files (*.*)|*.*"
        fileDialog = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        dialogResult = fileDialog.ShowModal()
        if dialogResult != wx.ID_OK:
            return
        path = fileDialog.GetPath()
        self.path.SetValue(path)
        global cf, sections
        cf.set("Setting", "path", self.path.GetValue())
        cf.write(open("Config.ini", "w"))

    def outputxls(self,event):
        try:
            global StopFlag
            self.DataBaseSelectTime2()
            now = datetime.datetime.now().strftime("%Y-%m-%d")
            # self.fileSend(self.path.GetValue() + "\\" + "流水线生产数据报表 " + now + ".xls")
        except Exception as e:
            print(e)
    # 数据页面
    def OnSelChange(self, event):
        event.Skip()
        if self.auinotebook.GetSelection() <= 1 :
            self.auinotebook.SetWindowStyle(wx.aui.AUI_NB_SCROLL_BUTTONS)
        else:
            self.auinotebook.SetWindowStyle(wx.aui.AUI_NB_CLOSE_ON_ACTIVE_TAB)

    # 绘图
    def draw(self, msg):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        global DataID, DataCheck, DataRec, D_Rank
        global DataRec_test, DataRec_yesterday
        self.axes.clear()
        for D in range(len(DataID)):
            if DataCheck[D] == 2 and self.tb_listbox.IsChecked(D_Rank[D]):
                self.axes.plot(msg[D_Rank[D]], label=DataID[D])
        self.axes.set_title("一小时产量变化推移图 更新时间："+ now)
        self.axes.legend(loc="upper right",ncol=2)

        self.axes.grid()
        self.figure.autofmt_xdate()
        self.figure.tight_layout()
        self.figure.set_canvas(self.canvas)
        self.canvas.draw()

    def draw2(self):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        global DataID, DataCheck, DataRec, D_Rank
        global DataRec_test, DataRec_yesterday
        self.axes1.clear()
        axes1_bar_x = []
        axes1_bar_y = []
        axes1_bar_b = []
        for D in range(len(DataID)):
            if DataCheck[D] == 2 and self.tb_listbox.IsChecked(D_Rank[D]):
                axes1_bar_x.append(DataID[D])
                axes1_bar_y.append(int(DataRec[D][2]))
                axes1_bar_b.append(DataRec_yesterday[D])
        self.axes1.bar(axes1_bar_x, axes1_bar_y, label='今日产量')
        self.axes1.bar(axes1_bar_x, axes1_bar_b, bottom=axes1_bar_y, label='昨日产量')

        self.axes1.set_title("实时产量总量图 更新时间："+ now)

        self.axes1.legend(loc="upper right",ncol=2)
        self.axes1.set_xticklabels(axes1_bar_x, fontsize=8)
        self.axes1.grid()
        self.figure1.autofmt_xdate()
        self.figure1.tight_layout()
        self.figure1.set_canvas(self.canvas1)

        self.canvas1.draw()

        global Data_15Day
        self.axes2.clear()
        for D in range(len(DataID)):
            if DataCheck[D] == 2 and self.tb_listbox.IsChecked(D_Rank[D]):
                self.axes2.plot(Data_15Day[D], label=DataID[D])
        self.axes2.set_title("15天日产量推移图 更新时间:"+ now)
        self.axes2.grid()
        self.axes2.legend(loc="upper right",ncol=2)
        self.figure2.tight_layout()
        self.figure2.set_canvas(self.canvas2)
        self.canvas2.draw()

        self.axes3.clear()
        axes1_bar_x = []
        axes1_bar_b = []
        for D in range(len(DataID)):
            if DataCheck[D] == 2 and self.tb_listbox.IsChecked(D_Rank[D]):
                axes1_bar_x.append(DataID[D])
                axes1_bar_b.append(DataRec_yesterday[D])
        labels = axes1_bar_x  # 定义标签
        sizes = axes1_bar_b  # 每块值

        patches, text1, text2 = self.axes3.pie(sizes,
                                               labels=labels,
                                               autopct='%3.2f%%',  # 数值保留固定小数位
                                               shadow=False,  # 无阴影设置
                                               startangle=90,  # 逆时针起始角度设置
                                               pctdistance=0.6)  # 数值距圆心半径倍数距离
        # patches饼图的返回值，texts1饼图外label的文本，texts2饼图内部的文本
        # x，y轴刻度设置一致，保证饼图为圆形
        self.axes3.set_title("昨日产量比例分布 更新时间:"+ now)
        self.axes3.legend(loc="upper right",ncol=2)
        self.figure3.tight_layout()
        self.axes3.axis('equal')
        self.figure3.set_canvas(self.canvas3)
        self.canvas3.draw()

    def draw1_start(self):
        global timer
        try:
            TestThread2()
            timer = threading.Timer(300,self.draw1_start)
            timer.start()
        except:
            timer.cancel()
            timer.start()

    def draw2_start(self):
        global timer2
        try:
            self.draw2()
            timer2 = threading.Timer(60, self.draw2_start)
            timer2.start()
        except:
            timer2.cancel()
            timer2.start()

    def datasta_yellow(self):
        global timer3
        try:
            global DataID, DataCheck, DataRec,Recflag
            for D in range(len(DataID)):
                if DataCheck[D] == 2 and Recflag[D] == 0:
                    self.tb_listbox.SetItemBackgroundColour(D_Rank[D], "Yellow")
                    self.tb_listbox.SetItem(D_Rank[D], 1, "无信号")
                else:
                    Recflag[D] = 0

            timer3 = threading.Timer(60, self.datasta_yellow)
            timer3.start()
        except:
            timer3.cancel()
            timer3.start()

    def refreshpic(self,event):
        self.draw2()
        TestThread2()

    def z_draw(self, msg):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        global DataID, DataCheck, DataRec, D_Rank, z_choice
        global DataRec_test, DataRec_yesterday
        self.z_axes.clear()

        self.z_axes.plot(msg, label=z_choice)
        self.z_axes.set_title("一小时产量变化推移图 更新时间："+ now)
        self.z_axes.legend(loc="upper right",ncol=2)

        self.z_axes.grid()
        self.z_figure.autofmt_xdate()
        self.z_figure.tight_layout()
        self.z_figure.set_canvas(self.z_canvas)
        self.z_canvas.draw()

    def z_draw2(self, msg):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        global DataID, DataCheck, DataRec, D_Rank, z_choice
        global DataRec_test, DataRec_yesterday
        self.z_axes1.clear()

        self.z_axes1.plot(msg[0], label=z_choice+": 白班")
        self.z_axes1.plot(msg[1], label=z_choice+": 夜班")
        self.z_axes1.set_title("白夜班产量变化图 更新时间：" + now)
        self.z_axes1.legend(loc="upper right", ncol=1)

        self.z_axes1.grid()
        self.z_figure1.autofmt_xdate()
        self.z_figure1.tight_layout()
        self.z_figure1.set_canvas(self.z_canvas1)
        self.z_canvas1.draw()

    def z_draw1_start(self):
        global z_timer
        try:
            TestThread4()
            z_timer = threading.Timer(60,self.z_draw1_start)
            z_timer.start()
        except:
            z_timer.cancel()
            z_timer.start()

    def z_draw2_start(self):
        global z_timer2
        try:
            TestThread5()
            z_timer2 = threading.Timer(180, self.z_draw2_start)
            z_timer2.start()
        except:
            z_timer2.cancel()
            z_timer2.start()


    def OnCombo(self,event):
        global z_choice,TestThread4_flag
        z_choice = self.z_comboBox1.GetValue()
        self.z_axes.clear()
        self.z_axes1.clear()
        self.z_axes2.clear()

        self.z_figure.set_canvas(self.z_canvas)
        self.z_canvas.draw()
        self.z_figure1.set_canvas(self.z_canvas1)
        self.z_canvas1.draw()
        self.z_figure2.set_canvas(self.z_canvas2)
        self.z_canvas2.draw()
        TestThread4()

if __name__ == '__main__':

    app = wx.App()
    main_win = mainWin(None)
    main_win.Show()
    app.MainLoop()