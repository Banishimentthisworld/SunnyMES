import configparser
import datetime
import os
import queue
import socket
import time

import matplotlib
import win32api
import win32con
import win32gui
import wx
import wx.aui

from GUI import MyFrame1
# 主线程
class mainWin(MyFrame1):
    def __init__(self, parent):
        MyFrame1.__init__(self, parent)

        # 界面绑定
        # 退出按钮
        self.bn1.Bind(wx.EVT_BUTTON, self.Exit)
        # 清空当前按钮
        self.bn2.Bind(wx.EVT_BUTTON, self.clear)
        # 报表输出路径按钮
        self.bn3.Bind(wx.EVT_BUTTON, self.Path)
        # 报表手动输出按钮
        self.bn_scbb.Bind(wx.EVT_BUTTON, self.outputxls)
        # 刷新按钮
        self.tb_shuaxin.Bind(wx.EVT_BUTTON, self.refreshpic)
        # 标签页切换事件
        self.Bind(wx.aui.EVT_AUINOTEBOOK_PAGE_CHANGED, self.OnSelChange)
        # 线体明细线体切换下拉框
        self.z_comboBox1.Bind(wx.EVT_COMBOBOX, self.OnCombo)

        # 汇总页面表格初始化
        self.tb_listbox.InsertColumn(0, '机种', width=150)
        self.tb_listbox.InsertColumn(1, '状态', wx.LIST_FORMAT_RIGHT, 50)
        self.tb_listbox.InsertColumn(2, '产量', wx.LIST_FORMAT_RIGHT, 50)

        # 子页面初始化
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

        # 线程响应函数

        # 参数初始化
        self.CSH()

    # 参数初始化
    def CSH(self):
        # 入库队列
        global DatabaseQ
        DatabaseQ = queue.Queue()

        # 配置初始化
        global cf, sections
        cf = configparser.ConfigParser()
        cf.read("Config.ini")
        sections = cf.sections()
        sections.remove("Setting")

        # 机种信息初始化，DataID为机种编号，DataCheck为是否启用，0为不启用，2为启用，DataRec为当前机种产量值，D_Rank为启用编号，last_Data为上次机种产量值
        global DataID, DataCheck, DataRec, D_Rank, last_Data
        DataID = sections
        item = 0
        DataCheck = [item] * 50
        item2 = 0
        last_Data = [item2] * 50
        DataRec = [(["0"] * 4) for i in range(50)]
        item3 = 0
        D_Rank = [item3] * 50

        # 初始化汇总界面左侧实时信息栏
        D_flag = 0
        for D in range(len(sections)):
            DataCheck[D] = int(cf.get(sections[D], "data_ischeck"))
            if DataCheck[D] == 2:
                self.listbox01.Append(DataID[D])
                self.listbox02.Append(cf.get(sections[D], "name"))
                self.tb_listbox.InsertItem(D_flag, DataID[D])
                self.tb_listbox.SetItem(D_flag, 1, "正常")
                self.tb_listbox.SetItemBackgroundColour(D_flag, "Green")
                self.tb_listbox.SetItem(D_flag, 2, "1")
                D_Rank[D] = D_flag
                D_flag += 1

        # 初始化推送参数
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

        # 绑定一个UDP端口
        global udpSocket
        udpSocket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        bindAdress = ('', 3007)
        udpSocket.bind(bindAdress)  # 绑定一个端口

        # 初始化接收信息
        global recvMsg, Hour_Msg
        recvMsg = ""
        Hour_Msg = ""

        # 初始化数据库输出路径
        global Output_path
        self.path.SetValue(cf.get("Setting", "path"))
        Output_path = cf.get("Setting", "path")

        # 选中全部线体
        num = self.tb_listbox.GetItemCount()
        for i in range(num):
            self.tb_listbox.CheckItem(i)

        # 子界面选项
        for D in range(len(sections)):
            DataCheck[D] = int(cf.get(sections[D], "data_ischeck"))
            if DataCheck[D] == 2:
                self.z_comboBox1.Append(DataID[D])

        # 初始化子界面线体选项
        global z_choice,z_CL
        z_CL = 0
        self.z_comboBox1.SetSelection(0)
        z_choice = self.z_comboBox1.GetValue()

    # 退出程序
    def Exit(self,event):
        try:
            global cnxn, crsr
            crsr.close()
            cnxn.close()
        except:
            pass
        os._exit(0)

    # 清空原始数据界面显示
    def clear(self,event):
        # 配置初始化
        global cf, sections
        global DataID, DataCheck, DataRec, D_Rank
        for D in range(len(sections)):
            if DataCheck[D] == 2:
                self.listbox01.SetString(D_Rank[D],DataID[D])

    # 报表输出路径
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

    # 输出报表
    def outputxls(self,event):
        try:
            global StopFlag
            self.DataBaseSelectTime2()
            now = datetime.datetime.now().strftime("%Y-%m-%d")
            # self.fileSend(self.path.GetValue() + "\\" + "流水线生产数据报表 " + now + ".xls")
        except Exception as e:
            print(e)

    # 汇总页面刷新图表
    def refreshpic(self,event):
        self.draw2()
        TestThread2()

    # 标签页切换事件
    def OnSelChange(self, event):
        event.Skip()
        if self.auinotebook.GetSelection() <= 2 :
            self.auinotebook.SetWindowStyle(wx.aui.AUI_NB_SCROLL_BUTTONS)
        else:
            self.auinotebook.SetWindowStyle(wx.aui.AUI_NB_CLOSE_ON_ACTIVE_TAB)

    # 线体明细下拉框响应事件
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

    # 保持不进入休眠
    def NOSleep(self, hwnd):
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
if __name__ == '__main__':

    app = wx.App()
    main_win = mainWin(None)
    main_win.Show()
    app.MainLoop()