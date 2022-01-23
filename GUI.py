# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Oct 26 2018)
## http://www.wxformbuilder.org/
##
## PLEASE DO *NOT* EDIT THIS FILE!
###########################################################################
import numpy
import wx
import wx.xrc
import wx.aui
import wx.adv
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg
from matplotlib.figure import Figure
from wx.lib.mixins.listctrl import CheckListCtrlMixin, ListCtrlAutoWidthMixin


class CheckListCtrl(wx.ListCtrl, CheckListCtrlMixin, ListCtrlAutoWidthMixin):
    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent, -1, style=wx.LC_EDIT_LABELS | wx.LC_HRULES | wx.LC_REPORT | wx.LC_VRULES)
        CheckListCtrlMixin.__init__(self)
        ListCtrlAutoWidthMixin.__init__(self)

class ListCtrlAutoCtrl(wx.ListCtrl, CheckListCtrlMixin, ListCtrlAutoWidthMixin):
    def __init__(self, parent):
        wx.ListCtrl.__init__(self, parent, -1, style=wx.LC_EDIT_LABELS | wx.LC_HRULES | wx.LC_REPORT | wx.LC_VRULES)
        ListCtrlAutoWidthMixin.__init__(self)


###########################################################################
## Class MyFrame1
###########################################################################

class MyFrame1(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"生产数据", pos=wx.DefaultPosition, size=wx.Size(693, 580),
                          style=wx.MAXIMIZE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        bSizer8 = wx.BoxSizer(wx.HORIZONTAL)

        self.auinotebook = wx.aui.AuiNotebook(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                              wx.aui.AUI_NB_SCROLL_BUTTONS)
        self.pan_yssj = wx.Panel(self.auinotebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer6 = wx.BoxSizer(wx.HORIZONTAL)

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        bSizer4 = wx.BoxSizer(wx.HORIZONTAL)

        self.m_staticText2 = wx.StaticText(self.pan_yssj, wx.ID_ANY, u"实时数据：", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText2.Wrap(-1)

        bSizer4.Add(self.m_staticText2, 0, wx.ALL, 5)

        self.path = wx.TextCtrl(self.pan_yssj, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        self.path.SetFont(
            wx.Font(wx.NORMAL_FONT.GetPointSize(), wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL,
                    False, wx.EmptyString))

        bSizer4.Add(self.path, 0, wx.ALL, 5)

        self.bn3 = wx.Button(self.pan_yssj, wx.ID_ANY, u"报表路径", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer4.Add(self.bn3, 0, wx.ALL, 5)

        self.bn2 = wx.Button(self.pan_yssj, wx.ID_ANY, u"清空当前", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer4.Add(self.bn2, 0, wx.ALL, 5)

        self.bn_scbb = wx.Button(self.pan_yssj, wx.ID_ANY, u"生成报表", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer4.Add(self.bn_scbb, 0, wx.ALL, 5)

        self.Qnum = wx.TextCtrl(self.pan_yssj, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer4.Add(self.Qnum, 0, wx.ALL, 5)

        bSizer2.Add(bSizer4, 0, wx.EXPAND, 5)

        bSizer5 = wx.BoxSizer(wx.HORIZONTAL)

        listbox02Choices = []
        self.listbox02 = wx.ListBox(self.pan_yssj, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, listbox02Choices, 0)
        self.listbox02.SetFont(
            wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))

        bSizer5.Add(self.listbox02, 0, wx.ALIGN_RIGHT | wx.ALL | wx.EXPAND, 5)

        listbox01Choices = []
        self.listbox01 = wx.ListBox(self.pan_yssj, wx.ID_ANY, wx.DefaultPosition, wx.Size(250, -1), listbox01Choices, 0)
        self.listbox01.SetFont(
            wx.Font(20, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))

        bSizer5.Add(self.listbox01, 1, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)

        bSizer2.Add(bSizer5, 1, wx.EXPAND, 5)

        bSizer6.Add(bSizer2, 1, wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        self.bn1 = wx.Button(self.pan_yssj, wx.ID_ANY, u"退出", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer3.Add(self.bn1, 0, wx.ALL | wx.EXPAND, 5)

        self.label01 = wx.TextCtrl(self.pan_yssj, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                   wx.TE_MULTILINE | wx.TE_READONLY)
        self.label01.SetFont(
            wx.Font(15, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))

        bSizer3.Add(self.label01, 1, wx.ALL | wx.EXPAND, 5)

        bSizer6.Add(bSizer3, 1, wx.ALIGN_CENTER_VERTICAL | wx.EXPAND, 5)

        self.pan_yssj.SetSizer(bSizer6)
        self.pan_yssj.Layout()
        bSizer6.Fit(self.pan_yssj)
        self.auinotebook.AddPage(self.pan_yssj, u"原始数据", False, wx.NullBitmap)
        self.pan_tb = wx.Panel(self.auinotebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer81 = wx.BoxSizer(wx.HORIZONTAL)

        tb_left = wx.BoxSizer(wx.VERTICAL)

        # self.tb_listbox = wx.ListCtrl(self.pan_tb, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
        #                               wx.LC_EDIT_LABELS | wx.LC_HRULES | wx.LC_REPORT | wx.LC_VRULES)
        # self.tb_listbox.Hide()
        self.tb_listbox = CheckListCtrl(self.pan_tb)
        tb_left.Add(self.tb_listbox, 1, wx.EXPAND, 5)

        bSizer10 = wx.BoxSizer(wx.HORIZONTAL)

        self.tb_shuaxin = wx.Button(self.pan_tb, wx.ID_ANY, u"刷新", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer10.Add(self.tb_shuaxin, 0, wx.ALL, 5)

        self.tb_lssj = wx.Button(self.pan_tb, wx.ID_ANY, u"历史数据", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer10.Add(self.tb_lssj, 0, wx.ALL, 5)

        tb_left.Add(bSizer10, 0, 0, 5)

        self.tb_ls = wx.adv.CalendarCtrl(self.pan_tb, wx.ID_ANY, wx.DefaultDateTime, wx.DefaultPosition, wx.DefaultSize,
                                         wx.adv.CAL_SHOW_HOLIDAYS)
        tb_left.Add(self.tb_ls, 0, wx.ALL | wx.EXPAND, 5)

        self.figure = Figure()
        self.figure1 = Figure()
        self.figure2 = Figure()
        self.figure3 = Figure()
        self.axes = self.figure.add_subplot(111)
        self.axes1 = self.figure1.add_subplot(111)
        self.axes2 = self.figure2.add_subplot(111)
        self.axes3 = self.figure3.add_subplot(111)
        t = numpy.arange(0.0, 3.0, 0.01)
        s = numpy.sin(2 * numpy.pi * t)
        self.axes.plot(t, s)
        self.axes1.plot(t, s)
        self.axes2.plot(t, s)
        self.axes3.plot(t, s)
        # self.canvas = FigureCanvasWxAgg(self.pan_tb, -1, self.figure)
        # self.canvas1 = FigureCanvasWxAgg(self.pan_tb, -1, self.figure1)
        # self.canvas2 = FigureCanvasWxAgg(self.pan_tb, -1, self.figure2)
        # self.canvas3 = FigureCanvasWxAgg(self.pan_tb, -1, self.figure3)

        bSizer81.Add(tb_left, 0, wx.EXPAND, 5)
        #
        # gSizer1 = wx.GridSizer(0, 2, 0, 0)
        # gSizer1.Add(self.canvas, 0, wx.LEFT | wx.TOP | wx.GROW, 5)
        #
        # gSizer1.Add(self.canvas1, 1, wx.LEFT | wx.TOP | wx.GROW, 5)
        #
        # gSizer1.Add(self.canvas2, 2, wx.LEFT | wx.TOP | wx.GROW, 5)
        #
        # gSizer1.Add(self.canvas3, 3, wx.LEFT | wx.TOP | wx.GROW, 5)
        #
        # bSizer81.Add(gSizer1, 1, wx.EXPAND, 5)

        bSizer101 = wx.BoxSizer(wx.VERTICAL)

        self.m_notebook1 = wx.Notebook(self.pan_tb, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_panel81 = wx.Panel(self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer11 = wx.BoxSizer(wx.VERTICAL)

        self.canvas = FigureCanvasWxAgg(self.m_panel81, -1, self.figure)
        bSizer11.Add(self.canvas, 1, wx.ALL | wx.EXPAND, 5)

        self.m_panel81.SetSizer(bSizer11)
        self.m_panel81.Layout()
        bSizer11.Fit(self.m_panel81)
        self.m_notebook1.AddPage(self.m_panel81, u"1", False)
        self.m_panel91 = wx.Panel(self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer12 = wx.BoxSizer(wx.VERTICAL)

        self.canvas1 = FigureCanvasWxAgg(self.m_panel91, -1, self.figure1)
        bSizer12.Add(self.canvas1, 1, wx.ALL | wx.EXPAND, 5)

        self.m_panel91.SetSizer(bSizer12)
        self.m_panel91.Layout()
        bSizer12.Fit(self.m_panel91)
        self.m_notebook1.AddPage(self.m_panel91, u"2", True)
        self.m_panel101 = wx.Panel(self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer13 = wx.BoxSizer(wx.VERTICAL)

        self.canvas2 = FigureCanvasWxAgg(self.m_panel101, -1, self.figure2)
        bSizer13.Add(self.canvas2, 1, wx.ALL | wx.EXPAND, 5)

        self.m_panel101.SetSizer(bSizer13)
        self.m_panel101.Layout()
        bSizer13.Fit(self.m_panel101)
        self.m_notebook1.AddPage(self.m_panel101, u"3", False)
        self.m_panel111 = wx.Panel(self.m_notebook1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer14 = wx.BoxSizer(wx.VERTICAL)

        self.canvas3 = FigureCanvasWxAgg(self.m_panel111, -1, self.figure3)
        bSizer14.Add(self.canvas3, 1, wx.ALL | wx.EXPAND, 5)

        self.m_panel111.SetSizer(bSizer14)
        self.m_panel111.Layout()
        bSizer14.Fit(self.m_panel111)
        self.m_notebook1.AddPage(self.m_panel111, u"4", False)

        bSizer101.Add(self.m_notebook1, 1, wx.EXPAND | wx.ALL, 5)

        bSizer81.Add(bSizer101, 1, wx.EXPAND, 5)


        self.pan_tb.SetSizer(bSizer81)
        self.pan_tb.Layout()
        bSizer81.Fit(self.pan_tb)
        self.auinotebook.AddPage(self.pan_tb, u"图表", False, wx.NullBitmap)
        self.pan_z = wx.Panel(self.auinotebook, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        bSizer101 = wx.BoxSizer(wx.HORIZONTAL)

        z_left = wx.BoxSizer(wx.VERTICAL)

        z_comboBox1Choices = []
        self.z_comboBox1 = wx.ComboBox(self.pan_z, wx.ID_ANY, u"Combo!", wx.DefaultPosition, wx.DefaultSize,
                                       z_comboBox1Choices, 0)
        z_left.Add(self.z_comboBox1, 0, wx.ALL | wx.EXPAND, 5)

        self.z_listbox = ListCtrlAutoCtrl(self.pan_z)
        self.z_listbox.SetFont(
            wx.Font(15, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        z_left.Add(self.z_listbox, 1, wx.ALL | wx.EXPAND, 5)
        self.z_listbox.SetMaxSize(wx.Size(300, -1))

        bSizer13 = wx.BoxSizer(wx.HORIZONTAL)

        self.z_shuaxin = wx.Button(self.pan_z, wx.ID_ANY, u"刷新", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer13.Add(self.z_shuaxin, 0, wx.ALL, 5)

        self.z_lssj = wx.Button(self.pan_z, wx.ID_ANY, u"历史数据", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer13.Add(self.z_lssj, 0, wx.ALL, 5)

        z_left.Add(bSizer13, 0, 0, 5)

        self.m_calendar2 = wx.adv.CalendarCtrl(self.pan_z, wx.ID_ANY, wx.DefaultDateTime, wx.DefaultPosition,
                                               wx.DefaultSize, wx.adv.CAL_SHOW_HOLIDAYS)
        z_left.Add(self.m_calendar2, 0, wx.ALL| wx.EXPAND, 5)

        bSizer101.Add(z_left, 0, wx.EXPAND, 5)

        gSizer2 = wx.GridSizer(0, 2, 0, 0)

        self.z_figure = Figure()
        self.z_figure1 = Figure()
        self.z_figure2 = Figure()

        self.z_axes = self.z_figure.add_subplot(111)
        self.z_axes1 = self.z_figure1.add_subplot(111)
        self.z_axes2 = self.z_figure2.add_subplot(111)

        t = numpy.arange(0.0, 3.0, 0.01)
        s = numpy.sin(2 * numpy.pi * t)
        self.z_axes.plot(t, s)
        self.z_axes1.plot(t, s)
        self.z_axes2.plot(t, s)

        self.z_canvas = FigureCanvasWxAgg(self.pan_z, -1, self.z_figure)
        self.z_canvas1 = FigureCanvasWxAgg(self.pan_z, -1, self.z_figure1)
        self.z_canvas2 = FigureCanvasWxAgg(self.pan_z, -1, self.z_figure2)
        self.z_TJlistbox = ListCtrlAutoCtrl(self.pan_z)
        gSizer2.Add(self.z_canvas, 0, wx.LEFT | wx.TOP | wx.GROW, 5)

        gSizer2.Add(self.z_canvas1, 1, wx.LEFT | wx.TOP | wx.GROW, 5)

        gSizer2.Add(self.z_canvas2, 2, wx.LEFT | wx.TOP | wx.GROW, 5)

        gSizer2.Add(self.z_TJlistbox, 3, wx.LEFT | wx.TOP | wx.GROW, 5)

        bSizer101.Add(gSizer2, 1, wx.EXPAND, 5)

        self.pan_z.SetSizer(bSizer101)
        self.pan_z.Layout()
        bSizer101.Fit(self.pan_z)
        self.auinotebook.AddPage(self.pan_z, u"线体明细", True, wx.NullBitmap)

        bSizer8.Add(self.auinotebook, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(bSizer8)
        self.Layout()

        self.Centre(wx.BOTH)

    def __del__(self):
        pass
