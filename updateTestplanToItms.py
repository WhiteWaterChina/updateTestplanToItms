#!/usr/bin/env python
# -*- coding:cp936 -*-
# Author:yanshuo@inspur.com

import xlsxwriter
import os
import time
import wx
import xlrd


class UpdateTestplanToItms(wx.Frame):

    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=u"将测试方案中配置同步到ITMS导入模板工具", pos=wx.DefaultPosition,
                          size=wx.Size(504, 680), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_APPWORKSPACE))

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        self.m_panel1 = wx.Panel(self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.TAB_TRAVERSAL)
        self.m_panel1.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOWFRAME))

        bSizer10 = wx.BoxSizer(wx.VERTICAL)

        bSizer3 = wx.BoxSizer(wx.VERTICAL)

        self.text_title1 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 1.请选择要导出配置的EXCEL表格！", wx.DefaultPosition,
                                         wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title1.Wrap(-1)

        self.text_title1.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title1.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title1.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer3.Add(self.text_title1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer3, 0, wx.EXPAND, 5)

        bSizer16 = wx.BoxSizer(wx.VERTICAL)

        bSizer10.Add(bSizer16, 0, 0, 5)

        bSizer9 = wx.BoxSizer(wx.HORIZONTAL)

        self.text_PathInputExcel = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                               wx.DefaultSize, wx.TE_MULTILINE)
        bSizer9.Add(self.text_PathInputExcel, 1, wx.ALL | wx.ALIGN_BOTTOM | wx.EXPAND, 5)

        self.btn_ChoseInputExcel = wx.Button(self.m_panel1, wx.ID_ANY, u"选择Excel", wx.DefaultPosition, wx.DefaultSize,
                                             0)
        bSizer9.Add(self.btn_ChoseInputExcel, 0, wx.ALL | wx.EXPAND, 5)

        bSizer10.Add(bSizer9, 0, wx.EXPAND, 5)

        bSizer101 = wx.BoxSizer(wx.VERTICAL)

        self.btn_GetSheetName = wx.Button(self.m_panel1, wx.ID_ANY, u"获取Excel中所有Sheet的名称", wx.DefaultPosition,
                                          wx.DefaultSize, 0)
        bSizer101.Add(self.btn_GetSheetName, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)

        bSizer10.Add(bSizer101, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.EXPAND, 5)

        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        bSizer19 = wx.BoxSizer(wx.VERTICAL)

        self.text_title11 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 2.请选择要导出配置的Sheet的名称！", wx.DefaultPosition,
                                          wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title11.Wrap(-1)

        self.text_title11.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title11.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title11.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer19.Add(self.text_title11, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer4.Add(bSizer19, 0, wx.EXPAND | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer20 = wx.BoxSizer(wx.VERTICAL)

        listbox_SheetNameChoices = []
        self.listbox_SheetName = wx.ListBox(self.m_panel1, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize,
                                            listbox_SheetNameChoices, 0)
        bSizer20.Add(self.listbox_SheetName, 1, wx.ALL | wx.EXPAND, 5)

        bSizer4.Add(bSizer20, 1, wx.EXPAND, 5)

        bSizer10.Add(bSizer4, 1, wx.EXPAND, 5)

        bSizer14 = wx.BoxSizer(wx.VERTICAL)

        bSizer15 = wx.BoxSizer(wx.VERTICAL)

        self.text_title121 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 3.请在此选择填写测试团队名称！\n一定要跟测试配置中的完全一样！",
                                           wx.DefaultPosition, wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title121.Wrap(-1)

        self.text_title121.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title121.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title121.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer15.Add(self.text_title121, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer14.Add(bSizer15, 0, wx.EXPAND, 5)

        bSizer161 = wx.BoxSizer(wx.VERTICAL)

        self.text_TeamName = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                         0)
        bSizer161.Add(self.text_TeamName, 0, wx.ALL | wx.EXPAND, 5)

        bSizer14.Add(bSizer161, 1, wx.EXPAND, 5)

        bSizer10.Add(bSizer14, 0, wx.EXPAND, 5)

        bSizer21 = wx.BoxSizer(wx.VERTICAL)

        bSizer211 = wx.BoxSizer(wx.VERTICAL)

        self.text_title12 = wx.StaticText(self.m_panel1, wx.ID_ANY, u"Step 4.请点击GO开始导出！或者点击EXIT退出程序！",
                                          wx.DefaultPosition, wx.DefaultSize, wx.ST_NO_AUTORESIZE)
        self.text_title12.Wrap(-1)

        self.text_title12.SetFont(
            wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        self.text_title12.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_title12.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer211.Add(self.text_title12, 0, wx.EXPAND, 5)

        bSizer21.Add(bSizer211, 0, wx.EXPAND, 5)

        bSizer22 = wx.BoxSizer(wx.HORIZONTAL)

        self.button_go = wx.Button(self.m_panel1, wx.ID_ANY, u"GO", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer22.Add(self.button_go, 0, wx.ALL, 5)

        self.button_exit = wx.Button(self.m_panel1, wx.ID_ANY, u"EXIT", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer22.Add(self.button_exit, 0, wx.ALL, 5)

        bSizer21.Add(bSizer22, 0, wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer10.Add(bSizer21, 0, wx.ALIGN_CENTER_HORIZONTAL | wx.ALIGN_CENTER_VERTICAL, 5)

        bSizer91 = wx.BoxSizer(wx.VERTICAL)

        self.textctrl_display = wx.TextCtrl(self.m_panel1, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition,
                                            wx.DefaultSize, wx.TE_MULTILINE | wx.TE_READONLY)
        bSizer91.Add(self.textctrl_display, 1, wx.ALL | wx.EXPAND, 5)

        bSizer10.Add(bSizer91, 1, wx.EXPAND, 5)

        self.m_panel1.SetSizer(bSizer10)
        self.m_panel1.Layout()
        bSizer10.Fit(self.m_panel1)
        bSizer2.Add(self.m_panel1, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(bSizer2)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.btn_ChoseInputExcel.Bind(wx.EVT_BUTTON, self.get_inputexcel)
        self.btn_GetSheetName.Bind(wx.EVT_BUTTON, self.get_inputexcelsheetname)
        self.button_go.Bind(wx.EVT_BUTTON, self.onbutton)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

    def __del__(self):
        pass

    def get_inputexcel(self, event):
        global filepath_inputexcel
        filepath_inputexcel_dialog = wx.FileDialog(self, message=u"选择测试方案Excel文件", defaultDir=os.getcwd(), defaultFile="")
        if filepath_inputexcel_dialog.ShowModal() == wx.ID_OK:
            filename_inputexcel = filepath_inputexcel_dialog.GetPath()
            self.text_PathInputExcel.SetValue(filename_inputexcel)
            filepath_inputexcel = filename_inputexcel
            filepath_inputexcel_dialog.Destroy()
            self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
            self.updatedisplay("测试方案的Excel的文件路径和文件名：{}".format(filename_inputexcel))

    def get_inputexcelsheetname(self, event):
        try:
            workbook_inputexcel = xlrd.open_workbook(filename=filepath_inputexcel)
            sheetnames_list = workbook_inputexcel.sheet_names()
            for item_sheetname in sheetnames_list:
                self.listbox_SheetName.Append(item_sheetname)
        except NameError:
            self.updatedisplay("未选择测试方案文件，请选择！")
            diag_error_input = wx.MessageDialog(None, "未选择测试方案文件，请选择！", '错误', wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
            diag_error_input.ShowModal()



    def onbutton(self, event):
        self.button_go.Disable()
        sheetname_seleted = self.listbox_SheetName.GetStringSelection()
        teamname_selected = self.text_TeamName.GetValue()
        if len(sheetname_seleted) == 0:
            self.updatedisplay("未选择Sheet名称，请选择！")
            diag_error_sheetname = wx.MessageDialog(None, "未选择Sheet名称，请选择！", '错误', wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
            diag_error_sheetname.ShowModal()
            self.button_go.Enable()
        else:
            if len(teamname_selected) == 0:
                self.updatedisplay("未填写测试团队名称，请填写！")
                diag_error_sheetname = wx.MessageDialog(None, "未填写测试团队名称，请填写！", '错误', wx.OK | wx.ICON_ERROR | wx.STAY_ON_TOP)
                diag_error_sheetname.ShowModal()
                self.button_go.Enable()
            else:
                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                self.updatedisplay("选择的Sheet名称是：{}".format(sheetname_seleted))
                self.updatedisplay("选择的测试团队名称是：{}".format(teamname_selected))
                timestamp = time.strftime('%Y%m%d', time.localtime())
                # add output workbook
                workbook_output = xlsxwriter.Workbook("测试配置转换结果-{}-{}-{}.xlsx".format(sheetname_seleted, teamname_selected, timestamp))
                self.updatedisplay("创建输出文档：《测试配置转换结果-{}-{}-{}.xlsx》".format(sheetname_seleted, teamname_selected, timestamp))
                formatOne = workbook_output.add_format({'border': 1})
                formatTitle = workbook_output.add_format({'bold': True, 'border': 1})
                merge_format = workbook_output.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'border': 1})
                # open input workbook excel
                workbook_inputexcel = xlrd.open_workbook(filename=filepath_inputexcel)
                sheet_selected = workbook_inputexcel.sheet_by_name(sheetname_seleted)
                all_rows = sheet_selected.nrows
                for item_line in range(0, all_rows):
                    teamname_excel = sheet_selected.cell(item_line, 1).value
                    if str(teamname_excel).strip() == teamname_selected:
                        configname = str(sheet_selected.cell(item_line+1, 1).value).strip()
                        # add new sheet to output workbook named after the config name
                        sheet_new_output = workbook_output.add_worksheet("{}".format(configname))
                        self.updatedisplay("输出文档增加sheet：{}".format(configname))
                        # set column width
                        sheet_new_output.set_column('A:A', 18)
                        sheet_new_output.set_column('B:B', 25)
                        sheet_new_output.set_column('C:C', 45)
                        sheet_new_output.set_column('D:D', 40)
                        sheet_new_output.set_column('E:E', 20)
                        sheet_new_output.set_column('F:F', 10)
                        sheet_new_output.set_column('G:G', 40)

                        sheet_new_output.merge_range(0, 0, 0, 6, configname, merge_format)
                        sheet_new_output.write(1, 0, "Hardware", formatTitle)
                        sheet_new_output.write(1, 1, "PN", formatTitle)
                        sheet_new_output.write(1, 2, "Model Name", formatTitle)
                        sheet_new_output.write(1, 3, "Location", formatTitle)
                        sheet_new_output.write(1, 4, "Firmware", formatTitle)
                        sheet_new_output.write(1, 5, "Qty", formatTitle)
                        sheet_new_output.write(1, 6, "Remarks", formatTitle)
                        line_to_write = 2
                        # get detail info from input excel and write to output workbook
                        for line_detail in range(item_line+3, all_rows):
                            title_line = sheet_selected.cell(line_detail, 0).value
                            if len(title_line) == 0:
                                break
                            else:
                                # get detail info for every data line
                                hardware_type = sheet_selected.cell(line_detail, 0).value
                                pn = sheet_selected.cell(line_detail, 1).value
                                modelname = sheet_selected.cell(line_detail, 2).value
                                location = sheet_selected.cell(line_detail, 3).value
                                firmware = sheet_selected.cell(line_detail, 4).value
                                qty = sheet_selected.cell(line_detail, 5).value
                                remarks = sheet_selected.cell(line_detail, 6).value
                                # write data line info to output sheet
                                sheet_new_output.write(line_to_write, 0, hardware_type, formatOne)
                                sheet_new_output.write(line_to_write, 1, pn, formatOne)
                                sheet_new_output.write(line_to_write, 2, modelname, formatOne)
                                sheet_new_output.write(line_to_write, 3, location, formatOne)
                                sheet_new_output.write(line_to_write, 4, firmware, formatOne)
                                sheet_new_output.write(line_to_write, 5, qty, formatOne)
                                sheet_new_output.write(line_to_write, 6, remarks, formatOne)

                                line_to_write = line_to_write + 1

                workbook_output.close()
                self.updatedisplay(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
                self.updatedisplay("".format(workbook_output))
                self.updatedisplay("配置信息获取完毕，保存在Excel：《测试配置转换结果-{}-{}-{}.xlsx》 中！".format(sheetname_seleted, teamname_selected, timestamp))
                diag_finish = wx.MessageDialog(None, "配置信息获取完毕，保存在Excel：《测试配置转换结果-{}-{}-{}.xlsx》 中！".format(sheetname_seleted, teamname_selected, timestamp), '提示', wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
                diag_finish.ShowModal()
                self.button_go.Enable()

    def close(self, event):
        self.Close()

    def updatedisplay(self, msg):
        t = msg
        if isinstance(t, int):
            self.textctrl_display.AppendText(u"完成第%s页" % t)
        elif t == "Finished":
            self.button_go.Enable()
        else:
            self.textctrl_display.AppendText(u"%s" % t)
        self.textctrl_display.AppendText(os.linesep)


if __name__ == '__main__':
    app = wx.App()
    frame = UpdateTestplanToItms(None)
    frame.Show()
    app.MainLoop()