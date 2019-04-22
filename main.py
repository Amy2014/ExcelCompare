# -*- coding: UTF-8 -*-

import sys
import getopt
import json
import time

from functools import partial

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter.ttk import *

from differ import ExcelDiffer, ExcelHelper

import tktable

limit = 6

class ScrollDummy(object):
    def __init__(self, table):
        self.table = table

    def xview(self, *args):
        if args[0] == 'scroll':
            self.table.xview_scroll(*args[1:])
        if args[0] == 'moveto':
            self.table.xview_moveto(*args[1:])

    def yview(self, *args):
        if args[0] == 'scroll':
            self.table.yview_scroll(*args[1:])
        if args[0] == 'moveto':
            self.table.yview_moveto(*args[1:])

class ScrollDataDummy(object):

    SCROLL_TYPE_COL = 1
    SCROLL_TYPE_ROW = 2
    SCROLL_TYPE_CELL = 3

    def __init__(self, tabFrame, data, tkTable, scrollType=None):
        self.tkTable = tkTable
        self.tabFrame = tabFrame
        self.data = data
        self.idx = 0
        self.scrollType = scrollType

    def yview(self, *args):
        print(len(self.data))
        #n = len(self.data) - limit
        n = len(self.data) - (self.idx+6)
        print("self.idx==="+str(self.idx))
        step = 0
        if args[0] == "scroll":
            step = int(args[1])
        if args[0] == "moveto":
            step = int(float(args[1]))
        print("step==="+str(step))

        if step >0 and n<=0:
            return
        if step < 0 and self.idx <=0:
            print("n=" + str(n))
            self.idx = 0
            return
        self.idx += step

        if self.scrollType == ScrollDataDummy.SCROLL_TYPE_CELL:
            row = 1
            for key, _data in list(
                    self.data.items())[self.idx:self.idx + limit]:
                coordinate = ExcelHelper.CoordinateFromStr(key)
                first = "%i,%i" % (coordinate[1],
                                   ExcelHelper.ColumnIndexFromStr(
                                       coordinate[0]))

                l = tk.Button(
                    self.tabFrame,
                    text=key,
                    bg='white',
                    relief='groove',
                    width=8,
                    command=partial(self.tkTable.SelectCells, first))

                l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=0)
                l = Label(self.tabFrame, text=_data[0], font=(8))
                l.grid(
                    sticky=tk.W + tk.E + tk.N + tk.S,
                    row=row,
                    column=1,
                    padx=8)
                l = Label(self.tabFrame, text=_data[1], font=(8))
                l.grid(
                    sticky=tk.W + tk.E + tk.N + tk.S,
                    row=row,
                    column=2,
                    padx=8)
                row += 1

        if self.scrollType == ScrollDataDummy.SCROLL_TYPE_COL:
            row = 1
            for _data in self.data[self.idx:self.idx+limit]:
                l = Label(self.tabFrame, text=_data["action"], font=(8))
                l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=0, pady=5)

                first = "%i,%i" % (0, ExcelHelper.ColumnIndexFromStr(_data["label"]))
                last = "%i,%i" % (10, ExcelHelper.ColumnIndexFromStr(_data["label"]))
                l = tk.Button(
                    self.tabFrame,
                    text=_data["label"],
                    bg='white',
                    relief='groove',
                    width=8,
                    command=partial(self.tkTable.SelectCells, first, last))

                l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=1, pady=5)
                row += 1

        if self.scrollType == ScrollDataDummy.SCROLL_TYPE_ROW:
            row = 1
            for _data in self.data[self.idx:self.idx+limit]:
                l = Label(self.tabFrame, text=_data["action"], font=(8))
                l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=0, pady=5)

                first = "%i,%i" % (_data["label"], 0)
                last = "%i,%i" % (_data["label"], 10)
                l = tk.Button(
                    self.tabFrame,
                    text=_data["label"],
                    bg='white',
                    relief='groove',
                    command=partial(self.tkTable.SelectCells, first, last))

                l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=1, pady=5)
                row += 1


class MyApp(tk.Tk):
    def __init__(self, srcPath=None, dstPath=None):
        super().__init__()

        self.srcPath = srcPath
        self.dstPath = dstPath
        self.tableFrame = None
        self.tabControl = None
        self.frame = None

        self.diffResults = {}
        self.lastSelectCells = None

        self.InitFrame()

        self.InitTableFlame(srcPath, dstPath, 0, 0)

        self.InitButtonFlame()

        self.InitTabFlame()

    def InitFrame(self):
        self.title("Excel Compare")

        w, h = self.maxsize()
        self.geometry("{}x{}".format(w, h))

        #self.columnconfigure(0, weight=1)
        #self.rowconfigure(0, weight=1)

    def InitTableTitleFlame(self, tableFrame,srcPath, dstPath):
        if srcPath is None:
            srcPath = ""
        if dstPath is None:
            dstPath = ""
        #tableTitleFrame = Frame(self)
        # srcPathLabel = Label(tableTitleFrame, text=srcPath)
        # srcPathLabel.grid(row=0, column=0)

        srcPathLabel = Label(tableFrame, text=srcPath)
        srcPathLabel.grid(row=0, column=0,sticky=tk.E)

        dstPathLabel = Label(tableFrame, text=dstPath)
        dstPathLabel.grid(row=0, column=2,sticky=tk.E)


    def InitTableSheetFlame(self, tableFrame, srcPath, dstPath, srcExcel, dstExcel, srcIndex, dstIndex):

        srcSheetNames = srcExcel.GetSheetsName()
        dstSheetNames = dstExcel.GetSheetsName()

        srcSheetFrame = Frame(tableFrame)
        srcSheetFrame.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=3, column=0)

        dstSheetFrame = Frame(tableFrame)
        dstSheetFrame.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=3, column=2)

        src_col_count = 0
        for sheet_title in srcSheetNames:
            if src_col_count == srcIndex:
                srcSheetButton = tk.Button(
                    srcSheetFrame,
                    text=sheet_title,
                    width=8,
                    height=1,
                    relief='groove',
                    font=("11"),
                    bg='LightGrey',
                    fg='black',
                )
            else:
                srcSheetButton = tk.Button(
                    srcSheetFrame,
                    text=sheet_title,
                    width=8,
                    height=1,
                    relief='groove',
                    font=("11"),
                    bg='WhiteSmoke',
                    fg='black',
                    command=partial(self.InitTableFlame, srcPath, dstPath, src_col_count, dstIndex)
                )
            srcSheetButton.grid(row=0, column=src_col_count, padx=5, pady=10)
            src_col_count += 1

        dst_col_count = 0
        for sheet_title in dstSheetNames:
            if dst_col_count == dstIndex:
                dstSheetButton = tk.Button(
                    dstSheetFrame,
                    text=sheet_title,
                    width=8,
                    height=1,
                    relief='groove',
                    font=("11"),
                    bg='LightGrey',
                    fg='black',
                )
            else:
                dstSheetButton = tk.Button(
                    dstSheetFrame,
                    text=sheet_title,
                    width=8,
                    height=1,
                    relief='groove',
                    font=("11"),
                    bg='WhiteSmoke',
                    fg='black',
                    command=partial(self.InitTableFlame, srcPath, dstPath, srcIndex, dst_col_count)
                )
            dstSheetButton.grid(row=0, column=dst_col_count, padx=5, pady=10)
            dst_col_count += 1


    def InitTableFlame(self, srcPath, dstPath, srcIndex, dstIndex):

        if self.tableFrame:
            self.tableFrame.destroy()

        if not srcPath or not dstPath:
            self.srcPath = srcPath
            self.dstPath = dstPath
            self.tableFrame = Frame(self)
            self.tableFrame.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=1, column=0)
            return

        self.srcExcel = ExcelHelper.OpenExcel(srcPath, srcIndex)
        self.srcPath = srcPath
        self.dstExcel = ExcelHelper.OpenExcel(dstPath, dstIndex)
        self.srcPath = srcPath

        self.tableFrame = Frame(self)

        self.InitTableTitleFlame(self.tableFrame,self.srcPath, self.dstPath)
        #分别显示表格sheet选项
        self.InitTableSheetFlame(self.tableFrame, self.srcPath, self.dstPath, self.srcExcel, self.dstExcel, srcIndex, dstIndex)

        maxRows = self.srcExcel.GetMaxRow() if self.srcExcel.GetMaxRow(
        ) >= self.dstExcel.GetMaxRow() else self.dstExcel.GetMaxRow()
        maxCols = self.srcExcel.GetMaxColumn() if self.srcExcel.GetMaxColumn(
        ) >= self.dstExcel.GetMaxColumn() else self.dstExcel.GetMaxColumn()

        maxRows += 1
        maxCols += 1

        self.maxRows = maxRows
        self.maxCols = maxCols

        self.table1, self.var1 = self.setTable(
            self.tableFrame,
            gridRow=1,
            gridColumn=0,
            rows=self.maxRows,
            cols=self.maxCols,
            excel=self.srcExcel)

        self.table2, self.var2 = self.setTable(
            self.tableFrame,
            gridRow=1,
            gridColumn=2,
            rows=self.maxRows,
            cols=self.maxCols,
            excel=self.dstExcel)

        self.tableFrame.grid(sticky="nsew", row=1, column=0)

        self.diffResults = {}
        diffResults = {}
        diffResults = ExcelDiffer.Diff2(self.srcExcel, self.dstExcel)

        self.SetDiffColor(diffResults)

        self.diffResults = diffResults
        self.InitTabFlame()

    def InitButtonFlame(self):
        buttonFrame = Frame(self)

        uploadFile1Button = tk.Button(
            buttonFrame,
            text="选择原文件",
            width=15,
            height=2,
            relief='groove',
            font=("13"),
            bg='DeepSkyBlue',
            fg='white',
            command=partial(self.UploadFile, "srcFile"))
        uploadFile1Button.grid(row=0, column=0, padx=5, pady=10)
        uploadFile1Button = tk.Button(
            buttonFrame,
            text="选择目标文件",
            width=15,
            height=2,
            relief='groove',
            font=("13"),
            bg='DeepSkyBlue',
            fg='white',
            command=partial(self.UploadFile, "dstFile"))
        uploadFile1Button.grid(row=0, column=2, padx=5, pady=10)
        #清空excel列表
        deleteFile1Button = tk.Button(
            buttonFrame,
            text="清空已选Excel",
            width=15,
            height=2,
            relief='groove',
            font=("13"),
            bg='DeepSkyBlue',
            fg='white',
            command=partial(self.DeleteFile))
        deleteFile1Button.grid(row=0, column=3, padx=5, pady=10)

        buttonFrame.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=2, column=0)

    def UploadFile(self, whitchFile):
        fileName = filedialog.askopenfilename()
        if whitchFile == "srcFile" and self.srcPath != fileName:
            self.srcPath = fileName
        if whitchFile == "dstFile":
            self.dstPath = fileName
        self.InitTableFlame(self.srcPath, self.dstPath, 0, 0)
        self.InitTabFlame()
    #清空表格
    def DeleteFile(self):
        self.srcPath = None
        self.dstPath = None
        self.InitTableFlame(self.srcPath, self.dstPath, 0, 0)
        self.diffResults = {}
        self.InitTabFlame()

    def setTable(self, tableFrame, gridRow, gridColumn, rows, cols, excel):
        tb = tktable.Table(
            tableFrame,
            selectmode="browse",
            state='disabled',
            width=8,
            height=11,
            font=(6),
            exportselection=0,
            titlerows=1,
            titlecols=1,
            rows=rows + 1,
            cols=cols + 1,
            colwidth=9)

        #### LIST OF LISTS DEFINING THE ROWS AND VALUES IN THOSE ROWS ####
        #### SETS THE DOGS INTO THE TABLE ####
        #DEFINING THE VAR TO USE AS DATA IN TABLE
        var = tktable.ArrayVar(tableFrame)

        row_count = 0
        col_count = 1
        #SETTING COLUMNS
        for col in range(0, cols):
            index = "%i,%i" % (row_count, col_count)
            var[index] = ExcelDiffer.GetColumnLeter(col)
            col_count += 1

        #SETTING ROWS
        row_count = 1
        col_count = 0
        for row in range(0, rows):
            index = "%i,%i" % (row_count, col_count)
            var[index] = row + 1
            row_count += 1

        #SETTING DATA IN ROWS
        row_count = 1
        col_count = 1
        for row in excel.data:
            for item in row:
                index = "%i,%i" % (row_count, col_count)
                ## PLACING THE VALUE IN THE INDEX CELL POSITION ##
                if item is None:
                    var[index] = ""
                else:
                    var[index] = item
                col_count += 1
            col_count = 1
            row_count += 1
        #### ABOVE CODE SETS THE DOG INTO THE TABLE ####
        ################################################
        #### VARIABLE PARAMETER SET BELOW ON THE 'TB' USES THE DATA DEFINED ABOVE ####
        tb['variable'] = var
        tb.tag_configure(
            'title',
            relief='raised',
            anchor='center',
            background='#D3D3D3',
            fg='BLACK',
            state='disabled')

        tb.width(**{"0": 5})

        xScrollbar = Scrollbar(tableFrame, orient='horizontal')
        xScrollbar.grid(row=gridRow + 1, column=gridColumn, sticky="ew")
        xScrollbar.config(command=tb.xview_scroll)
        tb.config(xscrollcommand=xScrollbar.set)

        yScrollbar = Scrollbar(tableFrame)
        yScrollbar.grid(row=gridRow, column=gridColumn+1, sticky="ns")
        yScrollbar.config(command=ScrollDummy(tb).yview)
        tb.config(yscrollcommand=yScrollbar.set)

        tb.grid(sticky="nsew", row=gridRow, column=gridColumn)
        return tb, var

    def _SetCommonHeader(self, tabControl, title, tabText, headers, data=None,scrollType=None):
        tab = ttk.Frame(tabControl)

        tabControl.add(tab, text=title, pad=5)

        monty = tk.LabelFrame(tab, text=tabText, font=(8))
        monty.grid(
            sticky=tk.W + tk.E + tk.N + tk.S, column=0, row=0, padx=5, pady=5)

        canvas = tk.Canvas(tab,width=800)  # 创建canvas
        canvas.grid(row=2, column=0)
        self.frame = Frame(canvas)  # 把frame放在canvas里
        self.frame.grid(row=0, column=0)  # frame的长宽，和canvas差不多的

        vbar = Scrollbar(tab, orient="vertical")  # 竖直滚动条
        vbar.grid(row=2, column=4, sticky="ns")
        vbar.configure(command=ScrollDataDummy(self.frame, data, self, scrollType).yview)
        canvas.config(yscrollcommand=vbar.set)  # 设置

        xbar = Scrollbar(tab, orient="horizon")  # 竖直滚动条
        xbar.grid(row=3, column=0, sticky="wes")
        xbar.configure(command=canvas.xview)

        canvas.config(yscrollcommand=vbar.set,xscrollcommand=xbar.set)  # 设置
        canvas.create_window(0, 0, window=self.frame, anchor="nw")

        row, col = 0, 0
        for header in headers:
            l = Label(self.frame, text=header, font=('', 15, 'bold'))
            l.grid(
                sticky=tk.W + tk.E + tk.N + tk.S,
                row=row,
                column=col,
                padx=10,
                pady=5)
            col += 1

        return self.frame


    def _SetRowTab(self, tabControl, title, tabText, headers, data=None):

        if not data:
            data = {"new":[], "del":[]}

        newData = []
        newData.extend([{"label":_data, "action": "新增"} for _data in data["new"]])
        newData.extend([{"label":_data, "action": "删除"} for _data in data["del"]])

        tabFrame = self._SetCommonHeader(tabControl, title, tabText, headers,
                                         newData, ScrollDataDummy.SCROLL_TYPE_ROW)

        if not newData:
            return

        row = 1
        for _data in newData[0:0+limit]:
            l = Label(tabFrame, text=_data["action"], font=(8))
            l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=0, pady=5)

            first = "%i,%i" % (_data["label"], 0)
            last = "%i,%i" % (_data["label"], 10)
            l = tk.Button(
                tabFrame,
                text=_data["label"],
                bg='white',
                relief='groove',
                command=partial(self.SelectCells, first, last))

            l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=1, pady=5)
            row += 1

    def _SetColumnTab(self, tabControl, title, tabText, headers, data=None):

        if not data:
            data = {"new":[], "del":[]}

        newData = []
        newData.extend([{"label":_data, "action": "新增"} for _data in data["new"]])
        newData.extend([{"label":_data, "action": "删除"} for _data in data["del"]])

        tabFrame = self._SetCommonHeader(tabControl, title, tabText, headers,
                                         newData, ScrollDataDummy.SCROLL_TYPE_COL)

        if not newData:
            return

        row = 1
        for _data in newData[0:0+limit]:
            l = Label(tabFrame, text=_data["action"], font=(8))
            l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=0, pady=5)

            first = "%i,%i" % (0, ExcelHelper.ColumnIndexFromStr(_data["label"]))
            last = "%i,%i" % (10, ExcelHelper.ColumnIndexFromStr(_data["label"]))
            l = tk.Button(
                tabFrame,
                text=_data["label"],
                bg='white',
                relief='groove',
                width=8,
                command=partial(self.SelectCells, first, last))

            l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=1, pady=5)
            row += 1

    def _SetCellTab(self, tabControl, title, tabText, headers, data=None):

        tabFrame = self._SetCommonHeader(tabControl, title, tabText, headers,
                                         data,
                                         ScrollDataDummy.SCROLL_TYPE_CELL)

        if not data:
            return

        row = 1
        for key, _data in list(data.items())[0:0 + limit]:
            coordinate = ExcelHelper.CoordinateFromStr(key)
            first = "%i,%i" % (coordinate[1],
                               ExcelHelper.ColumnIndexFromStr(coordinate[0]))

            l = tk.Button(
                tabFrame,
                text=key,
                bg='white',
                relief='groove',
                width=8,
                command=partial(self.SelectCells, first))

            l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=0)
            l = Label(tabFrame, text=_data[0], font=(8))
            l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=1, padx=8)
            l = Label(tabFrame, text=_data[1], font=(8))
            l.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=row, column=2, padx=8)
            row += 1

    def InitTabFlame(self):
        if self.tabControl:
            self.tabControl.destroy()
        if self.frame:
            self.frame.destroy()

        s = ttk.Style()
        s.configure('TNotebook', tabposition='nw')
        self.tabControl = ttk.Notebook(self)

        self._SetRowTab(self.tabControl, "行增删", "共计新增1行, 删除1行", ["改动", "行号"],
                        self.diffResults.get("rows"))
        self._SetColumnTab(self.tabControl, "列增删", "共计新增1列, 删除1列", ["改动", "列号"],
                           self.diffResults.get("columns"))

        self._SetCellTab(self.tabControl, "单元格改动", "共计2个单元格", ["坐标", "旧值", "新值"],
                         self.diffResults.get("cells"))

        self.tabControl.grid(sticky=tk.W + tk.E + tk.N + tk.S, row=3, column=0)

    def SetDiffColor(self, diffResults):
        for row in diffResults["rows"]["new"]:
            for col in range(0, self.maxCols + 1):
                index = "%i,%i" % (row, col)
                self.table1.tag_cell("new", index)
                self.table2.tag_cell("new", index)

        for row in diffResults["rows"]["del"]:
            for col in range(0, self.maxCols + 1):
                index = "%i,%i" % (row, col)
                self.table1.tag_cell("del", index)
                self.table2.tag_cell("del", index)

        for strCol in diffResults["columns"]["new"]:
            col = ExcelHelper.ColumnIndexFromStr(strCol)
            for row in range(0, self.maxRows + 1):
                index = "%i,%i" % (row, col)
                self.table1.tag_cell("new", index)
                self.table2.tag_cell("new", index)

        for strCol in diffResults["columns"]["del"]:
            col = ExcelHelper.ColumnIndexFromStr(strCol)
            for row in range(0, self.maxRows + 1):
                index = "%i,%i" % (row, col)
                self.table1.tag_cell("del", index)
                self.table2.tag_cell("del", index)

        for k, v in list(diffResults["cells"].items()):
            coordinate = ExcelHelper.CoordinateFromStr(k)
            index = "%i,%i" % (coordinate[1],
                               ExcelHelper.ColumnIndexFromStr(coordinate[0]))
            self.table1.tag_cell("mod", index)
            self.table2.tag_cell("mod", index)

        self.table1.tag_configure('new', background='green')
        self.table2.tag_configure('new', background='green')
        self.table1.tag_configure('del', background='red')
        self.table2.tag_configure('del', background='red')
        self.table1.tag_configure('mod', background='yellow')
        self.table2.tag_configure('mod', background='yellow')


    def SelectCells(self, first, last=None):
        if self.lastSelectCells:
            self.table1.selection_clear(self.lastSelectCells[0],
                                        self.lastSelectCells[1])
            self.table2.selection_clear(self.lastSelectCells[0],
                                        self.lastSelectCells[1])

        self.table1.selection_set(first, last)

        self.table2.selection_set(first, last)

        self.table1.see(first)

        self.table2.see(first)

        self.lastSelectCells = (first, last)


def Usage():
    print('''
-h/--help print this usage
-s/--src src excel path
-d/--dst dst excel path
    ''')
    return 0


def ParseArgv(argv):

    params = {}

    argvMap = {
        "h": "help",
        "s:": "src=",
        "d:": "dst=",
    }

    try:
        options, args = getopt.getopt(sys.argv[1:], "".join(
            list(argvMap.keys())), list(argvMap.values()))

        for name, value in options:
            if name in ('-h', '--help'):
                return Usage()
            elif name in ('-s', '--src'):
                params['srcPath'] = value
            elif name in ('-d', '--dst'):
                params['dstPath'] = value
        if not params:
            raise getopt.GetoptError("need path for excel")

    except getopt.GetoptError:
        return {}
        #return Usage()

    return params


if __name__ == '__main__':
    params = ParseArgv(sys.argv)

    if isinstance(params, int):
        sys.exit(params)

    if not params:
        MyApp().mainloop()
    else:
        MyApp(params["srcPath"], params["dstPath"]).mainloop()
