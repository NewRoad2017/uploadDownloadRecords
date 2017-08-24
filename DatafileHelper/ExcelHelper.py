#coding=utf-8
'''__author__ ChenLu'''

import xlwt
import xlrd
import csv
#from xlutils.copy import copy
import sys
sys.path.append(r'D:\Projects\Chinare\Datas\applications\python\classes')
from infolists import Infolist


class Excel:
    def __init__(self, datafilePath):
        self.filePath = datafilePath
        self.workbook = None
        self.sheets = []


class Excelset(Excel):
    def __setExcel(self, dataInfolists):
        try:
            self.workbook = xlwt.Workbook()
            for infolist in dataInfolists:
                sheet = self.workbook.add_sheet(infolist.infotype)
                colcount = len(infolist.listhead)
                for num in range(2, colcount):
                    sheet.write(0, num-2, infolist.listhead[num])
                self.sheets.append(sheet)#这里需要添加调试
            self.workbook.save(self.filePath)
        except Exception as e:
            print str(e)
            sys.exit()

    def __init__(self, dataFileDirpath, dataFilename, dataInfolists):
        self.__fileName = dataFilename
        filepath = dataFileDirpath + '/' + dataFilename + '.xls'
        Excel.__init__(self, filepath)
        self.__setExcel(dataInfolists)

    def excelWrite(self,sheet,dataList):#need to optimize
        for row in range(0,len(dataList)):
            for col in range(2,len(dataList[row])):
                sheet.write(row+1,col-2,dataList[row][col])
        self.workbook.save(self.filePath)


class Excelget(Excel):
    def __getExcel(self):
        try:
            self.workbook = xlrd.open_workbook(self.filePath)
            self.sheets = self.workbook.sheets()
            self.listinfos = []
        except Exception as e:
            print str(e)
            sys.exit()

    def __init__(self, dataFilepath):
        Excel.__init__(self, dataFilepath)
        self.__getExcel()

# excel -> Infolist[]
    def excel2list(self):
        for sheet in self.sheets:
            name = sheet.name
            listHead = []
            listValue = []
            rowcount = sheet.nrows
            colcount = sheet.ncols
            for row in range(0, rowcount):
                if row == 0:
                    for col in range(0, colcount):
                        value = sheet.cell(row, col).value
                        listHead.append(value)
                else:
                    listRow = []
                    for col in range(0, colcount):
                        value = sheet.cell(row, col).value
                        listRow.append(value)
                    listValue.append(listRow)
            self.listinfos.append(Infolist(name, listHead, listValue))


# class Excelappend(Excel):
#     def __init__(self,dataFilepath,sheetid):
#         Excel(dataFilepath)
#
#     def __appendExcel(self):
#         try:
#             self.workbook = copy(xlrd.open_workbook(self.filePath))
#             self.sheets = self.workbook.sheets()
#             self.listinfos = []
#         except Exception as e:
#             print str(e)
#             sys.exit()





class CSVset:
    __fileDirpath = ""
    __fileName = ""
#==============================================================================
#     #共有数据成员
#==============================================================================
    csvFilepath = ""
#    excel中所有的东西
    csvFile = None
    csvWriter = None
#==============================================================================
#     私有方法
#==============================================================================
    def __setCSV(self,dataFields):
        try:
            self.csvFile = file(self.csvFilepath, 'wb')
            self.csvWriter = csv.writer(self.csvFile)
            self.csvWriter.writerow(dataFields)
        except Exception as e:
            print str(e)
            sys.exit()
#==============================================================================
#     共有方法
#==============================================================================
    def __init__(self, dataFileDirpath, dataFilename, dataFields):
        self.csvFilepath = dataFileDirpath + '/' + dataFilename + '.csv'
        self.__setCSV(dataFields)
