import pandas as pd  # 用于分析数据
import os  # 导入os模块
import xlwt
import xlrd
import copy
import numpy as np
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
import sys

# 加载窗口
class MyWindow(QtWidgets.QWidget):
    def __init__(self):
        super(MyWindow, self).__init__()

    def msg(self):
        fileName1, filetype = QFileDialog.getOpenFileName(self, "选取文件", "./",
                                                          "All Files (*);;Excel Files (*.xls)")  # 设置文件扩展名过滤,注意用双分号间隔
        return fileName1



def load_keyname():  #返回一个字典
    if os.path.exists('data_name.xls') != True:
        workbook = xlwt.Workbook(encoding='utf-8')
        booksheet = workbook.add_sheet('Sheet1', cell_overwrite_ok=True)
        workbook.save('data_name.xls')
    read = pd.read_excel('data_name.xls')
    ts = read.iloc[:, 0:read.columns.size].values  # 读取所有
    key_dic = dict()
    for row in range(0, len(ts)):
        for col in range(0, len(ts[row])):
            key_dic.setdefault(ts[row][col],[])
    liKey = []   #取出字典的键，存入到一个列表中
    for key in key_dic.keys():
        liKey.append(key)  # 把关键字变成一行列表
    return key_dic, liKey #返回关键字字典和关键字列表

def load_data(keyStrs,filename): #返回的是数据字典
    FileObj = xlrd.open_workbook(filename)  # 打开文件
    sheet = FileObj.sheets()[0]  # 获取第一个工作表
    row_count = sheet.nrows  # 行数
    col_count = sheet.ncols  # 列数
    print("列数：",col_count)
    print("行数：",row_count)
    time_dic = {} #建立一个存储数据的总字典
    for index in range(1,row_count): #建立字典的键
        if sheet.row_values(index)[2] not in  time_dic.keys(): #是否有对应的键
            time_dic.setdefault(sheet.row_values(index)[2], copy.deepcopy(keyStrs)) #没有就添加新的键值对
        if sheet.row_values(index)[col_count - 2] in keyStrs.keys(): #值是否存在
            time_dic[sheet.row_values(index)[2]][sheet.row_values(index)[col_count - 2]] = sheet.row_values(index)[col_count - 1] #不存在就加入相应的值
    liContent = []
    for value in time_dic.values():
        for i in value.values():
            liContent.append(i) #把字典中的值存到列表中
    return time_dic,liContent #返回字典，和字典值的列表

def writeExcelFile(filename, header, content):
    # 因为输入都是Unicode字符，这里使用utf-8，免得来回转换
    workbook = xlwt.Workbook(encoding='utf-8')
    booksheet = workbook.add_sheet('Sheet1', cell_overwrite_ok=True)
    # 写列头,表示excel文件中设备数据名
    row = 0
    for col in range(len(header)):
        booksheet.write(row, col, header[col])
    # 写每一列
    for col in range(0, len(content)):
        row = int(col/len(header)) #确定行
        coll = col % len(header)   #确定列
        booksheet.write(row + 1, coll, content[col])
    # 保存文件
    workbook.save(filename)

#分开excel文件
def splictExcelFile(num):
    read = pd.read_excel('./Classification_File.xls')
    ncol = int(read.columns.size/num)#每个文件的列数
    for i in range(num):
        ts = read.iloc[:, ncol*i : ncol*(i+1)].values  # 读取所有列
        print(ts)
        workbook = xlwt.Workbook(encoding='utf-8') # 因为输入都是Unicode字符，这里使用utf-8，免得来回转换
        booksheet = workbook.add_sheet('Sheet_1', cell_overwrite_ok=True)
        #写第一行
        for col in range(ncol):
            booksheet.write(0, col, read.columns.values[ncol*i + col])

        #复制每一列
        matrix = [[0] * (len(ts)-1)] * len(ts)
        for row in range(1, len(ts)):
            for col in range(0, ncol):
                if (np.isnan(ts[row][col])):
                    print("")

        # 写每一列
        for row in range(1, len(ts)):
            for col in range(0, len(ts[row])):
                booksheet.write(row, col, ts[row][col])

        workbook.save(str(read.columns.values[ncol*i + ncol - 1])+ "#_file.xls")# 保存文件

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    myshow = MyWindow()
    filename = myshow.msg() #加载指定的文件
    keyStrs,liKey = load_keyname() #加载关键字文件
    data_dic,liContent = load_data(keyStrs, filename) #,参数是一个字典
    #写excel文件
    fileName =  "Classification_File.xls"
    writeExcelFile(fileName, liKey, liContent)
    splictExcelFile(4)
    os.remove(fileName)

