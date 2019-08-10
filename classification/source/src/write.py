# -*- coding: utf-8 -*-
"""
Created on Mon Jul 15 10:22:59 2019

@author: sunwf1114
"""
import xlwt
import pandas as pd

def writeExcelFile(num):
    read = pd.read_csv('./Classification_File.xls')
    ncol = int(read.columns.size/num)#每个文件的列数
    for i in range(num):
        ts = read.iloc[:, ncol*i : ncol*(i+1)].values  # 读取所有列
        workbook = xlwt.Workbook(encoding='utf-8') # 因为输入都是Unicode字符，这里使用utf-8，免得来回转换
        booksheet = workbook.add_sheet('Sheet_1', cell_overwrite_ok=True)
        #写第一行
        for col in range(ncol):
            booksheet.write(0, col, read.columns.values[ncol*i + col])
        # 写每一列
        for row in range(1, len(ts)):
            for col in range(0, len(ts[row])):
                booksheet.write(row, col, ts[row][col])
        workbook.save(str(read.columns.values[ncol*i + ncol - 1])+ "#_file.xls")# 保存文件


if __name__ == "__main__":
    writeExcelFile(4) #数字表示创造文件的数量

