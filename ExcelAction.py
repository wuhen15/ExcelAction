#!/usr/bin/env python
# -*- coding: utf-8 -*-
# 读取excel数据
# 取第二行以下的数据，然后取每行前2列的数据
import  xlrd
from  xlutils  import copy
import  xlwt
def  read__excel():
    data=xlrd.open_workbook(r'C:\Users\行掌天下\Desktop\新建 XLS 工作表.xls')  #打开excel文件
    table=data.sheets()[0]  # 打开第一张表
    nrows=table.nrows   # 获取表的行数
    for i in range(nrows):      # 循环逐行打印
        if  i==0:    # 跳过第一行
            continue
        # print(table.row_values(i)[:2])   # 取前两列
        print(tuple(table.row_values(i)))
    return  0
def  write__excel():
    rb = xlrd.open_workbook(r'C:\Users\行掌天下\Desktop\新建 XLS 工作表.xls')  # 打开excel文件
    wb=copy.copy(rb)
    ws=wb.get_sheet(0)   #获取sheet对象
    ws.write(11,10,'hello')
    wb.save(r'C:\Users\行掌天下\Desktop\新建 XLS 工作表.xls')
    return  0
def  Create_Excel_Write():
    # 创建workbook和sheet对象
    workbook = xlwt.Workbook()  # 注意Workbook的开头W要大写
    sheet1 = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
    sheet2 = workbook.add_sheet('sheet2', cell_overwrite_ok=True)
    # 向sheet页中写入数据
    sheet1.write(0, 0, '你好 ')
    sheet1.write(0, 1, 'aaaa')
    sheet2.write(0, 0, 'should ')
    sheet2.write(1, 2, 'bbbbb')
    workbook.save(r'C:\Users\行掌天下\Desktop\xxx.xls')
    return 0
if __name__=='__main__':
    # write__excel()    #写入数据
    # read__excel()     #读取文件
    Create_Excel_Write()  #创建并写入数据
