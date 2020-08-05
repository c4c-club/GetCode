#!/usr/bin/python
# -*- coding: UTF-8 -*-
import os
import re
import xlwt

'''获取子目录名'''
def getDirName (rootdir,list_all):
    for  a,dirs,b  in os.walk(rootdir,topdown=False):
        for name in dirs:
            list_all.append(name)

'''正则提取番号'''
def filter(list_fc2,list_yes,list_no,list_all):
    for i in range(len(list_all)):
        searchObj1 = re.search(r'\w+\d+\w+\W\d+',list_all[i],re.I)
        if searchObj1:
            list_fc2.append(searchObj1.group())
        else:
            searchObj = re.search(r'\w+\W\d+',list_all[i],re.I)
            if searchObj:
                list_yes.append(searchObj.group())
            else:
                list_no.append(list_all[i])

'''整理番号格式'''
def change(list_yes,list_change):
    for i in range(len(list_yes)):
        searchObj1 = re.search(r'\w+',list_yes[i],re.I)
        a=searchObj1.group().upper()
        searchObj2 = re.search(r'\d+',list_yes[i],re.I)
        b = searchObj2.group()
        c = a+'-'+b
        list_change.append(c)

'''整理特殊番号(FC2)格式'''
def change_fc2(list_fc2,list_change):
    for i in range(len(list_fc2)):
        searchObj1 = re.search(r'\w+\d+\w+',list_fc2[i],re.I)
        a=searchObj1.group().upper()
        searchObj2 = re.search(r'\d\d\d+',list_fc2[i],re.I)
        b = searchObj2.group()
        c = a+'-'+b
        list_change.append(c)

'''写入Excel'''
def insertToExcel(list_change,list_no):
    work_book = xlwt.Workbook(encoding='utf-8',style_compression=0)
    sheet = work_book.add_sheet('sheet1', cell_overwrite_ok=True)
    heading = ['Complete','Error']
    sheet.write(0, 0, heading[0])
    sheet.write(0,1,heading[1])
    for i in range(len(list_change)):
        sheet.write(i+1, 0, list_change[i])
    for i in range(len(list_no)):
        sheet.write(i+1, 1, list_no[i])
    work_book.save('D:/pytest/GetCode.xls')#D:/pytest/GetCode.xls

def main():
    list_all = []
    list_yes = []
    list_no = []
    list_change = []
    list_fc2 = []
    rootdir = "D:/pytest"#D:/pytest

    getDirName(rootdir,list_all)
    filter(list_fc2,list_yes,list_no,list_all)
    change(list_yes,list_change)
    change_fc2(list_fc2,list_change)
    insertToExcel(list_change,list_no)

    print('转换完毕\n'+
          '已转换',len(list_change),'\n'+
          '失败',len(list_no),'\n'+
          '总计',len(list_change)+len(list_no),'\n'+
          '请按任意键退出')
    input()

if __name__ == "__main__":
    main()

