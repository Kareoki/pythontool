# -*- coding: utf-8 -*-
import win32com.client as win32
import sys


excel = win32.gencache.EnsureDispatch("Excel.Application")
wb = excel.Workbooks.Open("C:\\Users\\Administrator\\Desktop\\python_excel.xlsx")
excel.Visible = True
ws = wb.Worksheets("工作表1")


f = open(r'C:\\Users\\Administrator\\Desktop\\Python_reading.txt')
s = open('save.txt','a+')
s.readline()
text = list()
for line in f:
    str_list = line.split(',')
  
    if str_list[0] == "-i":
        ws.Range(str_list[1]).Value = str_list[2]
        s.write(str_list[1]+"儲存成功"+"\n")
        wb.Save()
    
    elif str_list[0] == "-r":
         s.write(str_list[1]+"為"+ws.Range(str_list[1]).Value+"\n")
        
  
    else:
        print("wrong")

        
#excel 欄位取值
        #ws.Range(A1.Value)
