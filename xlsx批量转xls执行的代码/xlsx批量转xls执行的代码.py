# -*- coding:utf-8 -*-
import os
import win32com.client as win32

# 输入目录
inputdir = r'C:\Users\pdcfi\Desktop\test\收入确认表\生成的收入确认表'
# 输出目录
outputdir = r'C:\Users\pdcfi\Desktop\test\收入确认表\生成的收入确认表\收入确认表xls文件'
if not os.path.exists(outputdir):
    os.mkdir(outputdir)

# 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
for parent, dirnames, filenames in os.walk(inputdir):
    for fn in filenames:
        if fn.split('.')[-1] == "xlsx":
            filedir = os.path.join(parent, fn)
            print(filedir)
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            wb = excel.Workbooks.Open(filedir)
            # xlsx: FileFormat=51
            # xls:  FileFormat=56
            wb.SaveAs((os.path.join(outputdir, fn.replace('xlsx', 'xls'))), FileFormat=56)
            wb.Close()
            excel.Application.Quit()
