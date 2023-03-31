from glob import glob
from win32com.client import Dispatch
from win32com import client as win32
import os

# In[4]: Excel格式转pdf
xlApp = win32.Dispatch("Excel.Application")
xlApp.Visible = True
xlApp.ScreenUpdating = False
xlApp.DisplayAlerts = False

files = glob(f"./物资出库申请-*.xlsx")
print("加载excel结果数据")
first = None
try:
    for file in files:
        filepath = os.path.abspath(file)
        print(filepath)
        book = xlApp.Workbooks.Open(filepath, ReadOnly=1)
        book.ExportAsFixedFormat(0, filepath[:-4]+"pdf")
        print("保存到", filepath[:-4]+"pdf")
        book.Save()
        book.Close()
finally:
    xlApp.ScreenUpdating = True
    xlApp.Quit()
