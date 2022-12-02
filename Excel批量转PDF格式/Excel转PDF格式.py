# -*- coding:utf-8 -*-
import os
import re
from win32com.client import DispatchEx


def xls2pdf(filename):
    try:
        xlApp = DispatchEx("Excel.Application")
        # 后台运行, 不显示, 不警告
        xlApp.Visible = False
        xlApp.DisplayAlerts = 0
        books = xlApp.Workbooks.Open(filename, False)
        # 第一个参数0表示转换pdf
        books.ExportAsFixedFormat(0, re.subn('.xls', '.pdf', filename)[0])
        books.Close(False)
        print('保存 PDF 文件：', re.subn('.xls', '.pdf', filename)[0])
    except:
        input('转换出错了，按任意键退出')
    finally:
        xlApp.Quit()


if __name__ == '__main__':
    filepath = input('输入你的文件路径：')
    for dirs, subdirs, files in os.walk(filepath):
        for name in files:
            if re.search('.xls', name):
                xls2pdf(filepath + '\\' + name)
    input('转换成功，按任意键推出')
