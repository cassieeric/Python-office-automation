
import os
import pandas as pd

# 读取目标文件夹及子文件夹下的所有Excel文件
folder_path = r'C:\Users\pdcfi\Desktop\新建文件夹'
excel_files = []
for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.endswith('.xlsx') or file.endswith('.xls'):
            excel_files.append(os.path.join(root, file))

# 循环读取每个Excel并添加表头
for file_path in excel_files:
    df = pd.read_excel(file_path)  # 读取Excel
    df.columns = ['经度', '纬度']  # 添加表头
    df.to_excel(file_path, index=False)  # 写入Excel
