import os
import pandas as pd

# 定义文件夹路径
folder_path = r"文件夹路径"

# 获取所有子文件夹路径
subfolders = [f.path for f in os.scandir(folder_path) if f.is_dir()]

# 为每个Excel表格添加表头并保存
for subfolder in subfolders:
    # 获取该子文件夹中所有Excel表格的路径
    excel_paths = [f.path for f in os.scandir(subfolder) if f.is_file() and f.name.endswith(".xlsx")]
    for excel_path in excel_paths:
        # 读取Excel表格
        df = pd.read_excel(excel_path, header=None)
        # 添加表头
        df.columns = ['经度', '纬度']
        # 保存Excel表格
        df.to_excel(excel_path, index=False)