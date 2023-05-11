import os
import pandas as pd

# 定义文件夹路径
folder_path = r"文件夹路径"

# 获取文件夹下的所有子文件夹
subfolders = [f.path for f in os.scandir(folder_path) if f.is_dir()]

# 遍历每个子文件夹
for subfolder in subfolders:
    # 获取Excel文件路径
    excel_file = os.path.join(subfolder, "Excel表名.xlsx")

    # 读取Excel文件
    df = pd.read_excel(excel_file, header=None)

    # 添加表头
    df.columns = ["经度", "纬度"]

    # 保存Excel文件
    df.to_excel(excel_file, index=False)