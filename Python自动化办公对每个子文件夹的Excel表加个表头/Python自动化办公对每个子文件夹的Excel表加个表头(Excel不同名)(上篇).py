import os
import pandas as pd
import glob

# 定义文件夹路径
folder_path = r"C:\Users\pdcfi\Desktop\新建文件夹"

# 获取文件夹下的所有子文件夹
subfolders = [f.path for f in os.scandir(folder_path) if f.is_dir()]

excel_paths = []
# 遍历每个子文件夹
for subfolder in subfolders:
    # 获取Excel文件路径
    # excel_file = os.path.join(subfolder, "Excel表名.xlsx")
    excel_paths.extend(glob.glob(subfolder + "/*.xlsx"))

for excel_file in excel_paths:
    # 读取Excel文件
    df = pd.read_excel(excel_file, header=None)

    # 添加表头
    df.columns = ["经度", "纬度"]

    # 保存Excel文件
    df.to_excel(excel_file, index=False)