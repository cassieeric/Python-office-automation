import pandas as pd
import numpy as np
import os
import shutil

src_folder = os.path.abspath(r'C:\Users\pdcfi\Desktop\验收材料-文件夹模板\验收材料\模板')
for province in ["广东", "安徽", "湖南", "江苏"]:
    folder = os.path.abspath(rf'C:\Users\pdcfi\Desktop\{province}')

    df = pd.read_excel(f"{province}.xlsx", usecols="A:D")
    hetong_number = df["销售订单合同编号"].drop_duplicates()
    for i in hetong_number:
        if pd.isna(i):
            print("该字段为空，不创建文件")
            pass
        else:
            print(i)
            dst_folder = folder + "\\" + str(i).strip()
            shutil.copytree(src_folder, dst_folder)