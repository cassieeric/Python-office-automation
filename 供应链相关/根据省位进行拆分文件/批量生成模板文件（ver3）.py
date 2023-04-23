import pandas as pd
import numpy as np
import os
import shutil

src_folder = os.path.abspath(r'D:\供应链\成果平台资料录入\根据省位进行拆分文件\验收材料-文件夹模板\验收材料\模板')
df_1 = pd.read_excel("双路耦合器订单20230331（销售订单和采购订单对应）.xlsx", usecols="B:E")
provinces = df_1["省分"].drop_duplicates()
for province in provinces:
    if pd.isna(province):
        print("该字段为空，不创建文件")
        pass
    else:
        folder = os.path.abspath(rf'D:\供应链\成果平台资料录入\根据省位进行拆分文件\res\{province}')
        df = pd.read_excel(f"./res/{province}.xlsx", usecols="A:C")
        hetong_number = df["销售订单合同编号"].drop_duplicates()
        for i in hetong_number:
            if pd.isna(i):
                print("该字段为空，不创建文件")
                pass
            else:
                print(i)
                dst_folder = folder + "\\" + str(i).strip()
                shutil.copytree(src_folder, dst_folder)

