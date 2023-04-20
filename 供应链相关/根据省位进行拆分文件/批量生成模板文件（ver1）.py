# import os
# import shutil
#
#
# # 指定原始文件夹的路径
# src_folder = os.path.abspath(r'C:\Users\pdcfi\Desktop\验收材料-文件夹模板\验收材料')
#
# # 指定复制后的文件夹的名称前缀
# # dst_folder_prefix = "_"
# dst_folder = os.path.abspath(r'C:\Users\pdcfi\Desktop\整理好的\模板')
# # 指定需要复制的次数
# num_copies = 3
#
# # 循环复制文件夹和文件
# for i in range(num_copies):
#     # 构造新的文件夹名称
#     dst_folder = str(i)
#     if not os.path.exists(dst_folder):
#         os.makedirs(dst_folder)
#     # 使用shutil复制文件夹和文件
#     print(f"{i}")
#     shutil.copytree(src_folder, dst_folder)
#

import pandas as pd
import os
import shutil
src_folder = os.path.abspath(r'C:\Users\pdcfi\Desktop\验收材料-文件夹模板\验收材料\模板')
folder = os.path.abspath(r'C:\Users\pdcfi\Desktop\陕西省')

df = pd.read_excel(f"个人台账.xlsx", usecols="A:C")

for i in df["销售订单合同编号"]:
    print(i)
    dst_folder = folder + "\\" + str(i.strip())
    shutil.copytree(src_folder, dst_folder)
