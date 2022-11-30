import pandas as pd
import os


path = r"./新建文件夹/"
# 获取文件夹下的所有文件名
name_list = os.listdir(path)
# print(name_list)
# name_list = pd.DataFrame(name_list)
# file_path = [xxx, xxx, xxx, ......]

res = pd.read_excel(path+name_list[0])
res = res[res['id'] == '58666']

for file in name_list[1:]:
    temp = pd.read_excel(path+file)
    temp = temp[temp['id'] == '58666']
    res = pd.concat([res, temp], ignore_index=True)
res.to_excel('res.xlsx')
