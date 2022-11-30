import pandas as pd
import os

path = r"./新建文件夹/"
# 获取文件夹下的所有文件名
name_list = os.listdir(path)
name_list = pd.DataFrame(name_list)

# 计数器
res = []

# for循环遍历读取
for i in range(len(name_list)):  # len(name_list)等于21
    df = pd.read_excel(path + name_list[0][i])
    print('文件{}读取完成!'.format(i))
    target_data = df[df['id'] == '58666']
    # print(target_data)
    res.append(target_data)

final_df = pd.concat(res)
final_df.to_excel("target.xlsx")
