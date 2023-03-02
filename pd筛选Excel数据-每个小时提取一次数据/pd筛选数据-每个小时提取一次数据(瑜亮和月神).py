# -*- coding: utf-8 -*-
# @Time : 2022/4/21 18:00
# @Author : 瑜亮
# @FileName : pd筛选数据-每个小时提取一次数据.py

import pandas as pd

excel_filename = '数据.xlsx'
df = pd.read_excel(excel_filename)
# print(df)

# 方法一：分别取日期与小时，按照日期和小时删除重复项
# df['day'] = df['SampleTime'].dt.day   # 提取日期列
# df['hour'] = df['SampleTime'].dt.hour     # 提取小时列
# df = df.drop_duplicates(subset=['day', 'hour'])   # 删除重复项
# print(df)

# 方法二：把日期中的分秒替换为0
# SampleTime_new = df['SampleTime'].map(lambda x: x.replace(minute=0, second=0))
# data = df[SampleTime_new.duplicated() == False]
# print(df)

# 方法三：对日期时间按照小时进行分辨
# SampleTime_new = df['SampleTime'].dt.floor(freq='H')
# df = df[SampleTime_new.duplicated() == False]
# print(df)

# 方法四：对日期时间按照小时进行分辨
# SampleTime_new = df['SampleTime'].dt.to_period(freq='H')
# df = df[SampleTime_new.duplicated() == False]
# print(df)

# 方法五：对日期时间进行重新格式，并按照新的日期时间删除重复项（会引入新列）
# df['new'] = df['SampleTime'].dt.strftime('%Y-%m-%d %H')
# df = df.drop_duplicates(subset=['new'])
# print(df)

# 把筛选结果保存为excel文件
df.to_excel('数据筛选结果2.xlsx')
