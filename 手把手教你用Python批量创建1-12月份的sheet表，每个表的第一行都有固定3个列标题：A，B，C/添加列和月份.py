# coding: utf-8
import pandas as pd
import openpyxl

df = pd.DataFrame({'A': [], 'B': [], 'C': []})
for year in range(1999, 2022):
    path_name = f'./{year}年.xlsx'
    with pd.ExcelWriter(path_name, engine='openpyxl', mode='w+') as writer:
        for month in range(1, 13):
            df.to_excel(writer, index=False, sheet_name=f'{month}月份')
print('文件生成完成')
