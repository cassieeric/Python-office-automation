# 将A文件中名为a的sheet和B文件中名为b的sheet合并到一个sheet中去
from pathlib import Path
import pandas as pd

path = r'E:\PythonCrawler\有趣的代码\Python自动化办公\将A文件中名为a的sheet和B文件中名为b的sheet合并到一个sheet中去'
data_ex1 = pd.read_excel('ex1.xlsx', sheet_name='df1')
data_ex2 = pd.read_excel('ex2.xlsx', sheet_name='df2')
result = pd.concat([data_ex1, data_ex2], ignore_index=True)
result.to_excel('将A文件中名为a的sheet和B文件中名为b的sheet合并到一个sheet中去.xlsx', index=False, encoding='utf-8')
print('添加和合并完成！')
