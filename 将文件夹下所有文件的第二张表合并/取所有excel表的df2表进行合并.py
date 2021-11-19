# 合并所有表格中的第二张表格
from pathlib import Path
import pandas as pd

path = Path(r'E:\PythonCrawler\有趣的代码\Python自动化办公\将文件夹下所有文件的第二张表合并')
data_list = []
for i in path.glob("*.xls*"):
    # data = pd.read_excel(i, sheet_name='df2')
    data = pd.read_excel(i, sheet_name=1)
    data_list.append(data)

result = pd.concat(data_list, ignore_index=True)
result.to_excel(path.joinpath('取所有excel表的df2表进行合并.xlsx'), index=False, encoding='utf-8')
print('添加和合并完成！')


