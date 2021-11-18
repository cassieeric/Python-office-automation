# 给每个excel中的sheet增加一列，值为excel名.xlsx
from pathlib import Path
import pandas as pd

path = Path(r'E:\PythonCrawler\python_crawler-master\MergeExcelSheet\file\777')
excel_list = [(i.stem, pd.concat(pd.read_excel(i, sheet_name=None))) for i in path.glob("*.xls*")]
data_list = []
for name, data in excel_list:
    print(name)
    print(data)
    data['表名'] = name
    data_list.append(data)
result = pd.concat(data_list, ignore_index=True)
result.to_excel(path.joinpath('给每个excel中的sheet增加一列，值为excel名.xlsx'), index=False, encoding='utf-8')
print('添加和合并完成！')
