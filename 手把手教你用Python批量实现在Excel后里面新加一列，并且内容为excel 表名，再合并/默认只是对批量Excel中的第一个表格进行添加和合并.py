from pathlib import Path
import numpy as np
import pandas as pd

path = r'E:\PythonCrawler\python_crawler-master\MergeExcelSheet\file\777'
excel_list = [(i.stem, pd.read_excel(i, sheet_name=None).values()) for i in Path(path).glob("*.xls*")]
data_list = []
for name, data in excel_list:
    print(name)
    print(data)
    data['表名'] = name
    data_list.extend(data)
result = pd.concat(data_list, ignore_index=True)
result.to_excel('result.xlsx', index=False, encoding='utf-8')
print('添加和合并完成！')

