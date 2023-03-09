# 这里使用Pandas库进行实现
import pandas as pd

df = pd.read_excel('res.xlsx')
# df.set_index(["A"]).reset_index()
for i in range(len(df) // 3 + 1):
    df.iloc[3 * i: 3 * (i + 1)].to_excel(f'{i}.xlsx')
