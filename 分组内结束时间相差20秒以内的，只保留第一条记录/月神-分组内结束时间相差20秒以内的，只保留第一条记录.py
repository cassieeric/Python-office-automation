import pandas as pd


def func(date_s):
    # 筛选函数
    min_date = date_s.iloc[0]
    for num, i in enumerate(date_s):
        if num and (i - min_date).seconds <= 20:
            yield False
        else:
            min_date = i
            yield True


df = pd.read_excel("工作量计算.xlsx")
df.sort_values(["编号", "环节", "审核人", "金额", "结束时间"], inplace=True)
df = df[df.groupby(["编号", "环节", "审核人", "金额"])["结束时间"].transform(func)]
print(df)
