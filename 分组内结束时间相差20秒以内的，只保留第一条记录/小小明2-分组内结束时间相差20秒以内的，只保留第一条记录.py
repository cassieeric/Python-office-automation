import pandas as pd


def func(df_split):
    last_time = None
    idx = []
    for row in df_split.itertuples():
        if last_time is None or (row.结束时间 - last_time).total_seconds() > 20:
            idx.append(row.Index)
            last_time = row.结束时间
    return df_split.loc[idx]


df = pd.read_excel("工作量计算.xlsx", index_col=None)
res = (df.sort_values(["编号", "环节", "审核人", "金额", "结束时间"]).groupby(["编号", "环节", "审核人", "金额"], as_index=False).apply(func).droplevel(0))
print(res)
