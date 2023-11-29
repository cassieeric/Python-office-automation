# 该代码是小小明大佬优化瑜亮老师的代码得到后的优化代码
import pandas as pd


def filter_rows(group):
    diff = group.结束时间.diff()
    mask = diff.dt.total_seconds() < 20
    return group[~mask].drop_duplicates(keep="first")


df = pd.read_excel("工作量计算.xlsx")
res = df.sort_values(["编号", "环节", "审核人", "金额", "结束时间"]).groupby(["编号", "环节", "审核人", "金额"], as_index=False).apply(filter_rows).droplevel(0)
print(res)

