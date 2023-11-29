import pandas as pd

df = pd.read_excel("工作量计算.xlsx", index_col=None)

# # 按照"编号"、"环节"、"审核人"、"金额"分组，并对"结束时间"列做升序排列
df_spilt = df.groupby(["编号", "环节", "审核人", "金额"]).apply(lambda x: x.sort_values('结束时间', ascending=True))

df_spilt.reset_index(drop=True, inplace=True)
df_spilt['结束时间'] = pd.to_datetime(df_spilt['结束时间'])  # 转换为日期时间格式


def filter_rows(group):
    # 计算时间差，删除时间差小于20秒的记录，只保留第一条记录
    diff = group.groupby('编号')['结束时间'].diff()
    mask = (diff.dt.total_seconds() < 20)
    group = group[~mask].drop_duplicates(keep="first")
    return group


# 对每个分组中的'结束时间'列进行去重操作
result = df_spilt.groupby(['编号', '环节', '审核人', '金额']).apply(filter_rows)
# 重新设置索引
result.reset_index(drop=True, inplace=True)
# 输出结果
print(result)
