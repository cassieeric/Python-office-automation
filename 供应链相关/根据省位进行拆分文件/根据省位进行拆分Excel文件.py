import pandas as pd

df = pd.read_excel("移频MIMO订单20230331（销售订单和采购订单对应）.xlsx", usecols="B:E")
provinces = df["省分"].drop_duplicates()
# print(len(provinces))
# print(provinces)
# for province in ["北京", "福建", "广东", "安徽", "江西", "江苏", "甘肃",
#                  "广西", "贵州", "海南", "河北", "河南", "湖北", "湖南",
#                  "辽宁", "内蒙古", "宁夏", "青海", "陕西", "山西", "山东",
#                  "浙江", "上海", "四川", "天津", "西藏", "云南", "新疆", "重庆"]:
for province in provinces:
    if pd.isna(province):
        print("该字段为空，不创建文件")
        pass
    else:
        print(f"正在导出{province}的数据...")
        target_data = df[df['省分'] == province]
        # new_cols = ['到货证明', '初验', '终验']
        new_cols = ['付款通知书', '发票', '其他']
        # target_data = target_data.reindex(columns=[*target_data.columns.tolist(), *new_cols])
        target_data[new_cols] = None
        target_data.drop("省分", axis=1, inplace=True)
        target_data.to_excel(f"./excel_res/{province}.xlsx", index=False)
        print(f"{province}的数据已经导出完成")
