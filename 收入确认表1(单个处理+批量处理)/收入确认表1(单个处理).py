import openpyxl

workbook1 = openpyxl.load_workbook("上市公司收入确认表(研究院)_模板.xlsx")
worksheet1 = workbook1.worksheets[0]
print(worksheet1['C4'].value)  # # 含税金额：434704.22
print(worksheet1['D4'].value)  # 分公司：中国电信股份有限公司{}分公司
print(worksheet1['F4'].value)  # 销售订单编号：HIDD202209230175

workbook2 = openpyxl.load_workbook("小站订单.xlsx")
worksheet2 = workbook2['省对研究院']
print(worksheet2['C3'].value)  # 城市：海南1
print(worksheet2['D3'].value)  # 销售订单编号：HIDD202209230175
print(worksheet2['CU3'].value)  # 含税金额：434704.22
print(worksheet2['DM3'].value)  # 分公司：海南

print(f"正在处理订单：{worksheet2['C3'].value}...")
worksheet1['C4'].value = worksheet2['CU3'].value
worksheet1['D4'].value = f"中国电信股份有限公司{worksheet2['DM3'].value}分公司"
worksheet1['F4'].value = worksheet2['D3'].value
new_file_name = f"上市公司收入确认表(研究院)  （{worksheet2['C3'].value} {worksheet2['D3'].value}）"
workbook1.save(new_file_name + '.xlsx')
print(f"订单：{worksheet2['C3'].value}处理完成")
