#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
from copy import copy
from openpyxl import load_workbook
import re
from docxtpl import DocxTemplate
import math
from glob import glob

from win32com import client as win32
import os


src_dir = "."
os.makedirs("result", exist_ok=True)
tpl_file_path1 = "合格验收证书模板.docx"
tpl_file_path2 = "到货证明模板.docx"
df = pd.read_excel(f"{src_dir}/小站订单20230323.xlsx", sheet_name='佳贤', usecols="A:H")
df = df.dropna(subset="订单编号").ffill().set_index("序号")
df.tail()


# In[2]:


word = win32.Dispatch("Word.Application")
word.Visible = True
word.ScreenUpdating = False
word_data = {}
files = glob(f"{src_dir}/物资合同/物资合同*.doc*")
print("加载word数据")
first = None
for file in files:
    filepath = os.path.abspath(file)
    print(filepath)
    rDoc = word.Documents.Open(filepath)
    header, code = None, None
    data = []
    addr = ""
    for row in rDoc.Tables(1).Rows:
        row = [c.Range.Text.strip("\r\x07") for c in row.Cells]
        if header is None:
            if row[0] == "序号":
                header = row
            elif len(row) == 2 and row[1].startswith("供应商"):
                supplier = row[1].split("：")[1]
            elif len(row) > 2:
                if row[2].startswith("送货地址"):
                    addr = row[2]
                elif row[2].startswith("收货人"):
                    addr += "\n"+row[2]
                elif row[2].startswith("电话"):
                    addr += "\n"+row[2]
                elif row[1].startswith("发货通知单编号"):
                    code = row[1].split("：")[1]
        else:
            if row[0] == "货物费合计金额：":
                break
            data.append(row)
    data = pd.DataFrame(data, columns=header)
    print(code, addr)
    data["供应商"] = supplier
    data.采购数量 = pd.to_numeric(data.采购数量)
    data.价款小计 = pd.to_numeric(data.价款小计)
#     word_data[code] = data.iloc[:, [2, 1, 12, 5, 7, 8]]
    word_data[code] = [supplier, addr, data]
#     display(word_data[code])
    if first is None:
        first = rDoc
    else:
        rDoc.Close()
word.ScreenUpdating = True
first.Close()
word.Quit()


# In[3]:
'''
佳贤大概是这样算的：
包装类型好像都是：纸箱
件数：基带单元/扩展单元 都是1套1箱，射频单元是5套1箱
安装套件件数写：随主设备包装
GPS 件数6套1箱，
线材都是100根一箱
光模块，120对或者240个一箱
功分器是10个一箱
'''



def get_num(wdf):
    nums = []
    for name, num, unit in zip(wdf["货物名称（物料名称）"], wdf["采购数量"], wdf["计量单位"]):
        # t = None  # 如果设置成这个，整数存在空值就会成为浮点数，生成的到货中的件数是浮点数
        t = pd.NA  # 如果设置成这个，nullable类型的整数允许存在空值，生成的到货中的件数是整数
        if "安装套件" in name:
            # t = "随主设备包装"
            t = ""
        elif re.search("(基带单元|扩展单元|BBU交转直流电)", name):
            t = num
        elif "射频单元" in name:
            t = math.ceil(num/5)
        elif "GPS" in name:
            t = math.ceil(num/6)
        elif re.search("(尾纤|接地线|交流电源线)", name):
            t = math.ceil(num/100)
        elif "光模块" in name:
            t = math.ceil(num/(120 if unit == "个" else 240))
        elif "功分器" in name:
            t = math.ceil(num/10)
        nums.append(t)
    wdf["件数"] = nums
    return wdf


def get_text(wdf):
    text = ""
    flag = wdf["货物编码（物料编码）"].isin(["J01140300001", "J01140300003"]).any()
    if flag:
        text += "本订单中"
    if "J01140300001" in wdf["货物编码（物料编码）"].values:
        text += "5G小基站-自研EXT型-基带单元-2×100M-2TR-TDD J01140300001，需要供应商（乙方）装配NR基带加速卡升级成5G小基站-自研EXT型-基带单元-4×100M-2TR-TDD J01140300002，"
    if "J01140300003" in wdf["货物编码（物料编码）"].values:
        text += "5G小基站-自研EXT型-基带单元-2×100M+2×20M-2TR J01140300003，需要供应商（乙方）装配NR基带加速卡升级成5G小基站-自研EXT型-基带单元-4×100M+2×20M-2TR J01140300004，"
    if flag:
        text += "NR基带单元加速卡配送发货。"
    return text


# In[4]:


for i, s in df.loc[251:258].iterrows():
    tpl = DocxTemplate("合格验收证书模板.docx")
    code = s["订单编号"]
    supplier, addr, wdf = word_data[code]
    context = {"supplier": supplier}
    context["num"] = s["合同编号"]
    context["code"] = code
#     print(s.to_dict())
    wdf = wdf.iloc[:, [2, 1, 3, 5, 7, 4]].copy()
    wdf.原厂商规格型号 = wdf.原厂商规格型号.replace("-", "/")
    t1 = "装配NR基带加速卡升级成5G小基站-自研EXT型-基带单元-4×100M-2TR-TDD J01140300002"
    t2 = "装配NR基带加速卡升级成5G小基站-自研EXT型-基带单元-4×100M+2×20M-2TR J01140300004"
    wdf.备注 = wdf["货物编码（物料编码）"].map(
        {"J01140300001": t1, "J01140300003": t2}).fillna("")
    context["data"] = wdf.values.tolist()
    context["m"] = wdf.采购数量.sum()
    tpl.render(context, autoescape=True)
    save_name = f"result/{s['城市']}-{code}合格验收证书.docx"
    tpl.save(save_name)
    print("保存到", save_name)

    tpl = DocxTemplate("到货证明模板.docx")
    supplier, addr, wdf = word_data[code]
    context["addr"] = addr
    wdf = wdf.iloc[:, [0, 2, 1, 3, 7, 5]].copy()
    wdf.原厂商规格型号 = wdf.原厂商规格型号.replace("-", "/")
    wdf["包装类型"] = "纸箱"
    get_num(wdf)
    wdf.loc[wdf.件数.isnull(), "包装类型"] = "/"
    wdf.件数.fillna(wdf.采购数量, inplace=True)
    context["data"] = wdf.values.tolist()
    context["text"] = get_text(wdf)
    tpl.render(context, autoescape=True)
    save_name = f"result/{s['城市']}-{code}到货证明.docx"
    tpl.save(save_name)
    print("保存到", save_name)
    print()


# In[ ]:




