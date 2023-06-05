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
tpl_file_path1 = "证书模板.docx"
tpl_file_path2 = "证明模板.docx"
df = pd.read_excel(f"{src_dir}/xxx.xlsx", sheet_name='xxx', usecols="A:H")
df = df.dropna(subset="订单编号").ffill().set_index("序号")
df.tail()


# In[2]:


word = win32.Dispatch("Word.Application")
word.Visible = True
word.ScreenUpdating = False
word_data = {}
files = glob(f"{src_dir}/物资合同*.doc*")
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
                elif row[1].startswith("编号"):
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


def get_num(wdf):
    nums = []
    for name, num, unit in zip(wdf["货物名称"], wdf["采购数量"], wdf["计量单位"]):
        # t = None  # 如果设置成这个，整数存在空值就会成为浮点数，生成的到货中的件数是浮点数
        t = pd.NA  # 如果设置成这个，nullable类型的整数允许存在空值，生成的到货中的件数是整数
        if "套件" in name:
            # t = "随设备包装"
            t = ""
        elif re.search("(交转直流电)", name):
            t = num
        elif "单元" in name:
            t = math.ceil(num/5)
        elif "GPS" in name:
            t = math.ceil(num/6)
        elif re.search("(电源线)", name):
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
    flag = wdf["货物编码"].isin(["J01140300001", "J01140300003"]).any()
    if flag:
        text += "本订单中"
    if "0001" in wdf["货物编码"].values:
        text += "xxx"
    if "0003" in wdf["货物编码"].values:
        text += "xxx"
    if flag:
        text += "xxxxxx"
    return text


# In[4]:


for i, s in df.loc[261:262].iterrows():
    tpl = DocxTemplate("证书模板.docx")
    code = s["编号"]
    supplier, addr, wdf = word_data[code]
    context = {"supplier": supplier}
    context["num"] = s["编号"]
    context["code"] = code
#     print(s.to_dict())
    wdf = wdf.iloc[:, [2, 1, 3, 5, 7, 4]].copy()
    wdf.原厂商规格型号 = wdf.原厂商规格型号.replace("-", "/")
    t1 = "0002"
    t2 = "0004"
    wdf.备注 = wdf["物料编码"].map(
        {"0001": t1, "0003": t2}).fillna("")
    context["data"] = wdf.values.tolist()
    context["m"] = wdf.采购数量.sum()
    tpl.render(context, autoescape=True)
    save_name = f"result/{s['城市']}-{code}证书.docx"
    tpl.save(save_name)
    print("保存到", save_name)

    tpl = DocxTemplate("证明模板.docx")
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
