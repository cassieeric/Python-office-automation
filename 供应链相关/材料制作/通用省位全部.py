from docxtpl import DocxTemplate
import pandas as pd
import os

df = pd.read_excel("20230511.xlsx", sheet_name="Sheet4", usecols="B,E,F,M,Q,R,U,V,X,AN:AR,AX")
df.columns = [c.strip() for c in df.columns]
for (b, e, m), df_split in df.groupby(['单号', '编号', '单位']):
    tpl1 = DocxTemplate('1、证明-模板.docx')
    tpl2 = DocxTemplate('2、证明-模板.docx')
    tpl3 = DocxTemplate('4、证书-模板.docx')
    m = m.strip("*")
    # aq = "、".join(df_split.项目编码.unique())
    # aq = "、".join(df_split.项目编码.astype("string").unique())
    aq = "、".join(df_split.项目编码.dropna().astype("string").unique())
    # ar = "、".join(df_split.项目名称.unique())
    # ar = "、".join(df_split.项目名称.astype("string").unique())
    ar = "、".join(df_split.项目名称.dropna().astype("string").unique())
    an = df_split.送货地址.iat[0]
    ao = df_split.接货人.iat[0]
    ap = df_split.接货人联系电话.iat[0]
    ax = df_split.通知编号.iat[0]
    print(b, e, m, aq, an, ao, ap, ax, ar)
    title = "协议"
    items1 = []
    context1 = {"title": title, "B": b, "E": e, "M": m, "items": items1}
    items2 = []
    context2 = {"B": b, "E": e, "M": m,
                "AQ": aq, "AN": an, "AO": ao, "AP": ap, "AX": ax, "AR": ar,
                "items": items2}
    items3 = []
    context3 = {"title": title, "B": b, "E": e, "M": m, "items": items3}
    name = df_split.城市名.iat[0]
    try:
        province = name[:2]
        city = name[2:].rstrip("0123456789")
    except:
        province = name
        city = "暂无命名"
    os.makedirs(f"/{province}/{city}/{name}", exist_ok=True)
    t = df_split.groupby(["编号", "编码", "名称", "型号", "计量单位"]).数量.sum()
    for (e, q, r, u, v), x in t.items():
        items1.append([e, r, u, v, int(x)])
        items2.append([q, r, u, int(x)])
        items3.append([e, r, u, v, int(x)])
    tpl1.render(context1, autoescape=True)
    tpl1.save(f'/{province}/{city}/{name}/1、证明-{name}-{e}.docx')
    tpl2.render(context2, autoescape=True)
    tpl2.save(f'/{province}/{city}/{name}/2、证明-{name}-{e}.docx')
    tpl3.render(context3, autoescape=True)
    tpl3.save(f'/{province}/{city}/{name}/4、证书-{name}-{e}.docx')
    print(f'/{province}/{city}/{name}/{name}-{e}.docx')

