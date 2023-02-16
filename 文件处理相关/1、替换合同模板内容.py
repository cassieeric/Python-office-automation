import docx


def replace_data(batch, province, price, contract_number):
    for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                if batch_placeholder in run.text:
                    run.text = run.text.replace(batch_placeholder, batch)

                if province_placeholder in run.text:
                    run.text = run.text.replace(province_placeholder, province)

                if price_placeholder in run.text:
                    run.text = run.text.replace(price_placeholder, price)

                if contract_number_placeholder in run.text:
                    run.text = run.text.replace(contract_number_placeholder, contract_number)
    print("替换完成！")
    doc.save(f"{batch}.docx")


if __name__ == '__main__':
    template_docx_file = "2、小基站收入订单合同.docx"
    doc = docx.Document(template_docx_file)
    batch_placeholder = "<batch>"
    province_placeholder = "<province>"
    price_placeholder = "<price>"
    contract_number_placeholder = "<contract_number>"
    # replace_data('上海第一批', '上海', '421811.63', 'HQGF02200541BGY00-SH00001')
    # replace_data('广东惠州第五批', '广东', '274881.54', 'HQGF02200541BGY00-GD00040')
    # replace_data('广东惠州第六批', '广东', '412535.88', 'HQGF02200541BGY00-GD00041')

    # replace_data('福建龙岩第一批', '福建', '471888', 'HQGF02200541BGY00-FJ00009')
    # replace_data('福建龙岩第二批', '福建', '82642.55', 'HQGF02200541BGY00-FJ00008')
    # replace_data('福建龙岩第三批', '福建', '165285.1', 'HQGF02200541BGY00-FJ00007')
    # replace_data('福建厦门第一批', '福建', '756200.52', 'HQGF02200541BGY00-FJ00011')

    # replace_data('广东深圳第七批', '广东', '664608.37', 'HQGF02200541BGY00-GD00042')
    # replace_data('广东深圳第五批', '广东', '554738.47', 'HQGF02200541BGY00-GD00017')
    # replace_data('广东深圳第四批', '广东', '535076.47', 'HQGF02200541BGY00-GD00014')

    # replace_data('广东东莞第六批', '广东', '9501.04', 'HQGF02200541BGY00-GD00032')
    # replace_data('广东东莞第五批', '广东', '703054.36', 'HQGF02200541BGY00-GD00013')
    # replace_data('广东东莞第四批', '广东', '703054.36', 'HQGF02200541BGY00-GD00010')
    # replace_data('广东东莞第三批', '广东', '797217.26', 'HQGF02200541BGY00-GD00012')

    # replace_data('广东汕头第二批', '广东', '736196.13', 'HQGF02200541BGY00-GD00016')

    # replace_data('广东汕尾第二批', '广东', '2149.26', 'HQGF02200541BGY00-GD00033')

    # replace_data('福建泉州第一批', '福建', '209194.64', 'HQGF02200541BGY00-FJ00003 ')
    # replace_data('福建泉州第二批', '福建', '176605.44', 'HQGF02200541BGY00-FJ00001 ')
    # replace_data('福建泉州第三批', '福建', '187325.75', 'HQGF02200541BGY00-FJ00004 ')
    # replace_data('福建泉州第四批', '福建', '928505.18', 'HQGF02200541BGY00-FJ00005 ')
    # replace_data('福建泉州第五批', '福建', '147505.68', 'HQGF02200541BGY00-FJ00002 ')

    replace_data('江苏南通第二批', '江苏', '247772.84', 'HQGF02200541BGY00-JS00002')


