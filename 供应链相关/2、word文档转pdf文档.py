# 方法一
from docx2pdf import convert
from os import walk

# 方法二
doc_path = r'C:\Users\pdcfi\Desktop\出入库资料的制作\佳贤出入库资料的制作'
for root, dirs, filenames in walk(doc_path):
    # 遍历当前文件名称、并校验是否是word文档
    for file in filenames:
        # if file.endswith(".doc") or file.endswith(".docx"):
        if file.startswith("附件2") and file.endswith(".docx"):
            convert(file)
            print("物资开箱报告word格式转pdf格式已完成！")
            # word_file = str(root + "\\" + file)
            # # 如果当前文件是word文档则调用word转换函数
            # word = Dispatch('Word.Application')
            # # 以Doc对象打开文件
            # doc_ = word.Documents.Open(word_file)
            # # 另存为pdf文件
            # doc_.SaveAs(word_file.replace(".docx", ".pdf"), FileFormat=17)
            # # 关闭doc对象
            # doc_.Close()
            # # 退出word对象
            # word.Quit()
print("物资开箱报告word格式转pdf格式已完成！")
