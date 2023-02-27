from win32com.client import constants, gencache
import os


# Word转pdf方法,第一个参数代表word文档路径，第二个参数代表pdf文档路径
def Word_to_Pdf(Word_path, Pdf_path):
    word = gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(Word_path, ReadOnly=1)
    # 转换方法
    doc.ExportAsFixedFormat(Pdf_path, constants.wdExportFormatPDF)
    word.Quit()


# # 调用方法，进行单个文件的转换
# Word_to_Pdf('E:/B_pycharm/python_word_create.docx','E:/B_pycharm/python_word_create.pdf')
# 多个文件的转换
# print(os.listdir('.')) # 当前文件夹下的所有文件
Word_files = []
for file in os.listdir('.'):
    # 找出所有后缀为doc或者docx的文件
    # if file.endswith(('.doc','.docx')):
    if file.endswith('.doc'):
        Word_files.append(file)
print(Word_files)
for file in Word_files:
    file_path = os.path.abspath(file)
    index = file_path.rindex('.')
    pdf_path = file_path[:index] + '.pdf'
    print(file_path)
    print(pdf_path)
    Word_to_Pdf(file_path, pdf_path)
