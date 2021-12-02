from shutil import rmtree
import zipfile
from copy import deepcopy
from pathlib import Path
from win32com import client as wc  # pip install pypiwin32
import pandas as pd


# doc文档不包含所需xml文件，使用win32com转存为docx文档
def doctransform2docx(doc_path):
    docx_path = doc_path + 'x'
    suffix = doc_path.split('.')[1]
    assert 'doc' in suffix, '传入的不是word文档，请重新输入！'
    if suffix == 'docx':
        return Path(doc_path)
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(doc_path)
    doc.SaveAs2(docx_path, 16)  # docx为16
    doc.Close()
    word.Quit()
    return Path(docx_path)


# docx文档解压
def docx_unzip(docx_path):
    docx_path = Path(docx_path) if isinstance(docx_path, str) else docx_path
    upzip_path = docx_path.with_name(docx_path.stem)
    with zipfile.ZipFile(docx_path, 'r') as f:
        for file in f.namelist():
            f.extract(file, path=upzip_path)
    xml_path = upzip_path.joinpath('word/document.xml')
    with xml_path.open(encoding='utf-8') as f:
        xml_file = f.read()
    return upzip_path, xml_path, xml_file


# 将文件夹中的所有文件压缩成docx文档
def docx_zipped(docx_path, zipped_path):
    docx_path = Path(docx_path) if isinstance(docx_path, str) else docx_path
    with zipfile.ZipFile(zipped_path, 'w', zipfile.zlib.DEFLATED) as f:
        for file in docx_path.glob('**/*.*'):
            f.write(file, file.as_posix().replace(docx_path.as_posix() + '/', ''))


# 删除生成的解压文件夹
def remove_folder(path):
    path = Path(path) if isinstance(path, str) else path
    if path.exists():
        rmtree(path)
    else:
        raise "系统找不到指定的文件"


# 替换docx中的特定字符，重新保存document.xml至需要压缩的目录下
def replace_docx(name, values, xml_file, xml_path, unzip_path, path_name='Company'):
    xml_path = Path(xml_path) if isinstance(xml_path, str) else xml_path
    xml_file_copy = deepcopy(xml_file)  # 深复制xml内容
    for col_name, value in zip(name, values):
        if col_name == 'Company':
            path_name = str(value)
        xml_file_copy = xml_file_copy.replace(col_name, str(value))
    with xml_path.open(mode='w', encoding='utf-8') as f:
        f.write(xml_file_copy)
    # xml文档替换完毕，通过zipfile重新压缩另存为docx文档
    docx_zipped(unzip_path, f'{save_folder}/{path_name}.docx')


if __name__ == '__main__':
    # 定义需处理的文件路径
    doc_path = r"C:\Users\pdcfi\Desktop\yueliang\单位.doc"
    excel_path = r"C:\Users\pdcfi\Desktop\yueliang\信息.xls"
    save_folder = Path(r'C:\Users\pdcfi\Desktop\yueliang')
    save_folder.mkdir(parents=True, exist_ok=True)  # 文件夹没有时自动创建

    # 获取excel数据
    data = pd.read_excel(excel_path)
    docx_path = doctransform2docx(doc_path)
    unzip_path, xml_path, xml_file = docx_unzip(docx_path)
    data_save = data.apply(lambda x: replace_docx(x.index, x.values, xml_file, xml_path, unzip_path), axis=1)
    remove_folder(unzip_path)
