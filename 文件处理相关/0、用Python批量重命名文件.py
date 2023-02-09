import os
# 定义来源文件夹
path_src = r'C:\Users\pdcfi\Desktop\物资开箱报告_第二批_佳贤_中信科\佳贤'
# 定义目标文件夹
path_dst = r'C:\Users\pdcfi\Desktop\new'
# 自定义格式，例如“报告-第X份”，第一个{}用于放序号，第二个{}用于放后缀
rename_format = '_电子签'
begin_num = 1


def doc_rename(path_src, path_dst, begin_num):
    for i in os.listdir(path_src):
        print(f'正在重命名第{begin_num}个文件 >> {i}')
        # 获取原始文件名
        doc_src = os.path.join(path_src, i)
        # 重命名
        doc_name = os.path.splitext(i)[0] + rename_format + os.path.splitext(i)[-1]
        # 确定目标路径
        doc_dst = os.path.join(path_dst, doc_name)
        begin_num += 1
        os.rename(doc_src, doc_dst)


# 运行函数
doc_rename(path_src, path_dst, begin_num)

