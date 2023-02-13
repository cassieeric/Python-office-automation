# 参考博客：https://blog.csdn.net/weixin_44476410/article/details/116063808
import shutil
import os


def delete_file(path):
    # （root，dirs，files）分别为：遍历的文件夹，遍历的文件夹下的所有文件夹，遍历的文件夹下的所有文件
    for root, dirs, files in os.walk(path):
        for file in files:
            if "_双章" in file:  # 多了一层限定条件
            # if ".xls" in file:
                os.remove(file)
                print(f'{file} 文件删除成功')
        # for dir_in in dirs:
        #     copy_file(dir_in)


if __name__ == '__main__':
    # 文件夹路径
    source_path = r'D:\供应链\订单&需求单'
    # 输出路径
    target_path = r'C:\Users\pdcfi\Desktop\待制作开箱、到货、终验证书的小站订单'
    delete_file(target_path)


