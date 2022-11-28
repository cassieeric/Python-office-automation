# 参考博客：https://blog.csdn.net/weixin_44476410/article/details/116063808
import shutil
import os


def copy_file(path):
    # （root，dirs，files）分别为：遍历的文件夹，遍历的文件夹下的所有文件夹，遍历的文件夹下的所有文件
    for root, dirs, files in os.walk(path):
        for file in files:
            if "合格验收" in file:
                shutil.copyfile(root + '\\' + file, target_path + '\\' + file)
                print(root + '\\' + file + ' 复制成功-> ' + target_path)
        for dir_in in dirs:
            copy_file(dir_in)


if __name__ == '__main__':
    # 文件夹路径
    source_path = r'D:\供应链\订单&需求单'
    # 输出路径
    target_path = r'C:\Users\pdcfi\Desktop\test\res'
    copy_file(source_path)


