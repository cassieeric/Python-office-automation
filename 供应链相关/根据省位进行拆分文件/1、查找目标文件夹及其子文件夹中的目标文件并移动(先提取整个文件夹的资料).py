# # 参考博客：https://blog.csdn.net/weixin_44476410/article/details/116063808
import shutil
import os

# import sys  # 导入sys模块
# sys.setrecursionlimit(1000)  # 将默认的递归深度修改为3000


def copy_file(path):
    num = 1
    # （root，dirs，files）分别为：遍历的文件夹，遍历的文件夹下的所有文件夹，遍历的文件夹下的所有文件
    for root, dirs, files in os.walk(path):
        for dir in dirs:
            if "子文件夹1" in dir:
                shutil.copytree(root + '\\' + dir, target_path + '\\' + f"{dir}{num}")
            #     shutil.copytree(root + '\\' + dir, target_path + '\\' + dir)
                print(root + '\\' + dir + ' 复制成功-> ' + target_path)
                num += 1
        # for dir_in in dirs:
        # # for dir_in in root:
        #     copy_file(dir_in)


if __name__ == '__main__':
    # 文件夹路径
    source_path = r'D:\目标文件夹'
    # 输出路径
    target_path = r'C:\Users\Desktop\res'
    copy_file(source_path)
