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
            if "安科讯" in dir:
                shutil.copytree(root + '\\' + dir, target_path + '\\' + f"{dir}{num}")
            #     shutil.copytree(root + '\\' + dir, target_path + '\\' + dir)
                print(root + '\\' + dir + ' 复制成功-> ' + target_path)
                num += 1
        # for dir_in in dirs:
        # # for dir_in in root:
        #     copy_file(dir_in)


if __name__ == '__main__':
    # 文件夹路径
    source_path = r'D:\供应链\订单&需求单\前传小站'
    # 输出路径
    target_path = r'C:\Users\pdcfi\Desktop\res'
    copy_file(source_path)

#
#

"""
这段代码首先指定了源文件夹路径和目标文件夹路径，然后遍历源文件夹下的所有文件夹名称。
如果某个文件夹名称以"HQDD"开头且是一个文件夹，就将其复制到桌面上。最后输出复制成功的信息。
注意需要在字符串前面加上r，以表示原始字符串，可以避免转义字符等问题。
"""
# import os
# import shutil


# src_folder = r'D:\供应链\订单&需求单\前传小站'
# # 输出路径
# dst_folder = r'C:\Users\pdcfi\Desktop\res'
#
# for folder_name in os.listdir(src_folder):
#     print(folder_name)
#     if folder_name.startswith("安科讯"):
#         if folder_name.startswith("HQDD") and os.path.isdir(os.path.join(src_folder, folder_name)):
#             src_path = os.path.join(src_folder, folder_name)
#             dst_path = os.path.join(dst_folder, folder_name)
#             shutil.copytree(src_path, dst_path)
#             print(f"Copied folder {folder_name} to desktop.")


