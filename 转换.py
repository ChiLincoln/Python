# --** coding="UTF-8" **--
# 
# author:SueMagic  time:2019-01-01
import os
import re
import sys

path = r"C:\Users\11832\Desktop\captcha_data"
if __name__ == "__main__":
    old_names = os.listdir(path)  # 取路径下的文件名，生成列表
    for old_name in old_names:  # 遍历列表下的文件名
        if old_name != sys.argv[0]:  # 代码本身文件路径，防止脚本文件放在path路径下时，被一起重命名
            if old_name.endswith('.png'):  # 当文件名以.txt后缀结尾时
                new_name = old_name.replace(".png", "_new.png")
                os.rename(os.path.join(path, old_name), os.path.join(path, new_name))  # 重命名文件
                print(old_name, "has been renamed successfully! New name is: ", new_name)  # 输出提示

