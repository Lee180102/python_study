import os, re

# 替换的目录名以及文件名

def replace(path, str, repl):
    for root, dirs, files in os.walk(path):
        for f in files:
            old_file = os.path.join(root, f)
            new_file = os.path.join(root, re.sub(str, repl, f))
            os.rename(old_file, new_file)
        for d in dirs:
            old_dir = os.path.join(root, d)
            new_dir = os.path.join(root, re.sub(str, repl, d))
            os.rename(old_dir, new_dir)
            replace(new_dir, str, repl)
# 源字段
str = r'YF-QAXY2017-011'
# 替换的
repl = 'YF-QAXY2021-011'
# 目录
path = r'D:\Temps\word'
replace(path, str, repl)

print("end")
