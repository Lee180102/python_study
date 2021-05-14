# @author   Lee 
# @time     2020/11/1 2:00 下午

import os


def file_name(file_dir):
    file_list = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            file_list.append(root + '/' + file)
    return file_list

