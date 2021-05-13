# @author   Lee 
# @time     2020/11/1 3:02 下午

import ExcelTodict
import File
import ToExcel
fileList = File.file_name('D:/')
values = ExcelTodict.fileDict(fileList)
ToExcel.dict_to_excel(values)