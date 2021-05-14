# @author   Lee 
# @time     2020/11/1 3:21 下午
from openpyxl import Workbook


def dict_to_excel(value_dict):
    wb = Workbook()

    ws = wb.active
    ws.append([1, 2, 3])
    for item in value_dict:
        name = item.split('_')[0]
        department = item.split('_')[1]
        time = value_dict[item]
        print(name,department,time)
        ws.append([name, department, time])
    wb.save("sample.xlsx")
