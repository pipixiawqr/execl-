import os
from openpyxl import *

def delete_col(filename):
    black_list = ['1', '行', '要', '3', '时间', '姓名']
    wb = load_workbook(filename)

    ws = wb.active
    col_nums = ws.max_column  # 总列数
    list=[]
    for x in range(1, col_nums + 1):  # 循环所有列
        title = ws.cell(row=1, column=x).value  # 提取列名的
        for keyword in black_list:  # 判断是不是在要删除的里面
            # print(str(title))
            if keyword == str(title).strip():
                list.append(x)
    list.sort(reverse=True)
    for i in list:
        ws.delete_cols(i)
    wb.save(filename)

def delete_row(filename):
    black_list = ['1', '行', '要', '3', '时间', '姓名']
    wb = load_workbook(filename)
    ws = wb.active
    row_nums = ws.max_row  # 总行数
    list=[]
    for x in range(2, row_nums + 1):
        title = ws.cell(column=6, row=x).value 
        keyword = str(title).strip()
        if keyword in black_list:
            pass
        else:
            list.append(x)
    list.sort(reverse=True)
    # map(ws.delete_rows, list)
    for i in list:
        ws.delete_rows(i)
    wb.save(filename) 

def readFiles(path):
    files = os.listdir(path)
    for file in files:
        if ".DS_Store" in file:
            continue
        execls = os.listdir(path+"/"+file)
        for execl in execls:
            if ".xlsx" in execl:
                delete_col(path+"/"+file+"/"+execl)
                delete_row(path+"/"+file+"/"+execl)
            else:
                continue
if __name__ == "__main__":
    readFiles('./aaa')
