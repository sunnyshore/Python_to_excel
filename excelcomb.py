import os
import xlrd
from openpyxl import Workbook

workbook_list = []
workbook_path = r'C:\Users\ABOCN\Downloads'
for dir_path, subdir_path, files in os.walk(workbook_path):
    for file in files:
        file_path = os.path.join(dir_path, file)
        workbook_list.append(file_path)


wb_head = xlrd.open_workbook(workbook_list[0])
table_head = wb_head.sheets()[0]
wb_comb = Workbook()
wb_comb_table = wb_comb.active
wb_comb_table.title = r'Sheet1'
for j in range(5):
    wb_comb_table.append(table_head.row_values(j))
i = len(workbook_list)-1
while i > 0:
    wb_src = xlrd.open_workbook(workbook_list[i])
    table_src = wb_src.sheets()[0]
    wb_comb_table.append(table_src.row_values(4))
    i -= 1
wb_comb.save(r'还款计划表汇总')
for x in workbook_list:
    os.remove(x)
