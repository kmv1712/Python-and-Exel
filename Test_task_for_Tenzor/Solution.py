# Начало файла
import os
useful_sheet= 0
name_file_with_inf = os.listdir(path="./transform_Exel")

# Формируем путь нахождения файла с счетами
name_f_account_info = os.getcwd() + '\\transform_Exel\\' + name_f_account_info_two_part[0]
# print(name_f_account_info)

"""
Не работает если распознаный в exel файл открыт 
Необходимо проработать 
"""

open_f_account = xlrd.open_workbook(name_f_account_info)
sheet_f_account = open_f_account.sheet_by_index(0)
# получаем значение первой ячейки A1
# val = sheet_f_account.row_values(0)[0]
# print(val)
# получаем список значений из всех записей
info_in_account_f_Exel = [sheet_f_account.row_values(rownum) for rownum in range(sheet_f_account.nrows)]
# print(info_in_account_f_Exel)

