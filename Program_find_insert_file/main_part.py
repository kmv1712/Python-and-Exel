import os
import xlrd, xlwt
import re
import xlwings as xw

# print(os.getcwd())
# print(os.listdir(path="./transform_Exel"))
name_f_account_info_two_part = os.listdir(path="./transform_Exel")
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

"""
Модуль для поиска информации в exel файле счета
Обратить внимание, что иногда распознается с ошибками
Таблица состоит из 10 колонок
"""


# Поиск номера счета
def get_number_account(item_one_info_in_account_f_Exel):
    number_account = re.findall('(\d+)', item_one_info_in_account_f_Exel)
    # print(number_account[0])
    return number_account[0]


# Поиск даты счета
def get_date_account(item_one_info_in_account_f_Exel):
    date_account = re.findall('(\d\d\s\w\w+\s\d\d\d\d)', item_one_info_in_account_f_Exel)
    return date_account[0]



# Функция возращает список из (номер счета, дата, прибор, сумма с НДС)
def seach_need_info_in_account_f(info_in_account_f_Exel):
    i = 0
    n = 0
    list_name_account_date_nds = []
    all_list_name_account_date_nds = []
    # print(info_in_account_f_Exel)
    for item_info_in_account_f_Exel in info_in_account_f_Exel:
        # print(item_info_in_account_f_Exel)
        for item_one_info_in_account_f_Exel in item_info_in_account_f_Exel:
            if re.search(r'СЧЕТ', item_one_info_in_account_f_Exel):
                i += 1
                number_account = get_number_account(item_one_info_in_account_f_Exel)
                # print('Номер: ' + number_account)
                list_name_account_date_nds.append(number_account)
                date_account = get_date_account(item_one_info_in_account_f_Exel)
                list_name_account_date_nds.append(date_account)
                # print('Дата:' + date_account)
            if re.search(r'ШТ', item_one_info_in_account_f_Exel):
                n += 1
                list_name_account_date_nds.append(item_info_in_account_f_Exel[0])
                list_name_account_date_nds.append(item_info_in_account_f_Exel[7])
                # print(item_info_in_account_f_Exel)
                # print('Прибор: ' + item_info_in_account_f_Exel[0])
                # print('Сумма с НДС: ' + item_info_in_account_f_Exel[7])
    all_list_name_account_date_nds.append(list_name_account_date_nds)
    print('Количество счетов: ' + str(i))
    print('Количество позиций поверки: ' + str(n))
    if n == i:
        print('Отлично в распознаном файле Exel ошибок нет')
    else:
        print('Ошибка просим обратить внимание на распознанты файл в название не правильно распозналось слово СЧЕТ '
              'или в таблицы не верно распознались ШТ')

    return all_list_name_account_date_nds

# Функция открытия основной таблицы Уралтест.xlsx
def open_main_task():
    name_main_task = os.getcwd() + '\\Уралтест.xlsx'
    open_main_task = xlrd.open_workbook(name_main_task)
    sheet_main_task = open_main_task.sheet_by_index(0)
    # получаем значение первой ячейки A1
    # val = sheet_f_account.row_values(0)[0]
    # print(val)
    # получаем список значений из всех записей
    info_in_main_task_Exel = [sheet_main_task.row_values(rownum) for rownum in range(sheet_main_task.nrows)]
    print(info_in_main_task_Exel)










# Функция добавления информации в таблицу
# def add_info_in_main_f():


# all_list_name_account_date_nds = seach_need_info_in_account_f(info_in_account_f_Exel)
# print(all_list_name_account_date_nds)
open_main_task()

# wb = xw.Book('Уралтест.xlsx')
# xw.Range('A1').value = 'Foo'

# rb = xlrd.open_workbook('Уралтест.xlsx')

# выбираем активный лист
# sheet = rb.sheet_by_index(0)
