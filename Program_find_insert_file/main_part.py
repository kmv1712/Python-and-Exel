import os
import xlrd
import re
import xlwings as xw

# Выбранна чтраница Уралтест в основной таблице
name_sheet = 0
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
# Продумать как преобразовать 10 января 2018 в 10.01.2018 ---- СДЕЛАЛ 15082018
''' 15.08.
1) Оптимизировать (сделать код более компактным за счет list compression)
2) Сделать отдельную фунцию на обработку даты (добавить различные варианты для обработки месяца)
3) Продумать, как собирать данные с ошибочным распознанием и применить их при дальнейшем использование программы 
пример: распознал июл1 надо чтобы мы сохранили этот вариант и при следующем таком распознание мы получили 07, а не вариант с возможностью занести инф. самим
'''
# 11082018---------------1603
def get_date_account(item_one_info_in_account_f_Exel):
    date_account = re.findall('(\d\d\s\w\w+\s\d\d\d\d)', item_one_info_in_account_f_Exel)
    for date_account_one in date_account:
        date_account = date_account_one
    date_account = date_account.split(' ') 
    if date_account [1] == 'января':
        date_account [1] = '01'
    elif date_account [1] == 'февраля':
        date_account [1] = '02'
    elif date_account [1] == 'марта':
        date_account [1] = '03'
    elif date_account [1] == 'апреля':
        date_account [1] = '04'
    elif date_account [1] == 'мая':
        date_account [1] = '05'
    elif date_account [1] == 'июня':
        date_account [1] = '06'
    elif date_account [1] == 'июля':
        date_account [1] = '07'
    elif date_account [1] == 'августа':
        date_account [1] = '08'    
    elif date_account [1] == 'сентября':
        date_account [1] = '09'
    elif date_account [1] == 'октября':
        date_account [1] = '10'
    elif date_account [1] == 'ноября':
        date_account [1] = '11'
    elif date_account [1] == 'декабря':
        date_account [1] = '12'
    else:
        date_account [1] = input ('На экране указан нераспознаный месяц, если вы не можете его индефецировать то введите две цифры, соответсвующие этому месяцу, пример: 02')
    date_account = (str(date_account [0]) + '.' + date_account [1] + '.' + str(date_account [2])) 
    return date_account


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
    sheet_main_task = open_main_task.sheet_by_index(name_sheet)
    # получаем значение первой ячейки A1
    # val = sheet_f_account.row_values(0)[0]
    # print(val)
    # получаем список значений из всех записей
    info_in_main_task_Exel = [sheet_main_task.row_values(rownum) for rownum in range(sheet_main_task.nrows)]
    print(info_in_main_task_Exel)


# Вылавливаем ошибку для нахождения последней строки и вычетаем 1 для продолжения работы программы
# Находим пустую строку в тексте
def get_empty_line_in_table():
    try:
        rb = xlrd.open_workbook('Уралтест.xlsx')
        # выбираем активный лист
        sheet = rb.sheet_by_index(name_sheet)
        for i in range(0, 1000000000000):
            sheet.row_values(i)[0]
    except:
        i = i + 1
        return i


"""Функция генерирует список all_list_name_account_date_nds
по след маске [[номер_счета, дата, прибор, сумма с НДС], [-и-], ...]
"""
def get_sort_all_list_name_account_date_nds(all_list_name_account_date_nds):
    n = 0
    j = 4
    sort_all_list_name_account_date_nds = []
    count_insert_list = len(all_list_name_account_date_nds[0]) / 4
    # print(count_insert_list)
    # print(len(all_list_name_account_date_nds[0]))
    for i in range(0, int(count_insert_list)):
        sort_all_list_name_account_date_nds.append(all_list_name_account_date_nds[0][n:j])
        n = n + 4
        j = j + 4
    return (sort_all_list_name_account_date_nds)


"""Вставляем значение в файл Уралтест
Счет на оплату
(A)прибор
(E)номер
(F)дата
(G)Сумма с НДС
список вида [[номер_счета, дата, прибор, сумма с НДС], [-и-], ...]
"""
# Продумать закрытие и сохранение занесеных данных
# 1508 Продумать чтобы данные заносились в ячейки и фильтр их сортировал верно
def add_info_in_main_f(empty_line_in_table, sort_all_list_name_account_date_nds):
    xw.Book('Уралтест.xlsx')
    for i in range(0, len(sort_all_list_name_account_date_nds)):
        xw.Range('A' + str(empty_line_in_table)).value = sort_all_list_name_account_date_nds[i][2]
        xw.Range('E' + str(empty_line_in_table)).value = sort_all_list_name_account_date_nds[i][0]
        xw.Range('F' + str(empty_line_in_table)).value = sort_all_list_name_account_date_nds[i][1]
        xw.Range('G' + str(empty_line_in_table)).value = sort_all_list_name_account_date_nds[i][3]
        empty_line_in_table = empty_line_in_table + 1


# Функция выводит список в списке с дангыми номер, дата, прибор, сумма с НДС
all_list_name_account_date_nds = seach_need_info_in_account_f(info_in_account_f_Exel)
# print(all_list_name_account_date_nds)

# Функция возращает список с даными с листа
# open_main_task()

# Возращает номер последней строки таблицы
empty_line_in_table = get_empty_line_in_table()
# print(empty_line_in_table)

# Функция возращает список вида [[номер_счета, дата, прибор, сумма с НДС], [-и-], ...]
sort_all_list_name_account_date_nds = get_sort_all_list_name_account_date_nds(all_list_name_account_date_nds)
# print(sort_all_list_name_account_date_nds)

# Функция добавления информации в таблицу
add_info_in_main_f(empty_line_in_table, sort_all_list_name_account_date_nds)
