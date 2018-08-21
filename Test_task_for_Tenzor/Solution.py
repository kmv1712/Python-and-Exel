import os
import xlrd

# Выбираем лист с полезной информацией (ЛИСТ1)
useful_sheet = 0
# Название файла с EСXEL документами
name_dir = 'staff_efficiency'
# Получаем список названий файлов
name_file_with_inf = os.listdir(path="./" + name_dir)
# print(name_file_with_inf )
all_list_leader_in_doc = []

'''Поиск пустой строки в таблице 
useful_sheet - страница на которой ищем пустую строку
name_dir - имя директории с EXСEL файлами 
name_file - имя EXСEL документа в котором ищем пустую строку
'''


def get_empty_line_in_table(useful_sheet, name_dir, name_file):
    try:
        rb = xlrd.open_workbook(name_dir + '/' + name_file)
        # выбираем активный лист
        sheet = rb.sheet_by_index(useful_sheet)
        for i in range(0, 1000000000000):
            sheet.row_values(i)[0]
    except:
        return i


'''Открываем и получаем всю информацию с страницы useful_sheet, EXCEL документа в виде списка
useful_sheet - страница
name_dir - имя директории с EXСEL файлами 
name_file - имя EXСEL документа 
'''


def get_info_in_f_Ecxel(useful_sheet, name_dir, name_file):
    open_with_inf = xlrd.open_workbook(name_dir + '/' + name_file)
    sheet_with_inf = open_with_inf.sheet_by_index(useful_sheet)
    info_in_f_Ecxel = [sheet_with_inf.row_values(row_num) for row_num in range(sheet_with_inf.nrows)]
    return info_in_f_Ecxel


'''Получаем список всех руководителей с указанием выполнения проекта 
[Иванов, 1, 0] - проект выполнен в срок
[Иванов, 0, 1] - проект не выполнен в срок
Пример итог. вывода: [['Иванов Р.А.', 1, 0], ['Сидоров М.В.', 0, 1] ....]
empty_line_in_table см.функцию get_empty_line_in_table
info_in_f_Ecxel см.функцию get_info_in_f_Ecxel
'''


def get_all_list_leader_in_doc(empty_line_in_table, info_in_f_Ecxel):
    for i in range(2, empty_line_in_table):
        list_leader = []
        # print(info_in_f_Ecxel[i])
        date_finish_fact = info_in_f_Ecxel[i][3]
        if date_finish_fact != '':
            name_leader = info_in_f_Ecxel[i][1]
            date_finish_plan = info_in_f_Ecxel[i][2]
            date_finish_fact = info_in_f_Ecxel[i][3]
            if date_finish_plan >= date_finish_fact:
                good_finish = 1
                bad_finish = 0
                # print(name_leader)
                # print(date_finish_plan)
                # print(date_finish_fact)
            else:
                good_finish = 0
                bad_finish = 1
            list_leader.append(name_leader)
            list_leader.append(good_finish)
            list_leader.append(bad_finish)
            all_list_leader_in_doc.append(list_leader)
        else:
            continue
    return all_list_leader_in_doc


for name_file in name_file_with_inf:
    info_in_f_Ecxel = get_info_in_f_Ecxel(useful_sheet, name_dir, name_file)
    empty_line_in_table = get_empty_line_in_table(useful_sheet, name_dir, name_file)
    all_list_leader_in_doc = get_all_list_leader_in_doc(empty_line_in_table, info_in_f_Ecxel)
    print(all_list_leader_in_doc)

list_name_leader_one = []
for num_list in range(0, len(all_list_leader_in_doc) - 1):
    list_name_leader_one.append(all_list_leader_in_doc[num_list][0])
print(list_name_leader_one)
# НЕОБХОДИМО ПОЛУЧИТЬ СПИСОК ФАМИЛИЙ ИЗ ОБЩЕГО СПИСКА т.е. убрать дубли
