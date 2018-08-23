import os
import xlrd
from itertools import groupby

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


'''Получаем список всех руководителей с указанием успешности выполнения проектов 
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


'''Для получения списка персонала all_list_leader_in_doc см.функцию get_all_list_leader_in_doc - Получаем список всех 
руководителей с указанием успешности выполнения проектов '''


def get_list_leader(all_list_leader_in_doc):
    list_name_leader_one = []
    for num_list in range(0, len(all_list_leader_in_doc) - 1):
        list_name_leader_one.append(all_list_leader_in_doc[num_list][0])
    # print(list_name_leader_one)
    list_leader = list(set(list_name_leader_one))
    return list_leader


'''Для получения списка вида [[ФИО, кол-во проектов, кол-во вовремя завершенных проектов, кол-во не вовремя 
завершенные проектов], ...] all_list_leader_in_doc см.функцию get_all_list_leader_in_doc - Получаем список всех 
руководителей с указанием успешности выполнения проектов Улучшит: оптимизировать код (получать значения путем 
группировки списка all_list_leader_in_doc) '''


def get_list_with_eff_leader(all_list_leader_in_doc):
    list_leader = get_list_leader(all_list_leader_in_doc)
    list_with_eff_leader = [[item_list_leader, 0, 0, 0] for item_list_leader in list_leader]
    #     print(list_with_eff_leader)
    for item_list_leader in list_leader:
        for item_all_list_leader_in_doc in all_list_leader_in_doc:
            if item_list_leader == item_all_list_leader_in_doc[0]:
                for i_eff_leader in range(0, len(list_with_eff_leader)):
                    if list_with_eff_leader[i_eff_leader][0] == item_list_leader:
                        list_with_eff_leader[i_eff_leader][1] += 1
                        if item_all_list_leader_in_doc[1] == 1:
                            list_with_eff_leader[i_eff_leader][2] += 1
                        elif item_all_list_leader_in_doc[2] == 1:
                            list_with_eff_leader[i_eff_leader][3] += 1
    return list_with_eff_leader


''' 
Добавляем колонку с процентами характеризующих выполения проекта в срок.
list_with_eff_leader  [['Ф.И.О', 'кол-во проектов' , 'вып. в срок ','вып. не в срок'], ....]
'''


def add_percent_eff(list_with_eff_leader):
    list_with_eff_leader_and_percent = []
    # print(list_with_eff_leader)
    for item_list_with_eff_leader in list_with_eff_leader:
        #     print(item_list_with_eff_leader)
        percent_eff = round(
            ((item_list_with_eff_leader[1] - item_list_with_eff_leader[3]) / item_list_with_eff_leader[1]) * 100, 2)
        item_list_with_eff_leader.append(percent_eff)
        list_with_eff_leader_and_percent.append(item_list_with_eff_leader)
    return list_with_eff_leader_and_percent


'''Выводим список ответсвенных за проект отсортированых по (1) по имени (2) по кол-ву проектов (3) по кол-ву проектов 
вып. в срок. (4) по кол-ву проектов не вып. в срок. (5) по проценту выполения проекта в срок. 
list_with_eff_leader_and_percent  [['Ф.И.О', 'кол-во проектов' , 'вып. в срок ','вып. не в срок','проценту выполения 
проекта в срок'], ....] '''


def print_list_leader(list_with_eff_leader_and_percent):
    # print(list_with_eff_leader)
    n = input(
        'Сортировать специалиста в качестве руководителя по имени (1), по кол-ву проектов (2), по кол-ву проектов вып. в срок (3),  по кол-ву проектов вып. не в срок(4),по проценту выполения проекта в срок(5): ')
    n = int(n) - 1
    # t = input('По возрастанию (0), по убыванию (1): ')
    # t = int(t)
    t = 1
    list_with_eff_leader_and_percent.sort(key = lambda i: i[n], reverse = t)
    print('%20s | %15s | %15s | %15s | %15s  ' % (
    'Ф.И.О', 'кол-во проектов', 'вып. в срок ', 'вып. не в срок', 'проценту выполения проекта в срок'))
    for i in list_with_eff_leader_and_percent:
        print('%20s | %15d | %15d | %15d | %15d  ' % (str(i[0]), i[1], i[2], i[3], i[4]))

        
'''Генерируем БД с 
ФИО,
кол-во проектов
кол-во вып в срок
кол-во не вып в срок'''

def get_list_employee_name(name_file_with_inf):
    list_employee_name = []
    for item_inf in range(0, len(name_file_with_inf)):
        employee_name = (list_all_info_in_f_Ecxel[item_inf][0][4: len(list_all_info_in_f_Ecxel[item_inf][0])])
        for item_employee_name in employee_name:
            list_employee_name.append(item_employee_name)
        list_employee_name = [el for el, _ in groupby(list_employee_name)]
        list_employee_inf = [[item_list_employee_name, 0, 0, 0] for item_list_employee_name in list_employee_name]
        return list_employee_inf

'''Получаем список сотрудников ответсвенных за проект'''
def get_print_list_leader():    
    # Выбираем лист с полезной информацией (ЛИСТ1)
    useful_sheet = 0
    # Название файла с EСXEL документами
    name_dir = 'staff_efficiency'
    # Получаем список названий файлов
    name_file_with_inf = os.listdir(path="./" + name_dir)
    # print(name_file_with_inf )
    all_list_leader_in_doc = []
    for name_file in name_file_with_inf:
        info_in_f_Ecxel = get_info_in_f_Ecxel(useful_sheet, name_dir, name_file)
        empty_line_in_table = get_empty_line_in_table(useful_sheet, name_dir, name_file)
        all_list_leader_in_doc = get_all_list_leader_in_doc(empty_line_in_table, info_in_f_Ecxel)
    all_list_leader_in_doc = get_all_list_leader_in_doc(empty_line_in_table, info_in_f_Ecxel)
    list_with_eff_leader = get_list_with_eff_leader(all_list_leader_in_doc)
    list_with_eff_leader_and_percent = add_percent_eff(list_with_eff_leader)      
    print_list_leader (list_with_eff_leader_and_percent) 
    
    
'''
Создаем список на основе всез EXCEL файлов с параметрами 
ФИО, 
кол-во проектов,
кол-во проектов выполненых в срок,
кол-во проектов выполненых не в срок
'''
def get_list_all_info_in_f_Ecxel(name_file_with_inf):
    list_all_info_in_f_Ecxel = []
    for name_file in name_file_with_inf:
        info_in_f_Ecxel = get_info_in_f_Ecxel(useful_sheet, name_dir, name_file)
        list_all_info_in_f_Ecxel.append(info_in_f_Ecxel)
        return list_all_info_in_f_Ecxel
 
'''Выводим список сотрудников, как исполнителей '''
def get_print_list_employee():
    # Выбираем лист с полезной информацией (ЛИСТ1)
    useful_sheet = 0
    # Название файла с EСXEL документами
    name_dir = 'staff_efficiency'
    # Получаем список названий файлов
    name_file_with_inf = os.listdir(path="./" + name_dir)
    list_all_info_in_f_Ecxel = []
    all_list_employee_in_doc = []
    for name_file in name_file_with_inf:
        info_in_f_Ecxel = get_info_in_f_Ecxel(useful_sheet, name_dir, name_file)
        list_all_info_in_f_Ecxel.append(info_in_f_Ecxel)
    for item_inf in range(0, len(name_file_with_inf)):
        # print(list_all_info_in_f_Ecxel[item_inf][0])
        # print(len(list_all_info_in_f_Ecxel[item_inf][0]) - 1)
        # print(len(list_all_info_in_f_Ecxel[item_inf]))
        employee_name = (list_all_info_in_f_Ecxel[item_inf][0][4 : len(list_all_info_in_f_Ecxel[item_inf][0])])
        list_pl_fa_in_table = []
        for item_in_f_Ecxel in range (2 , len(list_all_info_in_f_Ecxel[item_inf])):
            list_with_plan_fact = list_all_info_in_f_Ecxel[item_inf][item_in_f_Ecxel][4: len(list_all_info_in_f_Ecxel[item_inf][0])]
            list_pl_fa_in_table.append(list_with_plan_fact)
        #print(list_pl_fa_in_table)
        employee_name = [el for el, _ in groupby(employee_name)]
        #print(employee_name)
        for item_employee_name in employee_name:
            position_list = employee_name.index(item_employee_name)
            #print (len(list_pl_fa_in_table))
            for item_list_pl_fa_in_table in range (0, len(list_pl_fa_in_table)):
                list_employee = []            
                if list_pl_fa_in_table[item_list_pl_fa_in_table][1] == '':
                    list_pl_fa_in_table[item_list_pl_fa_in_table].pop(1)
                    list_pl_fa_in_table[item_list_pl_fa_in_table].pop(0)
                    continue
                elif list_pl_fa_in_table[item_list_pl_fa_in_table][0] == '':
                    good_finish = 1
                    bad_finish = 0                 
                    list_pl_fa_in_table[item_list_pl_fa_in_table].pop(1)
                    list_pl_fa_in_table[item_list_pl_fa_in_table].pop(0)
                elif int(list_pl_fa_in_table[item_list_pl_fa_in_table][0]) >= int(list_pl_fa_in_table[item_list_pl_fa_in_table][1]):                
                    good_finish = 1
                    bad_finish = 0
                    list_pl_fa_in_table[item_list_pl_fa_in_table].pop(1)
                    list_pl_fa_in_table[item_list_pl_fa_in_table].pop(0)
                elif int(list_pl_fa_in_table[item_list_pl_fa_in_table][0]) < int(list_pl_fa_in_table[item_list_pl_fa_in_table][1]):                
                    good_finish = 0
                    bad_finish = 1
                    list_pl_fa_in_table[item_list_pl_fa_in_table].pop(1)
                    list_pl_fa_in_table[item_list_pl_fa_in_table].pop(0)
                list_employee.append(item_employee_name)
                list_employee.append(good_finish)
                list_employee.append(bad_finish)
                if len(list_employee) != 0:
                    #print(list_employee)
                    all_list_employee_in_doc.append(list_employee)
                else:
                    continue
    all_list_employee_in_doc = get_list_with_eff_leader(all_list_employee_in_doc)
    all_list_employee_in_doc = add_percent_eff(all_list_employee_in_doc)  
    print_list_leader (all_list_employee_in_doc) 

        


user_enter = int(input('Сортировать специалистов в качестве руководителя (1) \nСортировать специалистов в качестве исполнителя (2)'))
if user_enter == 1:
    get_print_list_leader() 
elif (user_enter == 2):
    get_print_list_employee() 
else:
    print('ok')
