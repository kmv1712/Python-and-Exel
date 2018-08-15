import os
import shutil

# Формируем список имен файлов для перемещения
name_file_what_need_sort = os.listdir(path = "C:/Users/KarpovM/Desktop/Sort")
# print (name_file_what_need_sort)

def seach_place_one(need_val, need_val_two):    
    way_list_file = 'W:/СИиЗП/1 Участок испытаний и измерений/1.1 Подстанции/Диагностика/МАСЛО/'
    list_file = [way_list_file + 'Восточный', way_list_file + 'Северный', way_list_file + 'Юго-Западный']
    for item in list_file:
        list_res = os.listdir(path = item)
        for list_res_name in list_res:
            if need_val == list_res_name:
                return (item + '/' + need_val + '/' + need_val_two)
            else:
                continue

# Дробим элементы списка по пробелу получаем ['Академическая', 'Т1', '03.07.18.jpg'] для дальнейшго поиска пример
for name_file in name_file_what_need_sort:
    PS_name = name_file.split(" ")
    need_val = PS_name[0]
    need_val_two = PS_name[1]
    wer = seach_place_one(need_val, need_val_two)
    print (wer)
    if (wer is not None ):         
        name_file_peren = 'C:/Users/KarpovM/Desktop/Sort/' + name_file
        print (name_file_peren)
        file_what_need_sort = wer
        print (file_what_need_sort)
        shutil.move(name_file_peren, file_what_need_sort)
                            
    
    
# import shutil
# name_file = 'C:/Users/KarpovM/Desktop/Sort/Авиатор Т2 22.06.18.jpg'
# file_what_need_sort = 'W:/СИиЗП/1 Участок испытаний и измерений/1.1 Подстанции/Диагностика/МАСЛО/Восточный'
# shutil.move(name_file, file_what_need_sort)    
    
    
    
#     for trans in PS_name:
#         if (trans == "Т or trans == "Т"  or trans == "Т" or trans == "Т" ):
            
#     PS_name[0]
#     vost_res = os.listdir(path = "W:/СИиЗП/1 Участок испытаний и измерений/1.1 Подстанции/Диагностика/МАСЛО/Восточный")
#     for item in vost_res:
#         if (item == PS_name[0]):
#             name_file
            
            
        
    
    
    
    
#     open W:\СИиЗП\1 Участок испытаний и измерений\1.1 Подстанции\Диагностика\МАСЛО\Восточный
#     open W:\СИиЗП\1 Участок испытаний и измерений\1.1 Подстанции\Диагностика\МАСЛО\Северный
#     open W:\СИиЗП\1 Участок испытаний и измерений\1.1 Подстанции\Диагностика\МАСЛО\Юго-Западный
    
#     print (PS_name)
    
