import time
import xlrd

def print_star():
    time.sleep (0.1)
    print ('*', end = '', flush = True)

def tkp_number():
    number = input (f'\nвведи номер КП либо "enter" для выхода\n')

    
    if number == "":
        print ("работа программы завершена. любая кнопка для выхода")
        global i
        i = False
        input ()
        return i 
    elif len (number) != 7:
        print ('введено не 7 знаков. повтори ввод')
        tkp_number()
    else:
        return number    

def open_and_check_tkp(number):
    if i == True:
        tkp_manager = number[4:6]
        print_star()
        tkp_year = number[-1]
        print_star()
        
        file = xlrd.open_workbook (f'//Server-bd/tkp/202{tkp_year}/{tkp_manager}/{number}/score/{number}.xls')
        print_star()
        sheet = file.sheet_by_index (0)
        print_star()
        vals_all = [sheet.row_values(rownum) for rownum in range(sheet.nrows)] #получаю все значения построчно
        print_star()
        val_number_in_tkp = vals_all[8][1][3:10]  
        print_star()
               
        if val_number_in_tkp != number:
            print (f'номера КП не совпадают')
            print (f'номер файла {number}, номер кп в файле {val_number_in_tkp}')
        print_star()
        val_date_in_tkp_b9 = vals_all[8][1][14:24]
        print_star()
        
        for v in vals_all:
            mistake = 0
            try:
                print_star()
                if 'Дата' in v[1]:
                    val_date_in_tkp_down = v[1][:-2].split()
                    val_date_in_tkp_down = val_date_in_tkp_down[-1]
                    break
            except TypeError:
                pass
        
        if val_date_in_tkp_b9 != val_date_in_tkp_down:
            print (f'даты не совпадают')
            print (f'в шапке {val_date_in_tkp_b9}, в подвале {val_date_in_tkp_down}')
            mistake += 1

        count_srok = 0
        for val in vals_all[16:]:
            try:
                if val[1] != '' and val[2] != '' and val[3] != '':
                    lenghth_artikul = len(val[9])
                    if lenghth_artikul != 10 and lenghth_artikul != 16:
                        print(f'\nневерная длина артикула, {lenghth_artikul} символов')
                        mistake += 1
                    count_srok_current = int(val[8])
                    if count_srok_current > count_srok:
                        count_srok = count_srok_current
                if 'Срок' in val[1]:
                    srok_down = val[1].split()
                    srok_down = srok_down[2]
                    if int(srok_down) != count_srok:
                        print (f'\nСроки не совпадают. Максимальный справа {count_srok}, в подвале {srok_down}')
                        break
            except TypeError:
                pass
        
        if mistake == 0:
            print ('__ok__')
        else:
            print (f'\nошибки, {mistake} шт')

i = True
while i == True:
    number = tkp_number()
    open_and_check_tkp(number)
   
