# ...
# INSTALL
# pip install openpyxl
# COMPILE
# pyinstaller -F -w cert_info_alert.py
# ...

import os
import datetime
import subprocess
import openpyxl

# переменные
etx_txt = '.txt'  # расширение файлов с текстом
dir_cers = r'cer_s'  # папка для сохранения сертификатов
dir_txts = r'txt_s'  # папка для дампов сертификатов
# команда для командной строки
cer_command = rf'for /r {dir_cers} %i in (*.cer) do certutil "%i" > "{dir_txts}\%~ni.txt"'
# файл шаблон для называния файла выгрузки
name_file_xlsx = 'cert.xlsx'
# конечный список со строками данных в нужном порядке который выгружается в эксель
list_of_strings_from_files = []


# функция очищающая папку {dir_txts} для создания новых дампов сертификатов
def clean_dir_txts():
    # переход в папку дампов
    os.chdir(os.path.join(os.getcwd(), dir_txts))
    # поиск файлов txt и их удаление
    for data_of_scan in os.scandir():
        if data_of_scan.is_file() and os.path.splitext(os.path.split(data_of_scan)[1])[1] == etx_txt:
            os.remove(data_of_scan)
    print('\n(1)...папка от старых дампов сертификатов очищена')


# функция создающая дампы файлов и складывающая их в папку {dir_txts}
def do_txt_from_cer():
    # переход в корневую папку
    os.chdir(os.path.dirname(os.path.realpath(__file__)))
    # запуск процесса создания дампов без вывода на экран результатов
    subprocess.run(cer_command, stdout=subprocess.DEVNULL, shell=True)
    print('\n(2)...дампы сертификатов вновь созданы')


# функция чтения дампов сертификатов и формирования конечной таблицы для вывода её в xlsx
def processing_txt_files():
    # поля и их порядок для поиска в дампах
    # если добавить или удалить поля (кроме последнего), то алгоритм не собъётся
    tuple_search_string = (
                           'SN',  # фамилия 'SN='
                           'G',   # имя отчество 'G='
                           'NotAfter',   # дата конца 'NotAfter:'
                           'NotBefore',  # дата начала 'NotBefore:'
                           'CN',  # организация выдавшая сертификат 'CN='
                           'Хеш сертификата(sha1)',  # отпечаток 'Хеш сертификата(sha1):'
                           'Серийный номер',  # отличие между органами выдавшими сертификат
                           'полный путь до дампа'  # полный путь до дампа
                          )
    # переход в папку с дампами
    os.chdir(os.path.join(os.getcwd(), dir_txts))

    # чтение и обработка каждого файла в список, и добавление в конец списка "путь к дампу"
    for data_of_scan in os.scandir():
        if data_of_scan.is_file() and os.path.splitext(os.path.split(data_of_scan)[1])[1] == etx_txt:
            # получил все строки из файла
            all_strings_from_file = []
            with open(data_of_scan.name, 'r') as txt_file:
                all_strings_from_file = txt_file.read().splitlines()

            # выбрал из всех строк те, которые имеются суфиксы из tuple_search_string
            # и заменил в строках где есть первый слева символ "=" на ":", для простоты в будущем
            list_of_need_strings = []
            for string_from_file in all_strings_from_file:
                string_from_file = string_from_file.strip()
                if (string_from_file.split(':', maxsplit=1)[0] in tuple_search_string) or\
                            (string_from_file.split('=', maxsplit=1)[0] in tuple_search_string):
                    list_of_need_strings.append(string_from_file.replace('=', ':', 1))

            # из списка list_of_need_strings нужно вычислить задвоенные суфиксы
            # И создать порядок для формирования конечного списка для выгрузки в xlsx
            list_of_need_strings_sorted = []
            for suffix in tuple_search_string:
                count_suffix = 0
                for string_from_need_list in list_of_need_strings:
                    suffix_of_need_string = string_from_need_list.split(':', maxsplit=1)[0]
                    if suffix_of_need_string == suffix:  # and (count_suffix == 0):
                        value_of_need_string = string_from_need_list.split(':', maxsplit=1)[1]
                        if count_suffix == 0:
                            if suffix_of_need_string == 'Хеш сертификата(sha1)':
                                list_of_need_strings_sorted.append(value_of_need_string.replace(' ',''))
                            else:
                                list_of_need_strings_sorted.append(value_of_need_string)
                        count_suffix += 1
                if count_suffix == 0:
                    list_of_need_strings_sorted.append('нет данных в дампе сертификата')
                    # print(f'{suffix = } ... {count_suffix = } ... {string_from_need_list}')
                    pass

            # print(string_from_need_list.split(':', maxsplit=1))
            # print(*list_of_need_strings_sorted, sep='\n')

            # и добавил в последний индекс ссылку на дамп сертификата
            list_of_need_strings_sorted[-1] = os.path.abspath(data_of_scan)

            # создал список списков
            list_of_strings_from_files.append(list_of_need_strings_sorted)

    # добавил в первую строку шапку из названий колонок
    list_of_strings_from_files.insert(0, list(val for val in tuple_search_string))

    # print(*list_of_strings_from_files, sep='\n')
    # print(list_of_strings_from_files)

    print('\n(3)...дампы сертификатов прочитаны и таблица для записи в xlsx готова')


# перенос списка данных в файл xlsx
def do_xlsx():
    # переход в корневую папку
    os.chdir(os.path.dirname(os.path.realpath(__file__)))

    # print(datetime.datetime.date(datetime.datetime.now()))

    file_xlsx = openpyxl.Workbook()
    file_xlsx_s = file_xlsx.active

    # собираю строку по правилу выгрузки и добавляю её в файл
    for xls_str in list_of_strings_from_files:
        file_xlsx_s.append(xls_str)
        # pass

    file_xlsx.save(name_file_xlsx.replace('cert', 'cert_'+str(datetime.datetime.date(datetime.datetime.now()))))
    # file_xlsx.save(name_file_xlsx)
    file_xlsx.close()

    print('\n(4)...файл с данными сертификатов собран')
    print('\n(5)...ГОТОВО!')

def run():
    clean_dir_txts()
    do_txt_from_cer()
    processing_txt_files()
    do_xlsx()


if __name__ == '__main__':
    run()

# **********************************************************************************************

# print(*list(os.scandir()), sep='\n')
# print(f'{data_of_scan.is_file() = }')
# print(f'{data_of_scan.name = }')
# # print(f'{os.path.split(data_of_scan) = }')
# # print(f'{os.path.splitext(data_of_scan) = }')
# # print(f'{os.path.splitext(os.path.split(data_of_scan)[1])[0] = }')
# print(f'{os.path.splitext(os.path.split(data_of_scan)[1])[1] = }')
# # print(f'{data_of_scan.path = }')
# # print(f'{dir(data_of_scan) = }')
# print()
