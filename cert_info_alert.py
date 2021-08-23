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
# команда для командной строки windows = for /r cer_s %i in (*.cer) do certutil "%i" > "txt_s\%~ni.txt"
cer_command = rf'for /r {dir_cers} %i in (*.cer) do certutil "%i" > "{dir_txts}\%~ni.txt"'
name_file_xlsx = 'cert.xlsx'
list_of_strings_from_files = []


# функция очищающая папку dir_txts для создания новых дампов сертификатов
def clean_dir_txts():
    os.chdir(os.path.join(os.getcwd(), dir_txts))
    for data_of_scan in os.scandir():
        if data_of_scan.is_file() and os.path.splitext(os.path.split(data_of_scan)[1])[1] == etx_txt:
            os.remove(data_of_scan)
    print('\n(1)...папка от старых дампов сертификатов очищена')


# функция создающая дампы файлов и складывающая их в dir_txts путь
def do_txt_from_cer():
    # переход в корневую папку
    os.chdir(os.path.dirname(os.path.realpath(__file__)))
    subprocess.run(cer_command, stdout=subprocess.DEVNULL, shell=True)
    print('\n(2)...дампы сертификатов вновь созданы\n')


# функция чтения дампов сертификатов и формирования конечной таблицы для вывода её в xlsx
def processing_txt_files():
    # Поля для поиска в файле
    tuple_search_string = (
                            'SN',  # фамилия 'SN='
                            'G',   # имя отчество 'G='
                            'NotAfter',   # дата конца 'NotAfter:'
                            'NotBefore',  # дата начала 'NotBefore:'
                            'CN',  # организация выдавшая сертификат 'CN='
                            'Хеш сертификата(sha1)',  # отпечаток 'Хеш сертификата(sha1):'
                            'datetime.datetime.date(datetime.datetime.now())',  # текущая дата
                            'Серийный номер',  # отличие между органами выдавшими сертификат
                            'полный путь до сертификата'  # полный путь до сертификата  = os.path.abspath(data_of_scan)
                            )
    # переход в папку с дампами
    os.chdir(os.path.join(os.getcwd(), dir_txts))

    # чтение и обработка каждого файла в список, и добавление в конец списка "путь к файлу"
    for data_of_scan in os.scandir():
        if data_of_scan.is_file() and os.path.splitext(os.path.split(data_of_scan)[1])[1] == etx_txt:

            # получил все строки из файла
            all_strings_from_file = []
            with open(data_of_scan.name, 'r') as txt_file:
                all_strings_from_file = txt_file.read().splitlines()

            # выбрал из всех строк те, которые имеют вхождение из tuple_search_string
            list_of_need_strings = []
            for string_from_file in all_strings_from_file:
                string_from_file = string_from_file.strip()
                if (string_from_file.split(':', maxsplit=1)[0] in tuple_search_string) or (string_from_file.split('=', maxsplit=1)[0] in tuple_search_string):
                    list_of_need_strings.append(string_from_file)
            list_of_need_strings.append(os.path.abspath(data_of_scan))

            # тут из списка list_of_need_strings нужно вычислить лишние
            # И создать порядок для формирования конечного списка для выгрузки в эксель
            for string_from_need_list in list_of_need_strings:
                pass

            # создал список списков
            list_of_strings_from_files.append(list_of_need_strings)

    # print(*list_of_strings_from_files, sep='\n')
    print(list_of_strings_from_files)

    print('\n(3)...дампы сертификатов прочитаны и таблица для записи в xlsx готова')


# перенос списка данных в файл xlsx
def do_xlsx():
    # переход в корневую папку
    os.chdir(os.path.dirname(os.path.realpath(__file__)))

    # print(datetime.datetime.date(datetime.datetime.now()))

    file_xlsx = openpyxl.Workbook()
    file_xlsx_s = file_xlsx.active

    # собираю строку по правилу выгрузки и добавляю её в файл !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    for xls_str in list_of_strings_from_files:
        file_xlsx_s.append(xls_str)
        # pass

    # file_xlsx.save(name_file_xlsx.replace('cert','cert_'+str(datetime.datetime.date(datetime.datetime.now()))))
    file_xlsx.save(name_file_xlsx)
    file_xlsx.close()
    print('\n(4)...файл с данными сертификатов собран\n')
    print(f'сделано')


if __name__ == '__main__':
    clean_dir_txts()
    do_txt_from_cer()
    processing_txt_files()
    do_xlsx()

# **********************************************************************************************

    # # формирование1 list_of_strings_from_files
    # for dump in list_of_strings_from_files:
    #     # print()
    #     # print(dump[-1])
    #     list_of_strings_from_files = []
    #     for dump_string in dump:
    #         dump_string = dump_string.strip()
    #
    #         if dump_string.split(':', maxsplit=1)[0] in tuple_search_string:
    #             # print(f'....{dump_string.split(":", maxsplit=1)[0]}....{dump_string.split(":", maxsplit=1)[1].strip()}')
    #             list_of_strings_from_files.append(dump_string.split(":", maxsplit=1)[1].strip())
    #
    #         if dump_string.split('=', maxsplit=1)[0] in tuple_search_string:
    #             # print(f'....{dump_string.split("=", maxsplit=1)[0]}....{dump_string.split("=", maxsplit=1)[1].strip()}')
    #             list_of_strings_from_files.append(dump_string.split("=", maxsplit=1)[1].strip())
    #
    #     list_of_strings_from_files.append(dump[-1])
    #
    #     list_of_strings_from_files.append(list_of_strings_from_files)
    #     list_of_strings_from_files = []

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
