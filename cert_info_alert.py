# ...
# INSTALL
# pip install openpyxl
# COMPILE
# pyinstaller -F -w cert_info_alert.py.py
# ...

import os
import openpyxl
import subprocess

# переменные
etx_txt = '.txt'  # расширение файлов с текстом
dir_cers = r'cer_s'  # папка для сохранения сертификатов
dir_txts = r'txt_s'  # папка для дампов сертификатов
# команда для командной строки windows = for /r cer_s %i in (*.cer) do certutil "%i" > "txt_s\%~ni.txt"
cer_command = rf'for /r {dir_cers} %i in (*.cer) do certutil "%i" > "{dir_txts}\%~ni.txt"'


# функция очищающая папку dir_txts для создания новых дампов сертификатов
def clean_dir_txts():
    os.chdir(os.path.join(os.getcwd(), dir_txts))
    # print(*list(os.scandir()), sep='\n')
    # print()
    for data_of_scan in os.scandir():
        if data_of_scan.is_file() and os.path.splitext(os.path.split(data_of_scan)[1])[1] == etx_txt:
            os.remove(data_of_scan)
        # print(f'{data_of_scan.is_file() = }')
        # print(f'{data_of_scan.name = }')
        # # print(f'{os.path.split(data_of_scan) = }')
        # # print(f'{os.path.splitext(data_of_scan) = }')
        # # print(f'{os.path.splitext(os.path.split(data_of_scan)[1])[0] = }')
        # print(f'{os.path.splitext(os.path.split(data_of_scan)[1])[1] = }')
        # # print(f'{data_of_scan.path = }')
        # # print(f'{dir(data_of_scan) = }')
        # print()
    print('(1)...папка от старых дампов сертификатов очищена')
    print()


# функция создающая дампы файлов и складывающая их в dir_txts путь
def do_txt_from_cer():
    os.chdir(os.path.dirname(os.path.realpath(__file__)))
    os.system(cer_command)
    print('(2)...дампы сертификатов вновь созданы')
    print()


def read_txt_files():
    os.chdir(os.path.join(os.getcwd(), dir_txts))
    for data_of_scan in os.scandir():
        if data_of_scan.is_file() and os.path.splitext(os.path.split(data_of_scan)[1])[1] == etx_txt:
            pass

    print('(3)...дампы сертификатов прочитаны и таблица для записи в xlx готова')
    print()


def do_xlsx():
    pass


if __name__ == '__main__':
    clean_dir_txts()
    do_txt_from_cer()
    read_txt_files()
    do_xlsx()

    print(f'сделано')
