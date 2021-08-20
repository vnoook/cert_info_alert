# ...
# INSTALL
# pip install openpyxl
# COMPILE
# pyinstaller -F -w cert_info_alert.py.py
# ...

import os
import openpyxl

# переменные
etx_txt = '.txt'
dir_cers = r'cer_s'
dir_txts = r'txt_s'
# for /r cer_s %%i in (*.cer) do certutil "%%i" > "txt_s\%%~ni.txt"
cer_command = 'for /r cer_s %%i in (*.cer) do certutil "%%i" > "txt_s\%%~ni.txt"'


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
    print('(1)...папка от txt файлов очищена')
    print()


def do_txt_from_cer():
    # os.chdir(os.getcwd())
    # os.system(cer_command)
    # print('(2)...файлы txt из сертификатов сделаны')
    # print()
    pass


def read_txt_files():
    pass


def do_xlsx():
    pass


if __name__ == '__main__':
    clean_dir_txts()
    do_txt_from_cer()
    read_txt_files()
    do_xlsx()

    print(f'сделано')
