# ...
# INSTALL
# pip install openpyxl
# COMPILE
# pyinstaller -F -w cert_info_alert.py.py
# ...

# переменные
dir_cers = ''
dir_txts = ''
cer_command = 'for /r cer_s %%i in (*.cer) do certutil "%%i" > "txt_s\%%~ni.txt"'

def do_txt_from_cer():
    pass


if __name__ == '__main__':
    do_txt_from_cer()


    print()
    print(f'сделано')
