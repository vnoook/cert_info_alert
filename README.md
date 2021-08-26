# cert_info_alert
Программа для проверки дат окончания сертификатов.
 
В папке с программой должны быть папки {dir_cers} и {dir_txts}.
 
В папку {dir_cers} скопируйте сертификаты с расширение ".cer".
 
Папка {dir_txts} нужна для хранения дампов сертификатов из {dir_cers}.
 
При наличии сертификатов в папке {dir_cers} с помощью программы Windows "certutil" создаются дампы с текстовом формате.
 
Эти дампы анализируются, и сортируются по порядку изложенному в переменной {tuple_search_string} и выгружаются в xlsx файл.
 
В полученном файле xlsx подсвечиваются ячейки дат со скором окончанием.
 
Красным подсвечиваются сроки "месяц до окончания" - 30 дней, "розовым" полтора месяца - 45 дней.
