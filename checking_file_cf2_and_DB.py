import logging
import ConMySQL

import os
from pathlib import Path

# Сравниваем список файлов CF2 с БД штампов и выводим отсутствующие в БД штампы


directory = '//Server-esko/AE_BASE/CUT-TOOLS/'
#Получаем список файлов в переменную files
files = os.listdir(directory)
files = filter(lambda x: x.endswith('.cf2'), files)
files = [os.path.splitext(os.path.basename(fn))[0] for fn in files]

# print(files)

all_id_dict = ConMySQL.get_list_id_cut()
all_id_list = list()
for id in all_id_dict:
    all_id_list.append(id['IDCUT'])
# print(all_id_list)

C = list(set(all_id_list) - set(files))
print(sorted(C))
print(len(files))