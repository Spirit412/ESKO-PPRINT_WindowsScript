import argparse
import logging

import ConMySQL

# inputs, outputFolder, params
# parser = argparse.ArgumentParser(description='inputs, outputFolder')
# parser.add_argument('inputs', type=str, help='Input dir for xls file')
# parser.add_argument('outputFolder', type=str, help='Output dir for xml file')
# args = parser.parse_args()



root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)  # or whatever
handler = logging.FileHandler('\\\\Server-esko\\ae_base\TEMP-Shuttle-IN\\logFile.txt', 'w', 'utf-8')  # or whatever
handler.setFormatter(
    logging.Formatter(u'%(filename)s[LINE:%(lineno)d]# %(levelname)-8s [%(asctime)s]  %(message)s'))  # or whatever
root_logger.addHandler(handler)

d = ConMySQL.get('A317')
print(d)
with open('\\\\Server-esko\\ae_base\TEMP-Shuttle-IN\\logFile.txt', "w", encoding='utf8') as f:
    f.write('- - - - - - данные из БД по штампу: ' + d['IDCUT'] + ' - - - - - ')
    f.close()
#
# print('- - - - - - данные из БД по штампу: ' + d['IDCUT'] + ' - - - - - ')
# for key, value in d.items():
#     tab = 1
#     probelkey = 0
#     probelvalue = 0
#     # print(str(key).__len__())
#     if str(key).__len__() < 15:
#         probelkey = int(15 - str(key).__len__())
#     else:
#         probelkey = 0
#     if str(value).__len__() < 15:
#         probelvalue = int(15 - str(value).__len__())
#     else:
#         probelvalue = 0
#     f.write('\t' + '| ' + key + str(probelkey * ' ') + '\t' + '->' + '\t' + str(value) + str(probelvalue * ' ') + ' |')
#
#     f.close()