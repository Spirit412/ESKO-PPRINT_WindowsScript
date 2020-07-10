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
handler = logging.FileHandler('\\\\esko\\ae_base\TEMP-Shuttle-IN\\logFile.txt', 'w', 'utf-8')  # or whatever
handler.setFormatter(
    logging.Formatter(u'%(filename)s[LINE:%(lineno)d]# %(levelname)-8s [%(asctime)s]  %(message)s'))  # or whatever
root_logger.addHandler(handler)

d = ConMySQL.get('A317')
print(d)
with open('\\\\esko\\ae_base\TEMP-Shuttle-IN\\logFile.txt', "w", encoding='utf8') as f:
    # f.write('- - - - - - данные из БД по штампу: ' + d['IDCUT'] + ' - - - - - ')
    # f.close()

    print('|{a}{b:<}|'.format(a='данные из БД по штампу: ', b=d['IDCUT']))
    for key, value in d.items():

        print('| {key:_<15}{value:_>15} |'.format(key=key, value=value), end='\n')
        f.write('| {key:_<15}{value:_>15} |\t\n'.format(key=key, value=value))
