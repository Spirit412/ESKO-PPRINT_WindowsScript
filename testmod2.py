import testmod1
import logging
import ConMySQL


def main():
    """
    The main entry point of the application
    """
    logger = logging.getLogger('testmod1')
    logger2 = logging.getLogger('ConMySQL1')
    logger3 = logging.getLogger('testmod2')
    logger.setLevel(logging.INFO)
    logger2.setLevel(logging.INFO)
    logger3.setLevel(logging.INFO)

    # create the logging file handler
    fh = logging.FileHandler("new_snake.log", 'w', 'utf-8')
    fh2 = logging.FileHandler("new_snake.log", 'a', 'utf-8')
    fh3 = logging.FileHandler("new_snake.log", 'a', 'utf-8')
    formatter = logging.Formatter(u'[%(asctime)s] - %(filename)s - - [LINE:%(lineno)d]# - %(message)s')
    formatter2 = logging.Formatter(u'[%(asctime)s] - %(filename)s - %(name)s - [LINE:%(lineno)d]# - %(levelname)4s - %(message)s')
    formatter3 = logging.Formatter(u'%(message)s')
    fh.setFormatter(formatter)
    fh2.setFormatter(formatter2)
    fh3.setFormatter(formatter3)

    # add handler to logger object
    logger.addHandler(fh)
    logger2.addHandler(fh2)
    logger3.addHandler(fh3)

    d = ConMySQL.get('A1270')
    print(testmod1.add(7, 8))
    print(d)
    d1={**d}

    replacements = {'ID': 'ID', 'IDCUT': 'Штамп', 'zub': 'зуб', 'HPrint': 'длина печати', 'HPolimer': 'Длина полимера 1.14', 'HDist': 'Дист 1.14', 'Hpolimer_17': 'Длина полимера 1.7', 'Hdist_17': 'Дист 1.7', 'Vsheet': 'Высота штампа', 'HCountItem': 'Эт-к по длине', 'VCountItem': 'Эт-к по ширина'}
    for i in list(d1):
        if i in replacements:
            d1[replacements[i]] = d1.pop(i)
    print(d)
    print(d1)

    logger3.info('- - - - - - данные из БД по штампу: ' + d['IDCUT'] + ' - - - - - ')
    # словарь замен: ключ - исходный ключ из d1, значение - на какой ключ его меняем
    for key, value in d1.items():
        tab = 2
        probelkey = 14
        probelvalue = 14
        # print(str(key).__len__())
        if str(key).__len__() < 20:
            probelkey = int(20 - str(key).__len__())
        else:
            probelkey = 0
        if str(value).__len__() < 10:
            probelvalue = int(10 - str(value).__len__())
        else:
            probelvalue = 0
        logger3.info('\t' + '| ' + key + str(probelkey * ' ') + '\t' + '->' + '\t' + str(value) + str(probelvalue * ' ') + ' |')

if __name__ == "__main__":
    main()

