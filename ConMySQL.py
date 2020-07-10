import pymysql
import logging
import configparser
import os


def get(cut_tool):
    """
Подключение к БД MySQL. Настройки подключения внутри класса
    host_con - адрес. По умолчанию 127.0.0.1
    user_con - имя пользователя
    password_con - пароль
    port - порт подключения к БД
    db_con - имя БД
    charset_con - кодировка. По умолчанию UTF8
    """
    logger = logging.getLogger('ConMySQL')
    config = configparser.ConfigParser()  # создаём объекта парсера
    thisfolder = os.path.dirname(os.path.abspath(__file__))
    inifile = os.path.join(thisfolder, 'config.ini')
    config.read(inifile)  # читаем конфиг
    try:
        connection = pymysql.connect(host=config.get("mysql", "host"),
                                     user=config.get("mysql", "user"),
                                     password=config.get("mysql", "password"),
                                     db=config.get("mysql", "db"),
                                     port=config.getint("mysql", "port"),
                                     charset=config.get("mysql", "charset"),
                                     cursorclass=pymysql.cursors.DictCursor)
        logger.info("Подсоединение к серверу успешно выполнено")

    except configparser.NoSectionError as err1:
        logger.error("Ошибка configparser. Нет запрашиваемой секции в конфиг-файле: {}".format(err1))
        exit()
    except pymysql.err.OperationalError as err2:
        logger.error("Ошибка подключения к БД: {}".format(err2))
        exit()
    except configparser.NoSectionError as err3:
        logger.error("Ошибка configparser в настройке доступа к конфигу: {}".format(err3))
        exit()
    except configparser.NoOptionError as err4:
        logger.error("Ошибка configparser. Нет опции секции в конфиг-файле: {}".format(err4))
        exit()
    except:
        logger.error("Ошибка подключения к БД!")
        exit()
    try:
        with connection.cursor() as cursor:
            # SQL
            # cursor.execute("SELECT * FROM tools  WHERE IDCUT=(%s)", (cut_tool))
            cursor.execute("SELECT cut.*, zub.HPrint, zub.HPolimer, zub.Hpolimer_17 FROM cut LEFT JOIN zub ON cut.zub "
                           "= zub.zub WHERE IDCUT = (%s)", (cut_tool))
            # Получаем результат сделанного запроса к БД MySQL. Формат list
            rows = cursor.fetchall()
    except:
        logger.error("SQL запрос не верный")
        # exit()
    finally:
        connection.close()
        # START Преобразуем тип rows list в dict
    for rows in rows:
        get_out = rows
    assert isinstance(get_out, object)
    return get_out


def get_list_id_cut():
    """
 Подключение к БД MySQL. Настройки подключения внутри класса
     host_con - адрес. По умолчанию 127.0.0.1
     user_con - имя пользователя
     password_con - пароль
     port - порт подключения к БД
     db_con - имя БД
     charset_con - кодировка. По умолчанию UTF8
     """
    logger = logging.getLogger('ConMySQL')
    config = configparser.ConfigParser()  # создаём объекта парсера
    thisfolder = os.path.dirname(os.path.abspath(__file__))
    inifile = os.path.join(thisfolder, 'config.ini')
    config.read(inifile)  # читаем конфиг
    try:
        connection = pymysql.connect(host=config.get("mysql", "host"),
                                     user=config.get("mysql", "user"),
                                     password=config.get("mysql", "password"),
                                     db=config.get("mysql", "db"),
                                     port=config.getint("mysql", "port"),
                                     charset=config.get("mysql", "charset"),
                                     cursorclass=pymysql.cursors.DictCursor)
        logger.info("Подсоединение к серверу успешно выполнено")

    except configparser.NoSectionError as err1:
        logger.error("Ошибка configparser. Нет запрашиваемой секции в конфиг-файле: {}".format(err1))
        exit()
    except pymysql.err.OperationalError as err2:
        logger.error("Ошибка подключения к БД: {}".format(err2))
        exit()
    except configparser.NoSectionError as err3:
        logger.error("Ошибка configparser в настройке доступа к конфигу: {}".format(err3))
        exit()
    except configparser.NoOptionError as err4:
        logger.error("Ошибка configparser. Нет опции секции в конфиг-файле: {}".format(err4))
        exit()
    except:
        logger.error("Ошибка подключения к БД!")
        exit()
    try:
        with connection.cursor() as cursor:
            # SQL
            cursor.execute("SELECT IDCUT FROM cut")
            # Получаем результат сделанного запроса к БД MySQL. Формат list
            rows = cursor.fetchall()
    except:
        logger.error("SQL запрос не верный")
        # exit()
    finally:
        connection.close()
    #     # START Преобразуем тип rows list в dict
    # for row in rows:
    #     list_id = list()
    #     list_id.append(row)

    # assert isinstance(get_out, object)
    # return list_id
    return rows


if __name__ == "__main__":
    d = get('D0076')
    print(d)
