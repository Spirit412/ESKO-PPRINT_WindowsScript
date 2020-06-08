import pymysql
import logging

try:
    import configparser
except ImportError:
    import ConfigParser as configparser
config = configparser.ConfigParser()  # создаём объекта парсера
config.read("config.ini")  # читаем конфиг


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
    try:
        logger = logging.getLogger('ConMySQL')
        connection = pymysql.connect(host=config.get('MySQL', 'host'),
                                     user=config.get('MySQL', 'user'),
                                     password=config.get('MySQL', 'password'),
                                     db=config.get('MySQL', 'db'),
                                     port=int(config.getint('MySQL', 'port')),
                                     # port=3360,
                                     charset=config.get('MySQL', 'charset'),
                                     cursorclass=pymysql.cursors.DictCursor)
        logger.info("Подсоединение к серверу успешно выполнено")
        print("Подсоединение к серверу успешно выполнено")

    except pymysql.err.OperationalError as pmye:
        logger.error("Ошибка подключения к БД: {}".format(pmye))
    except configparser.NoOptionError as pmye2:
        logger.error("Ошибки в настройке доступа к конфигу: {}".format(pmye2))
        exit()
    try:
        with connection.cursor() as cursor:
            # SQL
            cursor.execute("SELECT * FROM tools  WHERE IDCUT=(%s)", (cut_tool))
            # Получаем результат сделанного запроса к БД MySQL. Формат list
            rows = cursor.fetchall()
    except:
        logger.error("SQL запрос не верный")
        exit()
    finally:
        connection.close()
        # START Преобразуем тип rows list в dict
    for rows in rows:
        get_out = rows
    return get_out
