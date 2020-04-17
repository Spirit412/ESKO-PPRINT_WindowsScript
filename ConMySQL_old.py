import pymysql

# import logging
try:
    import configparser
except ImportError:
    import ConfigParser as configparser

# logging.basicConfig(
#     format=u'%(filename)s[LINE:%(lineno)d]# %(levelname)-8s [%(asctime)s]  %(message)s',
#     level=logging.DEBUG,
# )

config = configparser.ConfigParser()  # создаём объекта парсера
config.read("config.ini")  # читаем конфиг


class ConMySQL:
    """
Подключение к БД MySQL. Настройки подключения внутри класса
    host_con - адрес. По умолчанию 127.0.0.1
    user_con - имя пользователя
    password_con - пароль
    port - порт подключения к БД
    db_con - имя БД
    charset_con - кодировка. По умолчанию UTF8
    """

    def __init__(self, cutname):
        self.cutname = str(cutname)
        # self.row = row
        # self.rows = rows
        try:
            connection = pymysql.connect(host=config.get('MySQL', 'host'),
                                         user=config.get('MySQL', 'user'),
                                         password=config.get('MySQL', 'password'),
                                         db=config.get('MySQL', 'db'),
                                         port=3360,
                                         charset=config.get('MySQL', 'charset'),
                                         cursorclass=pymysql.cursors.DictCursor)
            print("connect successful!!")

        except pymysql.err.OperationalError as pmye:
            print("Ошибка подключения к БД: {}".format(pmye))
        except configparser.NoOptionError as pmye2:
            print("Ошибки в настройке доступа к конфигу: {}".format(pmye2))
            exit()

        try:
            with connection.cursor() as cursor:
                # SQL
                cursor.execute("SELECT * FROM tools  WHERE IDCUT=(%s)", (self.cutname))
                # Получаем результат сделанного запроса к БД MySQL. Формат list
            self.rows = cursor.fetchall()
            for self.rows in self.rows:
                get_out = self.rows

        except:
            logging.critical("SQL запрос не верный")
            exit()
        finally:
            # Закрыть соединение (Close connection).
            connection.close()
        return dict(get_out)


    def get(self):
        """
вывод данных в формате словаря dict
        """
        # START Преобразуем тип rows list в dict
        try:
            for self.rows in self.rows:
                get_out = self.rows
            return get_out
        except AttributeError:
            logging.critical("Ошибка: 'ConMySQL' object has no attribute 'rows'")
            exit()

if __name__ == "__main__":
    d = ConMySQL.get('A182')
    print(ConMySQL('A182'))
