import pymysql.cursors  
 
# Функция возвращает connection.
def getConnection():
     
# Подключиться к базе данных.
connection = pymysql.connect(host='127.0.0.1',
                             user='root',
                             password='',                             
                             db='pprint',
                             charset='utf8',
                             cursorclass=pymysql.cursors.DictCursor)
return connection
