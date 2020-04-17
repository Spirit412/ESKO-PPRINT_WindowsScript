import argparse
import pymysql
import xlrd
import os
import shutil
from xml.dom import minidom
import logging


try:
    import configparser
except ImportError:
    import ConfigParser as configparser

config = configparser.ConfigParser()  # создаём объекта парсера
config.read("config.ini")  # читаем конфиг

 # inputs, outputFolder, params
parser = argparse.ArgumentParser(description='inputs, outputFolder')
parser.add_argument('inputs', type=str, help='Input dir for xls file')
parser.add_argument('outputFolder', type=str, help='Output dir for xml file')
args = parser.parse_args()
logFile = str(args.outputFolder + "\\log.txt")
print(logFile)
logging.basicConfig(
    format = u'%(filename)s[LINE:%(lineno)d]# %(levelname)-8s [%(asctime)s]  %(message)s',
    level = logging.DEBUG,
    filename = logFile,
    filemode='w'
)


# print(args.inputs)
# print(args.outputFolder)



xls_file = args.inputs
xls_workbook = xlrd.open_workbook(str(xls_file))
xls_sheet = xls_workbook.sheet_by_index(0)


# Параметры из XLS файла в переменные.
# простая функция запроса к данным именной ячейки
def readNameCell(nameCell, xls):
    """
    Функция получает на входе имя ячейки и адрес XLS файла
    на выходе массив
    0 - строка ячейки
    1 - столбец ячейки
    2 - абсолютный адрес ячейки
    3 - имя ячейки
    4 - данные из ячейки
    """
    book = xlrd.open_workbook(xls)
    nameObj = book.name_and_scope_map.get((nameCell.lower(), -1))  # имя маленькими буквами
    r = [0, 1, 2, 3, 4]
    r[0] = nameObj.area2d()[1]
    r[1] = nameObj.area2d()[3]
    r[2] = nameObj.result.text
    r[3] = nameObj.name
    r[4] = nameObj.cell().value
    return (r)


logging.info(u'проверка')
logging.error(u'проверка')



JobNamber = readNameCell('JobNamber', xls_file)[4].strip()
CustomerName = readNameCell('CustomerName', xls_file)[4].strip()
CutTools = readNameCell('CutTools', xls_file)[4].strip()
StampManufacturer = readNameCell('StampManufacturer', xls_file)[4].strip()
Liniatura = readNameCell('Liniatura', xls_file)[4]
PrintMashine = readNameCell('PrintMashine', xls_file)[4]
ThicknessPolymer = readNameCell('ThicknessPolymer', xls_file)[4]
DieShape = "file://server-esko/AE_BASE/CUT-TOOLS/" + CutTools + ".cf2"
DieShapeMFG = "file://server-esko/AE_BASE/CUT-TOOLS/" + CutTools + ".MFG"
Bleed = str(readNameCell('Bleed', xls_file)[4]).strip()
Bleed = Bleed.replace('.', ',')
# Подключиться к базе данных.
try:
    connection = pymysql.connect(host='esko',
                                 user='root',
                                 port=3360,
                                 password='',
                                 db='pprint',
                                 charset='utf8',
                                 cursorclass=pymysql.cursors.DictCursor)
    with open(logFile, "w", encoding='utf8') as f:
        f.write("connect successful!!")
        f.close()

except:
    # запись в лог файл
    with open(logFile, "w", encoding='utf8') as f:
        f.write("ERROR connect MySQL")
        f.close()

try:
    with connection.cursor() as cursor:
        # SQL
        cursor.execute("SELECT * FROM tools  WHERE IDCUT=(%s)", (CutTools))
        # Получаем результат сделанного запроса к БД MySQL
    rows = cursor.fetchall()
    # print()
    #    s.replace(',', '.')
    for row in rows:
        Zub = row['zub']
        DPrint = row['HPrint']
        Polimer = row['HPolimer']
        Distorsia = row['HDist']
        # для полимера 1.7
        Polimer17 = row['Hpolimer_17']
        Distorsia17 = row['Hdist_17']
        Vsheet = row['Vsheet']

        HCountItem = row['HCountItem']
        VCountItem = row['VCountItem']

        # отступы
        HGap = row['HGap']
        VGap = row['VGap']
finally:
    # Закрыть соединение (Close connection).
    connection.close()


# определяем способ заполнения штампа строками/столбцами
Razmeshenie = readNameCell('Razmeshenie', xls_file)[4]
# если строки = h, если столбцами = v
if Razmeshenie == "строки":
    SequenceDirection = "h"
    Quantity = HCountItem
else:
    SequenceDirection = "v"
    Quantity = VCountItem


doc = minidom.Document()
root = doc.createElement('JOBS')
doc.appendChild(root)

leaf = doc.createElement('HCountItem')
text = doc.createTextNode(str(HCountItem))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('VCountItem')
text = doc.createTextNode(str(VCountItem))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Zub')
text = doc.createTextNode(str(Zub).strip())
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Distorsia')
text = doc.createTextNode(str(Distorsia))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Polimer')
text = doc.createTextNode(str(Polimer))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('DPrint')
text = doc.createTextNode(str(DPrint))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Vsheet')
text = doc.createTextNode(str(Vsheet))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Distorsia17')
text = doc.createTextNode(str(Distorsia17))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Polimer17')
text = doc.createTextNode(str(Polimer17))
leaf.appendChild(text)
root.appendChild(leaf)

# из экселевского файла
leaf = doc.createElement('JobNamber')
text = doc.createTextNode(str(JobNamber).strip())
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('CustomerName')
text = doc.createTextNode(str(CustomerName).strip())
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('CutTools')
text = doc.createTextNode(str(CutTools).strip())
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('StampManufacturer')
text = doc.createTextNode(str(StampManufacturer).strip())
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Razmeshenie')
text = doc.createTextNode(str(Razmeshenie))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('SequenceDirection')
text = doc.createTextNode(str(SequenceDirection))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Liniatura')
text = doc.createTextNode(str(Liniatura).strip())
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('PrintMashine')
text = doc.createTextNode(str(PrintMashine).strip())
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('ThicknessPolymer')
text = doc.createTextNode(str(ThicknessPolymer).strip())
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('DieShape')
text = doc.createTextNode(str(DieShape))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('DieShapeMFG')
text = doc.createTextNode(str(DieShapeMFG))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Bleed')
text = doc.createTextNode(str(Bleed))
leaf.appendChild(text)
root.appendChild(leaf)

# JOBы для спуска
iID = 1
x = readNameCell('FirstID', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
while x != "IndexError":
    cr = xls_sheet.cell(x, 1).value.strip()
    cr = os.path.basename(cr)
    textTest = str(iID) + "_" + cr
    job = doc.createElement('JOB')
    root.appendChild(job)
    leaf = doc.createElement('FileName')
    text = doc.createTextNode(textTest)
    leaf.appendChild(text)
    job.appendChild(leaf)
    leaf = doc.createElement('Quantity')
    text = doc.createTextNode(str(Quantity).strip())
    leaf.appendChild(text)
    job.appendChild(leaf)
    x += 1
    iID += 1
    # проверяем следующую ячейк. если она пустая, и запрос выдаёт ошибку IndexError, выходим из цикла
    try:
        cr = xls_sheet.cell(x, 1).value
    except IndexError:
        break


# делаем папку inPDF с проверкой
InPDFfolder = os.path.dirname(args.inputs) + "\\inPDF"
if os.path.exists(InPDFfolder) != True:
    os.mkdir(InPDFfolder, mode=0o777, dir_fd=None)

# а тут делаем список файлов, с углом поворота файла.
iID = 1
x = readNameCell('FirstID', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
while x != "IndexError":
    cr = xls_sheet.cell(x, 1).value.strip()
    textTest = cr
    if os.path.isabs(cr) == True:
        textTest = os.path.basename('r' + cr)

    leaf = doc.createElement('File')
    leaf.setAttribute("ID", str(iID).strip())
    leaf.setAttribute("Angle", str(int(xls_sheet.cell(x, 2).value)).strip())
    text = doc.createTextNode(textTest)
    print(textTest)
    leaf.appendChild(text)
    root.appendChild(leaf)


    #копируем файлы в папку .../inPDF/
    cr = xls_sheet.cell(x, 1).value.strip()
    if os.path.isabs(cr) == True:
        shutil.copy(cr, InPDFfolder)
    x += 1
    iID += 1
    # проверяем следующую ячейк. если она пустая, и запрос выдаёт ошибку IndexError, выходим из цикла
    try:
        cr = xls_sheet.cell(x, 1).value
    except IndexError:
        break
# input("Press Enter to continue...")
# делаем папку JobNumber_spusk с проверкой
FolderJobNumber = args.outputFolder
if os.path.exists(FolderJobNumber) != True:
    os.mkdir(FolderJobNumber, mode=0o777, dir_fd=None)
print(FolderJobNumber)
xml_out = FolderJobNumber + "\\" + str(JobNamber).strip() + ".xml"
print(xml_out)
# xml_out = "C:\\Temp\\"+ str(JobNamber) + ".xml"






xml_str = doc.toprettyxml(indent="  ")
with open(str(xml_out), "w", encoding='utf8') as f:
    f.write(xml_str)
    f.close()
