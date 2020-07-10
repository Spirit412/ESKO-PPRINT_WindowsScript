import argparse
import xlrd
import os
import shutil
from xml.dom import minidom
import ConMySQL
import logging

try:
    import configparser
except ImportError:
    import ConfigParser as configparser






# START настройка чтения конфиг файла
config = configparser.ConfigParser()  # создаём объекта парсера
thisfolder = os.path.dirname(os.path.abspath(__file__))
inifile = os.path.join(thisfolder, 'config.ini')
config.read(inifile)  # читаем конфиг
# END настройка чтения конфиг файла


# inputs, outputFolder, params
parser = argparse.ArgumentParser(description='inputs, outputFolder')
parser.add_argument('inputs', type=str, help='Input dir for xls file')
parser.add_argument('outputFolder', type=str, help='Output dir for xml file')
args = parser.parse_args()
logFile = str(args.outputFolder + "\\log.txt")
print(logFile)

# START настройка логгирования модуля и основного скрипта

logger = logging.getLogger('основной')
logger2 = logging.getLogger('ConMySQL')
logger3 = logging.getLogger('основной')

logger.setLevel(logging.INFO)
logger2.setLevel(logging.INFO)
logger3.setLevel(logging.INFO)

# create the logging file handler
fh = logging.FileHandler(logFile, 'w', 'windows-1251')
fh2 = logging.FileHandler(logFile, 'a+', 'windows-1251')
fh3 = logging.FileHandler(logFile, 'a+', 'windows-1251')

formatter = logging.Formatter(u'[%(asctime)s] - %(filename)s - [LINE:%(lineno)d]# - %(levelname)4s - %(message)s')
formatter2 = logging.Formatter(u'[%(asctime)s] - %(filename)s - [LINE:%(lineno)d]# - %(levelname)4s - %(message)s')
formatter3 = logging.Formatter(u'%(message)s')
fh.setFormatter(formatter)
fh2.setFormatter(formatter2)
fh3.setFormatter(formatter3)

# add handler to logger object
logger.addHandler(fh)
logger2.addHandler(fh)
logger3.addHandler(fh3)

xls_file = args.inputs
xls_workbook = xlrd.open_workbook(str(xls_file))
xls_sheet = xls_workbook.sheet_by_index(0)


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


# START чтение данных из XLS файла в переменные
JobNamber = readNameCell('JobNamber', xls_file)[4].strip()
CustomerName = readNameCell('CustomerName', xls_file)[4].strip()
CutTools = readNameCell('CutTools', xls_file)[4].strip()
StampManufacturer = readNameCell('StampManufacturer', xls_file)[4].strip()
Liniatura = readNameCell('Liniatura', xls_file)[4]
PrintMashine = readNameCell('PrintMashine', xls_file)[4]
ThicknessPolymer = readNameCell('ThicknessPolymer', xls_file)[4]
DieShape = "file://esko/AE_BASE/CUT-TOOLS/" + CutTools + ".cf2"
DieShapeMFG = "file://esko/AE_BASE/CUT-TOOLS/" + CutTools + ".MFG"
Bleed = str(readNameCell('Bleed', xls_file)[4]).strip()
Bleed = Bleed.replace('.', ',')
# END чтение данных из XLS файла в переменные


# Подключиться к базе данных.

# Подключиться к базе данных и получить словарь всех данных параметра.
try:
    CutTool = ConMySQL.get(CutTools)
except:
    logger.info('нет соединения с БД')
# START ------------------------Вывод данных таблицей в лог файл--------------------------------#
# создаём дубликат словаря длоя оформления таблици в логе
try:
    d1 = {**CutTool}
except NameError:
    logger.error("Ошибка: {}".format(NameError))
    exit()
# словарь замены ключей
replacements = {'ID': 'ID', 'IDCUT': 'Штамп', 'zub': 'зуб', 'HPrint': 'длина печати',
                'HPolimer': 'Длина полимера 1.14', 'Hpolimer_17': 'Длина полимера 1.7',
                'Vsheet': 'Высота штампа', 'HCountItem': 'Эт-к по длине',
                'VCountItem': 'Эт-к по ширина'}
# цикл замены ключей в словаре.
for i in list(d1):
    if i in replacements:
        d1[replacements[i]] = d1.pop(i)

logger3.info(str('\t' + 'Таблица данных из БД на штамп ' + CutTool['IDCUT']))

for key, value in d1.items():
    # tab = 2
    # probelkey = 5
    # probelvalue = 5
    # if str(key).__len__() < 30:
    #     probelkey = int(30 - str(key).__len__())
    # else:
    #     probelkey = 0
    # if str(value).__len__() < 10:
    #     probelvalue = int(10 - str(value).__len__())
    # else:
    #     probelvalue = 0
    # logger3.info('| ' + key + str(probelkey * '. ') + '  ->  ' + str(value) + str(probelvalue * ' ') + ' |')
    logger3.info('| {key:_<15}{value:_>15} |\t\n'.format(key=key, value=value))

# END ------------------------Вывод данных таблицей в лог файл--------------------------------#



Zub = CutTool['zub']
DPrint = CutTool['HPrint']
Polimer = CutTool['HPolimer']
Distorsia = round((Polimer / DPrint) * 100, 4)
# для полимера 1.7
Polimer17 = CutTool['Hpolimer_17']
Distorsia17 = round((Polimer17 / DPrint) * 100, 4)
Vsheet = CutTool['Vsheet']

HCountItem = CutTool['HCountItem']
VCountItem = CutTool['VCountItem']

# отступы
HGap = CutTool['HGap']
VGap = CutTool['VGap']







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

    # копируем файлы в папку .../inPDF/
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
