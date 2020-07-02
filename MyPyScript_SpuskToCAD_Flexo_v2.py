import argparse
import logging
import os
import shutil
from random import randint
import xlrd
from reportlab.lib.colors import PCMYKColorSep
from reportlab.lib.units import mm
from reportlab.pdfgen.canvas import Canvas

from xml.dom import minidom
import ConMySQL

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


inputsFolder = os.path.dirname(args.inputs)

outputFolder = args.outputFolder
# logFile = outputFolder + "\\log.txt"



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

# END настройка логгирования модуля и основного скрипта

try:
    xls_file = args.inputs
    xls_workbook = xlrd.open_workbook(str(xls_file))
    xls_sheet = xls_workbook.sheet_by_index(0)
except:
    logger.error(u'ошибка чтения XLS файла')


def round(xin, y=0):
    """ Функция математического округления """
    m = int('1' + '0' * y)  # multiplier - how many positions to the right
    q = xin * m  # shift to the right by multiplier
    c = int(q)  # new number
    i = int((q - c) * 10)  # indicator number on the right
    if i >= 5:
        c += 1
    return c / m


# Параметры из XLS файла в переменные.
# простая функция запроса к данным именной ячейки
try:
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
        try:
            r[0] = nameObj.area2d()[1]
            r[1] = nameObj.area2d()[3]
            r[2] = nameObj.result.text
            r[3] = nameObj.name
            r[4] = nameObj.cell().value
        except AttributeError:
            logger.error(u'не найдена именованная ячейка по запрашиваемому имени')
        return (r)
except:
    logger.error(u'не правильные атрибуты XLS файла. Ошибка функции readNameCell')

table = str.maketrans("", "", "?()!@#/$%^&*+|+\:;[]{}<>")
"""
список запрещённых символов в названии файла.
"""


JobNamber = str(readNameCell('JobNamber', xls_file)[4]).replace(".0", "")
# удаляем запрещенные символы из номера заказа
JobNamber = JobNamber.translate(table).replace(".0", "")
CustomerName = readNameCell('CustomerName', xls_file)[4]
CutTools = readNameCell('CutTools', xls_file)[4]
StampManufacturer = readNameCell('StampManufacturer', xls_file)[4]
Liniatura = readNameCell('Liniatura', xls_file)[4]
PrintMashine = readNameCell('PrintMashine', xls_file)[4]
ThicknessPolymer = readNameCell('ThicknessPolymer', xls_file)[4]
TypyPolimer = readNameCell('TypyPolimer', xls_file)[4]
DieShape = "file://esko/AE_BASE/CUT-TOOLS/" + CutTools + ".cf2"
DieShapeMFG = "file://esko/AE_BASE/CUT-TOOLS/" + CutTools + ".MFG"
Bleed = str(readNameCell('Bleed', xls_file)[4])
Bleed = Bleed.replace('.', ',')
PositionMark_6x4 = readNameCell('PositionMark_6x4', xls_file)[4]
RotateCUT = readNameCell('RotateCUT', xls_file)[4]
UpOffset = readNameCell('UpOffset', xls_file)[4]
BotOffset = readNameCell('BotOffset', xls_file)[4]
psFile = readNameCell('psFile', xls_file)[4]
onebitTiff = readNameCell('onebitTiff', xls_file)[4]
addDGC = readNameCell('addDGC', xls_file)[4]


try:
    mirror_layout = readNameCell('mirror_layout', xls_file)[4]
except AttributeError:
    mirror_layout = ''


try:
    OutFileName = readNameCell('OutFileName', xls_file)[4]
except AttributeError:
    OutFileName = os.path.splitext(os.path.basename(inputsFolder))[0]

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
                'HPolimer': 'Длина полимера 1.14', 'HDist': 'Дист 1.14', 'Hpolimer_17': 'Длина полимера 1.7',
                'Hdist_17': 'Дист 1.7', 'Vsheet': 'Высота штампа', 'HCountItem': 'Эт-к по длине',
                'VCountItem': 'Эт-к по ширина'}
# цикл замены ключей в словаре.
for i in list(d1):
    if i in replacements:
        d1[replacements[i]] = d1.pop(i)

logger3.info(str('\t' + 'Таблица данных из БД на штамп ' + CutTool['IDCUT']))

for key, value in d1.items():
    tab = 2
    probelkey = 5
    probelvalue = 5
    if str(key).__len__() < 30:
        probelkey = int(30 - str(key).__len__())
    else:
        probelkey = 0
    if str(value).__len__() < 10:
        probelvalue = int(10 - str(value).__len__())
    else:
        probelvalue = 0
    logger3.info('| ' + key + str(probelkey * '. ') + '  ->  ' + str(value) + str(probelvalue * ' ') + ' |')

# END ------------------------Вывод данных таблицей в лог файл--------------------------------#

Zub = CutTool['zub']
DPrint = CutTool['HPrint']
Polimer = CutTool['HPolimer']

# вычисляем % дисторсии до 4го знака после. В старой версии % брался из БД
# для полимера 1.7
Distorsia = round((Polimer / DPrint) * 100, 4)
Polimer17 = CutTool['Hpolimer_17']

# вычисляем % дисторсии до 4го знака после. В старой версии % брался из БД
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
text = doc.createTextNode(str(Zub))
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
text = doc.createTextNode(str(JobNamber).replace(".0", ""))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('CustomerName')
text = doc.createTextNode(str(CustomerName.strip()))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('CutTools')
text = doc.createTextNode(str(CutTools))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('StampManufacturer')
text = doc.createTextNode(str(StampManufacturer.strip()))
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
text = doc.createTextNode(str(int(Liniatura)))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('PrintMashine')
text = doc.createTextNode(str(PrintMashine))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('ThicknessPolymer')
text = doc.createTextNode(str(ThicknessPolymer))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('TypyPolimer')
text = doc.createTextNode(str(TypyPolimer))
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

leaf = doc.createElement('PositionMark_6x4')
text = doc.createTextNode(str(PositionMark_6x4))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('RotateCUT')
text = doc.createTextNode(str(int(RotateCUT)))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('UpOffset')
text = doc.createTextNode(str(UpOffset))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('BotOffset')
text = doc.createTextNode(str(BotOffset))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('psFile')
text = doc.createTextNode(str(psFile))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('onebitTiff')
text = doc.createTextNode(str(onebitTiff))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('addDGC')
text = doc.createTextNode(str(addDGC))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('mirror_layout')
text = doc.createTextNode(str(mirror_layout))
leaf.appendChild(text)
root.appendChild(leaf)

# OutFileName - имя выходных файлов
if OutFileName == "":
    OutFileName = JobNamber
else:
    # защита от использования запрещённых символов
    table = str.maketrans("", "", "?()!@#$%^&*+|+\/:;[]{}<>")  # список запрещённых символов в названии файла.
    OutFileName = OutFileName.translate(table)  # удаляем запрещенные символы из номера заказа
print("Имя файла на выходе " + OutFileName)

leaf = doc.createElement('OutFileName')
text = doc.createTextNode(str(OutFileName))
leaf.appendChild(text)
root.appendChild(leaf)

# ---------------------функция проверки на наличие недопустимых символов в строке--------------------------------#
import re


def is_ok(text_in):
    """
    функция проверки на наличие недопустимых символов в строке.
    Возвращает список найденных запрещенных символов
    """
    test_str = text_in
    exception_chars = '?()!@#$%^&*+|+\/:;{}<>'
    # exception_chars = '?()!@#$%^&*+|+\/:;[]{}<>'
    find_exceptions = re.compile('([{}])'.format(exception_chars))
    res = find_exceptions.findall(test_str)
    return res


# ---------------------проверка на наличие недопустимых символов в названии файла--------------------------------#


# делаем папку inPDF с проверкой. При отсутствии - создаём
InPDFfolder = os.path.dirname(args.inputs) + "\\" + config.get("ESKO", "inPDF")
print(InPDFfolder)
if os.path.exists(InPDFfolder) != True:
    os.mkdir(InPDFfolder, mode=0o777, dir_fd=None)
    logger.info(f'Папка {config.get("ESKO", "inPDF")} отсутствует. Создаём')

# JOBы для спуска
# получим словарь дизайнов с параметрами из XLS
iID = 1
mydict = {}
# возвращаем номер строки первой ячейки в таблице файлов дизайнов
x = readNameCell('FirstID', xls_file)[0]
while x != "IndexError":
    FileName = xls_sheet.cell(x, 1).value.strip()  # значение из ячейки с именем дизайна
    CellName = xlrd.cellname(x, 1)

    # Проверка на наличие запрещенных символов в ячейке
    if is_ok(FileName):
        logger.error('Ячейка {} - содержит запрещенный символ: {}'.format(CellName, is_ok(FileName)))
    try:
        other_var = readNameCell('file_angle_rotate', xls_file)[1]
    except:
        logger.error('Ошибка')

    AngleFile = str(int(xls_sheet.cell(x, other_var).value))  # значение у
    mydict[iID] = {'File': FileName, 'Angle': AngleFile, 'Quantity': Quantity}
    try:
        # копируем файлы в папку .../inPDF/
        cr = xls_sheet.cell(x, 1).value.strip()
        print(cr)
        print(InPDFfolder)
        if os.path.isabs(cr):
            shutil.copy(cr, InPDFfolder)
    except FileNotFoundError:
        logger.error(f'Не найдены файлы {cr} для копирования в папку {config.get("ESKO", "inPDF")}')

    # Формируем список
    x += 1
    iID += 1
    # проверяем следующую ячейк. если она пустая, и запрос выдаёт ошибку IndexError, выходим из цикла
    try:
        FileName = xls_sheet.cell(x, 1).value
    except IndexError:
        break

# START проверяем значения ключей в словаре. Если все равны, то True и Quantity = maximum
values = [(d['File'], d['Angle']) for d in mydict.values()]
if all(f == values[0][0] and a == values[0][1] for f, a in values):
    allFileSame = True
    logger.info("наспуске один макет с одинаковым углом поворота. Параметр Quantity = Maximum")
else:
    allFileSame = False

if allFileSame == True:
    x = readNameCell('FirstID', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
    FileName = xls_sheet.cell(x, 1).value.strip()
    FileName = os.path.basename(FileName)
    FileName = "1_" + FileName
    job = doc.createElement('JOB')
    root.appendChild(job)
    leaf = doc.createElement('FileName')
    text = doc.createTextNode(FileName)
    leaf.appendChild(text)
    job.appendChild(leaf)
    leaf = doc.createElement('Quantity')
    text = doc.createTextNode(str('Maximum'))
    leaf.appendChild(text)
    job.appendChild(leaf)

else:
    iID = 1
    x = readNameCell('FirstID', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
    while x != "IndexError":
        FileName = xls_sheet.cell(x, 1).value.strip()
        FileName = os.path.basename(FileName)
        FileName = str(iID) + "_" + FileName
        job = doc.createElement('JOB')
        root.appendChild(job)
        leaf = doc.createElement('FileName')
        text = doc.createTextNode(FileName)
        leaf.appendChild(text)
        job.appendChild(leaf)
        leaf = doc.createElement('Quantity')
        text = doc.createTextNode(str(Quantity))
        leaf.appendChild(text)
        job.appendChild(leaf)
        x += 1
        iID += 1
        # проверяем следующую ячейк. если она пустая, и запрос выдаёт ошибку IndexError, выходим из цикла
        try:
            FileName = xls_sheet.cell(x, 1).value
        except IndexError:
            break

comment = doc.createComment("список файлов на поворот")
root.appendChild(comment)

if not allFileSame:
    # а тут делаем список файлов, с углом поворота файла.
    x = readNameCell('FirstID', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
    # files = doc.createElement('Files')
    # root.appendChild(files)
    while x != "IndexError":
        # FileName = xls_sheet.cell(x, 1).value
        FileName = str(xls_sheet.cell(x, 1).value)
        FileName = os.path.basename(FileName)
        leaf = doc.createElement('File')
        leaf.setAttribute("ID", str(int(xls_sheet.cell(x, 0).value)).strip())
        leaf.setAttribute("File", (xls_sheet.cell(x, 1).value.strip()))
        leaf.setAttribute("Angle", str(int(xls_sheet.cell(x, 8).value)))
        text = doc.createTextNode(FileName)
        leaf.appendChild(text)
        root.appendChild(leaf)
        x += 1
        # проверяем следующую ячейк. если она пустая, и запрос выдаёт ошибку IndexError, выходим из цикла
        try:
            FileName = xls_sheet.cell(x, 0).value
        except IndexError:
            break
else:
    # список на поворот файлов когда Quantity = Maximum
    x = readNameCell('FirstID', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
    FileName = str(xls_sheet.cell(x, 1).value)
    FileName = os.path.basename(FileName)
    leaf = doc.createElement('File')
    leaf.setAttribute("ID", str(int(xls_sheet.cell(x, 0).value)).strip())
    leaf.setAttribute("File", (FileName))
    leaf.setAttribute("Angle", str(int(xls_sheet.cell(x, 8).value)))
    text = doc.createTextNode(FileName)
    leaf.appendChild(text)
    root.appendChild(leaf)

# список красок и их очередность
y = readNameCell('relcolor1', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
x = readNameCell('relcolor1', xls_file)[1]  # возвращаем номер строки первой ячейки в таблице дизайнов
iID = 1
colors = doc.createElement('Colors')
root.appendChild(colors)
SumLenColorPole = 0
while x <= 9:
    cr = xls_sheet.cell(y, x).value
    textTest = str(xls_sheet.cell(y, x).value).replace(".0", "")
    if textTest != "":
        textTest = textTest
    else:
        textTest = "NaN"

    leaf = doc.createElement("Color")
    color = str(xls_sheet.cell(y, x).value).strip()
    if color != "":
        color = color.replace(".0", "")
    else:
        color = "NaN"
    leaf.setAttribute("Color", color)
    leaf.setAttribute("LPI", str(xls_sheet.cell(y + 1, x).value))
    leaf.setAttribute("Angle", str(xls_sheet.cell(y + 2, x).value))
    ThicknessPolimer_color = str(xls_sheet.cell(y + 4, x).value)
    if ThicknessPolimer_color == "":
        ThicknessPolimer_color = str(ThicknessPolymer)
    leaf.setAttribute("ThicknessPolimer_color", ThicknessPolimer_color)

    TypyPolimer_color = str(xls_sheet.cell(y + 5, x).value)
    if TypyPolimer_color == "":
        TypyPolimer_color = str(TypyPolimer)
    leaf.setAttribute("TypyPolimer_color", TypyPolimer_color)

    ColorSep_delete = str(xls_sheet.cell(y + 6, x).value)
    if ColorSep_delete == "да":
        ColorSep_delete = color
    leaf.setAttribute("ColorSep_delete", ColorSep_delete)

    ColorPole = str(xls_sheet.cell(y + 3, x).value).strip().replace(".0", "")

    if ColorPole == "":
        ColorPole = 'not'
    s = ColorPole
    ColorPole = s.split(',')

    # LenColorPole - количество полей. Условие: если цвет не заполнен, присваивается NaN, но ячейки с цифрами считаются.
    leaf.setAttribute("LenColorPole", str(len(ColorPole)))

    if len(ColorPole) == 1:
        leaf.setAttribute("ColorPole1", str(ColorPole[0]))
        leaf.setAttribute("ColorPole2", 'not')
        leaf.setAttribute("ColorPole3", 'not')
        # print(ColorPole[0])
        if ColorPole[0] == 'not':
            leaf.setAttribute("ColorPole1", 'not')
            leaf.setAttribute("LenColorPole", 'not')
        elif ColorPole[0] != 'not':
            SumLenColorPole += 1
    elif len(ColorPole) == 2:
        leaf.setAttribute("ColorPole1", str(ColorPole[0]))
        leaf.setAttribute("ColorPole2", str(ColorPole[1]))
        leaf.setAttribute("ColorPole3", 'not')
        SumLenColorPole += 2
    elif len(ColorPole) == 3:
        leaf.setAttribute("ColorPole1", str(ColorPole[0]))
        leaf.setAttribute("ColorPole2", str(ColorPole[1]))
        leaf.setAttribute("ColorPole3", str(ColorPole[2]))
        SumLenColorPole += 3
    # print('сумма' + str(SumLenColorPole))
    leaf.setAttribute("ID", str(iID))
    text = doc.createTextNode(textTest)
    leaf.appendChild(text)
    colors.appendChild(leaf)
    x += 1
    iID += 1

leaf = doc.createElement('SumLenColorPole')
text = doc.createTextNode(str(SumLenColorPole))
leaf.appendChild(text)
colors.appendChild(leaf)

# список красок с кращенными именами для рельс и их очередность
y = readNameCell('relcolor1', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
x = readNameCell('relcolor1', xls_file)[1]  # возвращаем номер строки первой ячейки в таблице дизайнов
iID = 0
colors = doc.createElement('ColorsRelsi')
root.appendChild(colors)
while x <= 9:
    cr = xls_sheet.cell(y, x).value
    textTest = str(xls_sheet.cell(y, x).value).replace(".0", "")
    if textTest != "":
        textTest = textTest.strip()

        arr1 = ['Pantone ', ' C', 'PANTONE ', 'U']
        for xx1 in arr1:
            textTest = textTest.replace(xx1, "")

    leaf = doc.createElement("Color")
    color = str(xls_sheet.cell(y, x).value)
    if color != "":
        color = color.replace(".0", "").strip()
    leaf.setAttribute("Color", color)
    text = doc.createTextNode(textTest)
    leaf.appendChild(text)
    colors.appendChild(leaf)
    x += 1
    iID += 1

# Определяем общее количество полей в контрольном поле для расчёта длины поля


# print(' ОБЩЕЕ КОЛИЧЕСТВО ПОЛЕЙ ' + str(SumLenColorPole))
PageX = int(3.5 * SumLenColorPole)
# print(' Ширина страницы ='+ '3.5 mm* '+ str(SumLenColorPole) +' полей = ' + str(PageX) + " mm")
PageY = 3.5
# print(' Высота страницы ' + str(PageY) + " mm")

c = Canvas(outputFolder + "\\" + "MarkColorPole.pdf", (PageX * mm, PageY * mm))

y = readNameCell('relcolor1', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
x = readNameCell('relcolor1', xls_file)[1]  # возвращаем номер столбца первой ячейки в таблице дизайнов
IDpole = 0
colorspole = doc.createElement('ColorsPole')
root.appendChild(colorspole)

while x <= 9:
    ColorPoleTint = str(xls_sheet.cell(y + 3, x).value).strip().replace(".0", "")
    Colorpole = str(xls_sheet.cell(y, x).value).strip().replace(".0", "")
    if (ColorPoleTint == "" and Colorpole == "") or (ColorPoleTint != "" and Colorpole == "") or (
            ColorPoleTint == "" and Colorpole != ""):
        x += 1
    else:
        # x += 1
        if ColorPoleTint != "" and Colorpole == "":
            Colorpole = "NaN"
        s = ColorPoleTint
        ColorPoleTint = s.split(',')
        xx = len(ColorPoleTint)
        for i in range(xx):
            ColorPoleTintID = ColorPoleTint[i].strip()
            ColorPoleTintID = float(ColorPoleTintID)
            ColorPoleTintID = str(ColorPoleTintID / 100).replace("1.0", "1").replace("0.", ".")
            leaf = doc.createElement("ColorPole")
            # print(' IDpole ' + str(IDpole) + ' tint ' + str(ColorPoleTintID) +' Цвет '+ Colorpole)

            c.setFillColor(
                PCMYKColorSep(randint(0, 100), randint(0, 100), randint(0, 100), randint(0, 100), spotName=Colorpole,
                              density=100 * float(ColorPoleTintID)))
            c.rect(3.5 * mm * IDpole, 0 * mm, 3.5 * mm, 3.5 * mm, fill=True, stroke=False)

            leaf.setAttribute("tint", ColorPoleTintID)
            leaf.setAttribute("IDpole", str(IDpole))

            Colorpole = str(xls_sheet.cell(y, x).value).strip().replace(".0", "")
            if Colorpole != "":
                leaf.setAttribute("Color", Colorpole)
                text = doc.createTextNode(textTest)
                leaf.appendChild(text)
                colorspole.appendChild(leaf)
            IDpole += 1
        x += 1
c.save()

# ----------------------------------------------------------------#
if os.path.exists(outputFolder) != True:
    os.mkdir(outputFolder, mode=0o777, dir_fd=None)

xml_out = outputFolder + "\\" + str(JobNamber).replace(".0", "") + ".xml"
xml_str = doc.toprettyxml(indent="    ")
try:
    with open(str(xml_out), "w", encoding='utf8') as f:
        f.write(xml_str)
        f.close()
except FileNotFoundError:
    logging.error(f'{xml_out} неправильный путь')
