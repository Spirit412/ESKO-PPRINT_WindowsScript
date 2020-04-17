import argparse
import xlrd
from xml.dom import minidom

import os
import pathlib
import shutil
import logging

# inputs, outputFolder, params
try:
    parser = argparse.ArgumentParser(description='inputs, outputFolder')
    parser.add_argument('inputs', type=str, help='Input dir for xls file')
    parser.add_argument('outputFolder', type=str, help='Output dir for xml file')
    args = parser.parse_args()
except:
    logging.error(u'Ошибка в аргументах к PY файлу')

logFile = str(args.outputFolder + "\\log.txt")
# сделаем лог файл в нужной кодировке.
try:
    with open(logFile, "w", encoding='utf8') as f:
        f.close()
except FileNotFoundError:
    logging.error(f'{logFile} неправильный путь')

logging.basicConfig(
    format=u'%(filename)s[LINE:%(lineno)d]# %(levelname)-8s [%(asctime)s]  %(message)s',
    level=logging.DEBUG,
    filename=logFile,
    filemode='w'
)

# # Сообщение отладочное
    # logging.debug(u'This is a debug message')
# # Сообщение информационное
    # logging.info(u'This is an info message')
# # Сообщение предупреждение
    # logging.warning(u'This is a warning')
# # Сообщение ошибки
    # logging.error(u'This is an error message')
# # Сообщение критическое
    # logging.critical(u'FATAL!!!')


logging.info(u'Имя входящего XLS файла: ' + os.path.basename(args.inputs))
logging.info(u'Адрес вывода: ' + os.path.normpath(args.outputFolder))
if os.path.exists(args.inputs) == False:
    logging.error(u'Не найден XLS файл указанный в первом аргументе PY скрипта')

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
    try:
        book = xlrd.open_workbook(xls)
        nameObj = book.name_and_scope_map.get((nameCell.lower(), -1))  # имя маленькими буквами
        r = [0, 1, 2, 3, 4]
        r[0] = nameObj.area2d()[1]
        r[1] = nameObj.area2d()[3]
        r[2] = nameObj.result.text
        r[3] = nameObj.name
        r[4] = nameObj.cell().value
    except FileNotFoundError:
        logging.error("Не найден XLS файл")
        print((logging.error("Не найден XLS файл")))
    return (r)


class readNameCell_XLS(object):
    """ункция чтения XLS файла"""
    def __init__(self, nameCell, url_xls_file):
        """Constructor"""
        self.nameCell = nameCell
        self.url_xls_file = url_xls_file
    def readNameCell(self, id):
        self.table = str.maketrans("", "", "()!@#$%^&*+|+\/:;[]{}<>")  # список запрещённых символов в названии файла.
        self.book = xlrd.open_workbook(self.url_xls_file)
        nameObj = self.book.name_and_scope_map.get((self.nameCell.lower(), -1))  # имя маленькими буквами
        r = [0, 1, 2, 3, 4]
        d = dict({0: nameObj.area2d()[1], 1: nameObj.area2d()[3], 2: nameObj.result.text, 3: nameObj.name, 4: nameObj.cell().value})
        r[0] = nameObj.area2d()[1]
        r[1] = nameObj.area2d()[3]
        r[2] = nameObj.result.text
        r[3] = nameObj.name
        r[4] = nameObj.cell().value
        return ((d[id].strip()).translate(table))



table = str.maketrans("", "", "()!@#$%^&*+|+\/:;[]{}<>")  # список запрещённых символов в названии файла.
JobNamber = str(readNameCell('Номер_заказа', xls_file)[4]).strip().replace(".0", "")
JobNamber = JobNamber.translate(table)  # удаляем запрещенные символы из номера заказа

CustomerName = readNameCell_XLS('Заказчик', args.inputs).readNameCell(4)
Substrate = readNameCell('Тип_материала', xls_file)[4]
PrintTech = readNameCell('Способ_печати', xls_file)[4]
ICCprof = readNameCell('ICC_профиль', xls_file)[4]
CutTools = readNameCell('Номер_штампа', xls_file)[4]
LabelSize = readNameCell('Размер_этикетки', xls_file)[4]
Winding = int(readNameCell('Вариант_намотки', xls_file)[4])
Designer = readNameCell('Дизайнер', xls_file)[4]

try:
    Dvtulka = int(readNameCell('Dvtulka', xls_file)[4])
    if Dvtulka == "":
        Dvtulka = " 76 / 40 "
except AttributeError:
    Dvtulka = " 76 / 40 "
except ValueError:
    Dvtulka = str(readNameCell('Dvtulka', xls_file)[4])

try:
    Vrulone = int(readNameCell('Vrulone', xls_file)[4])
    if Vrulone == "":
        Vrulone = "по заявке"
except AttributeError:
    Vrulone = "по заявке"
except ValueError:
    Vrulone = str(readNameCell('Vrulone', xls_file)[4])

ColorPrint = readNameCell('ColorPrint', xls_file)[4]
if ColorPrint == "Black":
    ColorPrintLogo = "\BW_logo_01.pdf"
else:
    ColorPrintLogo = "\CMYK_logo_01.pdf"

TextLayerScale = readNameCell('TextLayerScale', xls_file)[4]
TextLayerScale = str(TextLayerScale).replace(',', '.')

TextLayerPos = readNameCell('TextLayerPos', xls_file)[4]
OutTextLay = readNameCell('OutTextLay', xls_file)[4]
FileName1 = readNameCell('FileName1', xls_file)[4].strip()
if FileName1.lower().endswith(".pdf") == 0 and FileName1 != "":
    FileName1 = FileName1 + ".pdf"
FileName2 = readNameCell('FileName2', xls_file)[4].strip()
if FileName2.lower().endswith(".pdf") == 0 and FileName2 != "":
    FileName2 = FileName2 + ".pdf"

FileNameCount = 0
if FileName1 != "":
    FileNameCount = FileNameCount + 1
if FileName2 != "":
    FileNameCount = FileNameCount + 1

FileName1_podpis = readNameCell('FileName1_podpis', xls_file)[4]
if FileName1_podpis == "":
    FileName1_podpis = FileName1

FileName2_podpis = readNameCell('FileName2_podpis', xls_file)[4]
if FileName2_podpis == "":
    FileName2_podpis = FileName2

FileName1_rot = int(readNameCell('FileName1_rot', xls_file)[4])
FileName2_rot = int(readNameCell('FileName2_rot', xls_file)[4])

DopColor = readNameCell('DopColor', xls_file)[4]
if DopColor == "":
    DopColor == "."

PostPrintText = readNameCell('PostPrintText', xls_file)[4]
if PostPrintText == "":
    PostPrintText = " "

DopColor2 = readNameCell('DopColor2', xls_file)[4]
if DopColor2 == "":
    DopColor2 = "_"

PostPrintForm = readNameCell('ФормаДопОбр', xls_file)[4]
if PostPrintForm == "":
    PostPrintForm = "."

if PostPrintForm == "нов.":
    SummNewForm = 1
else:
    SummNewForm = 0

komment = readNameCell('комментарий', xls_file)[4]
if komment == "":
    komment = " "

# /////////////////// ФОРМИРУЕМ XML ///////////////////#
doc = minidom.Document()
root = doc.createElement('JOB')
doc.appendChild(root)

leaf = doc.createElement('JobNamber')
text = doc.createTextNode(str(JobNamber))
leaf.appendChild(text)
root.appendChild(leaf)

comment = doc.createComment("Комментарий")
root.appendChild(comment)

leaf = doc.createElement('CustomerName')
text = doc.createTextNode(str(CustomerName))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Substrate')
text = doc.createTextNode(str(Substrate))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('PrintTech')
text = doc.createTextNode(str(PrintTech))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('ICCprof')
text = doc.createTextNode(str(ICCprof))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('CutTools')
text = doc.createTextNode(str(CutTools))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('LabelSize')
text = doc.createTextNode(str(LabelSize))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Designer')
text = doc.createTextNode(str(Designer))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('ColorPrint')
text = doc.createTextNode(str(ColorPrint))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Dvtulka')
text = doc.createTextNode(str(Dvtulka))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Vrulone')
text = doc.createTextNode(str(Vrulone))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('ColorPrintLogo')
text = doc.createTextNode(str('\\\ESKO\\bg_data_marks_v010\\dat') + str(ColorPrintLogo))
leaf.appendChild(text)
root.appendChild(leaf)

comment = doc.createComment("Слой text")
root.appendChild(comment)

leaf = doc.createElement('TextLayerScale')
text = doc.createTextNode(str(TextLayerScale))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('TextLayerPos')
text = doc.createTextNode(str(TextLayerPos))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('OutTextLay')
text = doc.createTextNode(str(OutTextLay))
leaf.appendChild(text)
root.appendChild(leaf)

comment = doc.createComment("END Слой text")
root.appendChild(comment)

comment = doc.createComment("Размещение файлов на бланке")
root.appendChild(comment)
leaf = doc.createElement('FileNameCount')
text = doc.createTextNode(str(FileNameCount))
leaf.appendChild(text)
root.appendChild(leaf)





# делаем папку inPDF с проверкой
InPDFfolder = os.path.dirname(args.inputs) + "\\inPDF"
if os.path.exists(InPDFfolder) != True:
    os.mkdir(InPDFfolder, mode=0o777, dir_fd=None)
    logging.info(u'папки inPDF нет, создаём')
    print(os.path.exists(InPDFfolder))
else:
    logging.info(u'папка есть')
print(os.path.exists(FileName1))
# проверяем является ли значение абс путём к файлу. если да, возвращаем имя файла
if os.path.isabs(FileName1) == True:
    logging.info(u'Поле имени первого дизайна является абсолютным путём')
    # проверяем, есть ли файл или папка по адресу
    if os.path.exists(FileName1) == True:
        # если есть, копируем
        shutil.copy(FileName1, InPDFfolder)
        # в переменной заменяем абсолютный путь на имя файла
        FileName1 = os.path.basename('r' + FileName1)
    else:
        logging.error(u'файл отсутствует ' + FileName1)
    # FileName1 = os.path.basename('r' + FileName1)
else:
    logging.info(u'Поле имени первого дизайна является названием файла')
    # проверяем, есть ли этот файл в папке inPDF
    if os.path.exists(InPDFfolder + "\\" + FileName1) == True:
        logging.info(f'файл {FileName1} есть в папке inPDF')
    else:
        logging.error(f'файл {FileName1}  отсутствует в папке inPDF')

leaf = doc.createElement('FileName1')
text = doc.createTextNode(str(FileName1))
leaf.appendChild(text)
root.appendChild(leaf)

if os.path.isabs(FileName2) == True:
    logging.info(u'Поле имени первого дизайна является абсолютным путём')
    # проверяем, есть ли файл или папка по адресу
    if os.path.exists(FileName2) == True:
        # если есть, копируем
        shutil.copy(FileName2, InPDFfolder)
        # в переменной заменяем абсолютный путь на имя файла
        FileName2 = os.path.basename('r' + FileName2)
    else:
        logging.error(u'файл отсутствует ' + FileName2)
    # FileName2 = os.path.basename('r' + FileName2)
else:
    logging.info(u'Поле имени первого дизайна является названием файла')
    # проверяем, есть ли этот файл в папке inPDF
    if os.path.exists(InPDFfolder + "\\" + FileName2) == True:
        logging.info(f'файл {FileName2} есть в папке inPDF')
    else:
        logging.error(f'файл {FileName2}  отсутствует в папке inPDF')

leaf = doc.createElement('FileName2')
text = doc.createTextNode(str(FileName2))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('FileName1_podpis')
text = doc.createTextNode(str(FileName1_podpis))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('FileName2_podpis')
text = doc.createTextNode(str(FileName2_podpis))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('FileName1_rot')
text = doc.createTextNode(str(FileName1_rot))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('FileName2_rot')
text = doc.createTextNode(str(FileName2_rot))
leaf.appendChild(text)
root.appendChild(leaf)

comment = doc.createComment("END Размещение файлов на бланке")
root.appendChild(comment)

comment = doc.createComment("Дополнительная обработка")
root.appendChild(comment)

leaf = doc.createElement('DopColor')
text = doc.createTextNode(str(DopColor))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('DopColor2')
text = doc.createTextNode(str(DopColor2))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('PostPrintText')
text = doc.createTextNode(str(PostPrintText))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('PostPrintForm')
text = doc.createTextNode(str(PostPrintForm))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('SummNewForm')
text = doc.createTextNode(str(SummNewForm))
leaf.appendChild(text)
root.appendChild(leaf)

comment = doc.createComment("END Дополнительная обработка")
root.appendChild(comment)

comment = doc.createComment("Намотка")
root.appendChild(comment)

leaf = doc.createElement('Winding')
text = doc.createTextNode(str(Winding))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('WindingIMGURL')
text = doc.createTextNode("\\\ESKO\\bg_data_marks_v010\\dat\\Winding_" + str(Winding) + ".pdf")
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('WindingFile')
text = doc.createTextNode("Winding_" + str(Winding) + ".pdf")
leaf.appendChild(text)
root.appendChild(leaf)

comment = doc.createComment("END Намотка")
root.appendChild(comment)

leaf = doc.createElement('komment')
text = doc.createTextNode(str(komment))
leaf.appendChild(text)
root.appendChild(leaf)

xml_out = args.outputFolder + "\\" + str(JobNamber) + ".xml"
xml_str = doc.toprettyxml(indent="  ")
try:
    with open(str(xml_out), "w", encoding='utf8') as f:
        f.write(xml_str)
        f.close()
except FileNotFoundError:
    logging.error(f'{xml_out} неправильный путь')