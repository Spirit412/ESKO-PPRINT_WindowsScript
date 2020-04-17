import argparse
import pymysql
import xlrd
from xml.dom import minidom
import os

#reportlab - енерация PDF
from reportlab.lib.units import mm
from reportlab.lib.colors import PCMYKColor, PCMYKColorSep, CMYKColor, opaqueColor
from reportlab.pdfgen.canvas import Canvas
from random import randint


 # inputs, outputFolder, params
parser = argparse.ArgumentParser(description='inputs, outputFolder')
parser.add_argument('inputs', type=str, help='Input dir for xls file')
parser.add_argument('outputFolder', type=str, help='Output dir for xml file')

args = parser.parse_args()
# print(os.path.splitext(os.path.basename(args.inputs))[0])
# print(os.path.basename(args.inputs))
# print(args.outputFolder)


xls_file = args.inputs
try:
    xls_workbook = xlrd.open_workbook(str(xls_file))
except FileNotFoundError:
    print('скрипт не нашел Excel файла по адресу :' + args.inputs)
    exit()

xls_sheet = xls_workbook.sheet_by_index(0)
def rd(x, y=0):
    ''' Функция математического округления '''
    m = int('1' + '0' * y)  # multiplier - how many positions to the right
    q = x * m  # shift to the right by multiplier
    c = int(q)  # new number
    i = int((q - c) * 10)  # indicator number on the right
    if i >= 5:
        c += 1
    return c / m

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

table = str.maketrans("", "", "?()!@#$%^&*+|+\/:;[]{}<>") #список запрещённых символов в названии файла.
JobNamber = str(readNameCell('JobNamber', xls_file)[4]).replace(".0","")
JobNamber = JobNamber.translate(table) #удаляем запрещенные символы из номера заказа
print(JobNamber)
CustomerName = readNameCell('CustomerName', xls_file)[4]
CutTools = readNameCell('CutTools', xls_file)[4]
StampManufacturer = readNameCell('StampManufacturer', xls_file)[4]
Liniatura = readNameCell('Liniatura', xls_file)[4]
PrintMashine = readNameCell('PrintMashine', xls_file)[4]
ThicknessPolymer = readNameCell('ThicknessPolymer', xls_file)[4]
print(ThicknessPolymer)
TypyPolimer = readNameCell('TypyPolimer', xls_file)[4] #Тип полимера
DieShape = "file://server-esko/AE_BASE/CUT-TOOLS/" + CutTools + ".cf2"
DieShapeMFG = "file://server-esko/AE_BASE/CUT-TOOLS/" + CutTools + ".MFG"
Bleed = str(readNameCell('Bleed', xls_file)[4])
Bleed = Bleed.replace('.', ',')
PositionMark_6x4 = readNameCell('PositionMark_6x4', xls_file)[4] #Положение метки или её отсутствие
RotateCUT = readNameCell('RotateCUT', xls_file)[4] #Поворот штампа
UpOffset = readNameCell('UpOffset', xls_file)[4] #Кор-ка отступа до верхней рельсы
BotOffset = readNameCell('BotOffset', xls_file)[4] #Кор-ка отступа до нижней рельсы
psFile = readNameCell('psFile', xls_file)[4] #Какой PS файл на выходе - сепарированны/композитный или оба варианта
onebitTiff = readNameCell('onebitTiff', xls_file)[4] #PS риповать в 1битники или нет
addDGC = readNameCell('addDGC', xls_file)[4] #Применять DGC кривую при риповании
try:
    OutFileName = readNameCell('OutFileName', xls_file)[4] #Название файла
except AttributeError:
    OutFileName = os.path.splitext(os.path.basename(args.inputs))[0]


# Подключиться к базе данных.
connection = pymysql.connect(host='esko',
                             port=3360,
                             user='root',
                             password='',
                             db='pprint',
                             charset='utf8',
                            cursorclass=pymysql.cursors.DictCursor)
print("connect successful!!")

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
        Distorsia =  rd((Polimer/DPrint)*100, 4) #вычисляем % дисторсии до 4го знака после. В старой версии % брался из БД
        # для полимера 1.7
        Polimer17 = row['Hpolimer_17']
        Distorsia17 = rd((Polimer17/DPrint)*100, 4) #вычисляем % дисторсии до 4го знака после. В старой версии % брался из БД
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


if OutFileName == "":
    OutFileName = os.path.splitext(os.path.basename(args.inputs))[0]
else:
    # защита от использования запрещённых символов
    table = str.maketrans("", "", "?()!@#$%^&*+|+\/:;[]{}<>")  # список запрещённых символов в названии файла.
    OutFileName = OutFileName.translate(table) #удаляем запрещенные символы из номера заказа
print("Имя файла на выходе " + OutFileName)

leaf = doc.createElement('OutFileName')
text = doc.createTextNode(str(OutFileName))
leaf.appendChild(text)
root.appendChild(leaf)




# JOBы для спуска
iID = 1
x = readNameCell('FirstID', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
while x != "IndexError":
    cr = xls_sheet.cell(x, 1).value.strip()
    textTest = str(iID) + "_" + cr
    job = doc.createElement('JOB')
    root.appendChild(job)
    leaf = doc.createElement('FileName')
    text = doc.createTextNode(textTest)
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
        cr = xls_sheet.cell(x, 1).value
    except IndexError:
        break
comment = doc.createComment("список файлов на поворот")
root.appendChild(comment)

# а тут делаем список файлов, с углом поворота файла.
x = readNameCell('FirstID', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
# files = doc.createElement('Files')
# root.appendChild(files)
while x != "IndexError":
    cr = xls_sheet.cell(x, 1).value
    textTest = str(xls_sheet.cell(x, 1).value)
    leaf = doc.createElement('File')
    leaf.setAttribute("ID", str(int(xls_sheet.cell(x, 0).value)).strip())
    leaf.setAttribute("File", (xls_sheet.cell(x, 1).value.strip()))
    leaf.setAttribute("Angle", str(int(xls_sheet.cell(x, 8).value)))
    text = doc.createTextNode(textTest)
    leaf.appendChild(text)
    root.appendChild(leaf)
    x += 1
    # проверяем следующую ячейк. если она пустая, и запрос выдаёт ошибку IndexError, выходим из цикла
    try:
        cr = xls_sheet.cell(x, 0).value
    except IndexError:
        break

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
    leaf.setAttribute("ID",str(iID))
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


#Определяем общее количество полей в контрольном поле для расчёта длины поля



print(' ОБЩЕЕ КОЛИЧЕСТВО ПОЛЕЙ ' + str(SumLenColorPole))
PageX = int(3.5 * SumLenColorPole)
print(' Ширина страницы ='+ '3.5 mm* '+ str(SumLenColorPole) +' полей = ' + str(PageX) + " mm")
PageY = 3.5
print(' Высота страницы ' + str(PageY) + " mm")

c = Canvas(args.outputFolder + "\\" + "MarkColorPole.pdf", ( PageX*mm,PageY*mm))


y = readNameCell('relcolor1', xls_file)[0]  # возвращаем номер строки первой ячейки в таблице дизайнов
x = readNameCell('relcolor1', xls_file)[1]  # возвращаем номер столбца первой ячейки в таблице дизайнов
IDpole = 0
colorspole = doc.createElement('ColorsPole')
root.appendChild(colorspole)

while x <= 9:
    ColorPoleTint = str(xls_sheet.cell(y + 3, x).value).strip().replace(".0", "")
    Colorpole = str(xls_sheet.cell(y, x).value).strip().replace(".0", "")
    if (ColorPoleTint == "" and Colorpole =="") or (ColorPoleTint != "" and Colorpole =="") or (ColorPoleTint == "" and Colorpole !=""):
        x += 1
    else:
        # x += 1
        if ColorPoleTint != "" and Colorpole =="":
            Colorpole = "NaN"
        s = ColorPoleTint
        ColorPoleTint = s.split(',')
        xx = len(ColorPoleTint)
        for i in range(xx):
            ColorPoleTintID = ColorPoleTint[i].strip()
            ColorPoleTintID = float(ColorPoleTintID)
            ColorPoleTintID = str(ColorPoleTintID / 100).replace("1.0", "1").replace("0.", ".")
            leaf = doc.createElement("ColorPole")
            print(' IDpole ' + str(IDpole) + ' tint ' + str(ColorPoleTintID) +' Цвет '+ Colorpole)

            c.setFillColor( PCMYKColorSep(randint(0, 100),randint(0, 100),randint(0, 100),randint(0, 100), spotName=Colorpole, density=100*float(ColorPoleTintID)) )
            c.rect( 3.5 * mm * IDpole, 0 * mm, 3.5 * mm, 3.5 * mm, fill=True, stroke=False)

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

#----------------------------------------------------------------#
xml_out = args.outputFolder + "\\" + str(JobNamber).replace(".0", "") + ".xml"
xml_str = doc.toprettyxml(indent="  ")
with open(str(xml_out), "w", encoding='utf8') as f:
    f.write(xml_str)
    f.close()
