import argparse
import xlrd
from xml.dom import minidom

 # inputs, outputFolder, params
parser = argparse.ArgumentParser(description='inputs, outputFolder')
parser.add_argument('inputs', type=str, help='Input dir for xls file')
parser.add_argument('outputFolder', type=str, help='Output dir for xml file')

args = parser.parse_args()
print(args.inputs)
print(args.outputFolder)

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

table = str.maketrans("", "", "()!@#$%^&*+|+\/:;[]{}<>") #список запрещённых символов в названии файла.
JobNamber = str(readNameCell('Номер_заказа', xls_file)[4]).strip().replace(".0", "")
JobNamber = JobNamber.translate(table) #удаляем запрещенные символы из номера заказа
print(JobNamber)

CustomerName = readNameCell('Заказчик', xls_file)[4].strip()
Substrate = readNameCell('Тип_материала', xls_file)[4]
PrintTech = readNameCell('Способ_печати', xls_file)[4]
ICCprof = readNameCell('ICC_профиль', xls_file)[4]
CutTools = readNameCell('Номер_штампа', xls_file)[4].strip()
LabelSize = readNameCell('Размер_этикетки', xls_file)[4]
LabelPart = readNameCell('часть_этикетки', xls_file)[4]
Winding = int(readNameCell('Вариант_намотки', xls_file)[4])
Designer = readNameCell('Дизайнер', xls_file)[4]
komment = readNameCell('комментарий', xls_file)[4]
if komment == "":
    komment = "нет"

try:
    Dvtulka = int(readNameCell('Dvtulka', xls_file)[4])
    if Dvtulka == "":
        Dvtulka = " 76 / 40 "
except AttributeError:
    Dvtulka = " 76 / 40 "
try:
    Vrulone = readNameCell('Vrulone', xls_file)[4]
    if Vrulone == "":
        Vrulone = "по заявке"
except AttributeError:
    Vrulone = "по заявке"

TextLayerScale = readNameCell('TextLayerScale', xls_file)[4]
TextLayerScale = str(TextLayerScale).replace(',', '.')

TextLayerPos = readNameCell('TextLayerPos', xls_file)[4]
OutTextLay = readNameCell('OutTextLay', xls_file)[4]

kongrev = readNameCell('конгрев', xls_file)[4]
tisnenie = readNameCell('тиснение', xls_file)[4]
vibor_LAK = readNameCell('выб_лак', xls_file)[4]
sploshnoy_LAK = readNameCell('сплош_лак', xls_file)[4]
belila = readNameCell('белила', xls_file)[4]

try:
    PostPrint1 = readNameCell('PostPrint1', xls_file)[4]
    if PostPrint1 == "":
        PostPrint1 = "  "
    elif PostPrint1 != "":
        PostPrint1 = "1 - " + PostPrint1
except AttributeError:
    PostPrint1 = readNameCell('конгрев', xls_file)[4]
    if PostPrint1 == "":
        PostPrint1 = "  "
    elif PostPrint1 != "":
        PostPrint1 = "конгрев: " + PostPrint1

try:
    PostPrint2 = readNameCell('PostPrint2', xls_file)[4]
    if PostPrint2 == "":
        PostPrint2 = "  "
    elif PostPrint2 != "":
        PostPrint2 = "2 - " + PostPrint2
except AttributeError:
    PostPrint2 = readNameCell('тиснение', xls_file)[4]
    if PostPrint2 == "":
        PostPrint2 = "  "
    elif PostPrint2 != "":
        PostPrint2 = "тиснение: " + PostPrint2

try:
    PostPrint3 = readNameCell('PostPrint3', xls_file)[4]
    if PostPrint3 == "":
        PostPrint3 = "  "
    elif PostPrint3 != "":
        PostPrint3 = "3 - " + PostPrint3
except AttributeError:
    PostPrint3 = readNameCell('выб_лак', xls_file)[4]
    if PostPrint3 == "":
        PostPrint3 = "  "
    elif PostPrint3 != "":
        PostPrint3 = "выб. лак - " + PostPrint3

try:
    PostPrint4 = readNameCell('PostPrint4', xls_file)[4]
    if PostPrint4 == "":
        PostPrint4 = "  "
    elif PostPrint4 != "":
        PostPrint4 = "4 - " + PostPrint4
except AttributeError:
    PostPrint4 = readNameCell('сплош_лак', xls_file)[4]
    if PostPrint4 == "":
        PostPrint4 = "  "
    elif PostPrint4 != "":
        PostPrint4 = "сплошной лак - " + PostPrint4

try:
    PostPrint5 = readNameCell('PostPrint5', xls_file)[4]
    if PostPrint5 == "":
        PostPrint5 = "  "
    elif PostPrint5 != "":
        PostPrint5 = "5 - " + PostPrint5
except AttributeError:
    PostPrint5 = readNameCell('белила', xls_file)[4]
    if PostPrint5 == "":
        PostPrint5 = "  "
    elif PostPrint5 != "":
        PostPrint5 = "белила - " + PostPrint5

# /////////////////// ФОРМИРУЕМ XML ///////////////////#
doc = minidom.Document()
root = doc.createElement('JOB')
doc.appendChild(root)

leaf = doc.createElement('JobNamber')
text = doc.createTextNode(str(JobNamber).replace(".0", ""))
leaf.appendChild(text)
root.appendChild(leaf)

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

leaf = doc.createElement('LabelPart')
text = doc.createTextNode(str(LabelPart))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Designer')
text = doc.createTextNode(str(Designer))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('komment')
text = doc.createTextNode(str(komment))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Dvtulka')
text = doc.createTextNode(str(Dvtulka))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('Vrulone')
text = doc.createTextNode(str(Vrulone).replace(".0", ""))
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
text = doc.createTextNode("Winding_" + str(Winding) + ".pdf" )
leaf.appendChild(text)
root.appendChild(leaf)

comment = doc.createComment("END Намотка")
root.appendChild(comment)

comment = doc.createComment("Отделка")
root.appendChild(comment)

leaf = doc.createElement('kongrev')
text = doc.createTextNode(str(kongrev))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('tisnenie')
text = doc.createTextNode(str(tisnenie))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('vibor_LAK')
text = doc.createTextNode(str(vibor_LAK))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('sploshnoy_LAK')
text = doc.createTextNode(str(sploshnoy_LAK))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('belila')
text = doc.createTextNode(str(belila))
leaf.appendChild(text)
root.appendChild(leaf)

comment = doc.createComment("END отделка")
root.appendChild(comment)

comment = doc.createComment("Непечатные краски")
root.appendChild(comment)

leaf = doc.createElement('PostPrint1')
text = doc.createTextNode(str(PostPrint1))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('PostPrint2')
text = doc.createTextNode(str(PostPrint2))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('PostPrint3')
text = doc.createTextNode(str(PostPrint3))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('PostPrint4')
text = doc.createTextNode(str(PostPrint4))
leaf.appendChild(text)
root.appendChild(leaf)

leaf = doc.createElement('PostPrint5')
text = doc.createTextNode(str(PostPrint5))
leaf.appendChild(text)
root.appendChild(leaf)

comment = doc.createComment("END Непечатные краски")
root.appendChild(comment)
# Краски макета
job = doc.createElement('Inks')
root.appendChild(job)
yy = ["нов", "нов.", "новая", "new"]
SummNewForm = 0
iID = 1
x = readNameCell('FirstColor', xls_file)[0] # возвращаем номер строки первой ячейки в таблице дизайнов
while x != "IndexError":
    ColorName = ""
    cr = xls_sheet.cell(x, 1).value.strip()
    textTest = cr

    leaf = doc.createElement('Ink')
    leaf.setAttribute("ID" , str(iID))

    ColorName = str(xls_sheet.cell(x, 1).value.strip())
    if ColorName == "":
        ColorName = "NaN"
    leaf.setAttribute("ColorName", ColorName)

    leaf.setAttribute("Frequency", str(xls_sheet.cell(x, 2).value).replace(".0", "")) # удаляем из строки ".0"
    leaf.setAttribute("Angle", str(xls_sheet.cell(x, 3).value).replace(".0", "")) # удаляем из строки ".0"
    leaf.setAttribute("InkParam", str(xls_sheet.cell(x, 4).value).strip(' ')) #удаляем пробелы вначале/конце строки
    if str(xls_sheet.cell(x, 4).value) in yy: # идет сравнение со значениями элементов массива yy
        SummNewForm += 1
    text = doc.createTextNode(textTest)
    leaf.appendChild(text)
    job.appendChild(leaf)
    x += 1
    iID += 1
    # проверяем следующую ячейк. если она пустая, и запрос выдаёт ошибку IndexError, выходим из цикла
    try:
        cr = int(xls_sheet.cell(x, 0).value) == ''
    except ValueError:
        break

leaf = doc.createElement('SummNewForm')
text = doc.createTextNode(str(SummNewForm))
leaf.appendChild(text)
root.appendChild(leaf)



xml_out = args.outputFolder + "\\" + str(JobNamber).replace(".0", "") + ".xml"
# xml_out = outputFolder + "\\" + str(JobNamber) + ".xml"
xml_str = doc.toprettyxml(indent="  ")
with open(str(xml_out), "w", encoding='utf8') as f:
    f.write(xml_str)
    f.close()