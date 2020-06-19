import argparse
import fnmatch
import os
import xml.etree.ElementTree as ET

from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.pdf import PageObject
import PyPDF2


def create_pdf_marks_layers(url_xml):
    # Поиск в XML с помощью XPath
    tree = ET.parse(url_xml)
    root = tree.getroot()
    cut_name = root.find('.//JOBNAME').text
    size_cut = {'width': float(root.find('.//WIDTH').attrib['number']),
                'height': float(root.find('.//HEIGHT').attrib['number'])}
    number_up = root.find('.//NUMBERUP').text
    size_up = {'width': float(root.find('.//HSIZE').text), 'height': float(root.find('.//VSIZE').text)}

    print('название штампа cut_name: ', cut_name[:-4])
    print('количество этикеток number_up: ', number_up)
    print('размер единички size_up: ', size_up)
    print('размер штампа с блидами size_cut: ', size_cut)

    def label_coord_list():
        items = list(root.findall('.//POSITION'))
        item_list = list(map(lambda x: [float(x.find('H').text), float(x.find('V').text)], items))
        return item_list

    print('label_coord_list', label_coord_list())

    def colom_coord_list():
        items = list(root.findall('.//POSITION'))
        item_list = list(set(map(lambda x: float(x.find('H').text), items)))
        item_list = sorted(item_list)
        return item_list

    def row_coord_list():
        items = list(root.findall('.//POSITION'))
        item_list = list(set(map(lambda x: float(x.find('V').text), items)))
        item_list = sorted(item_list)
        return item_list

    print('================колонки===================')
    print('colom_coord_list', colom_coord_list())
    print('=================строки====================')
    print('row_coord_list: ', row_coord_list())

    def offset_row():
        if len(row_coord_list()) > 1:
            offset = (row_coord_list()[1] - row_coord_list()[0] - size_up['height'])
        else:
            offset = 0
        return offset

    print('offset_row:', offset_row())

    def offset_col():
        if len(colom_coord_list()) > 1:
            offset = (colom_coord_list()[1] - colom_coord_list()[0]) - size_up['width']
        else:
            offset = (size_cut['width']-size_up['width']) / 2
        return offset

    print('offset_col:', offset_col())

    # for i in range(1, int(number_up), 4):
    #     print(label_coord_list().__getitem__(i))
    # print(label_coord_list())

    # for element in tree.findall(".//JOBNAME"):
    #     print(element.tag)
    # tree.
    # print(tree.getroot().tag)

    # 1 дюйм = 72, 0000000000005
    # пункт НИС / PostScript 1 пункт НИС / PostScript = 0.3527777777778 миллиметр
    ######################################### START генерируем PDF ###########################################

    bigpage = '\\\Server-esko\\ae_base\\TEMP-Shuttle-IN\\' + cut_name[:-4] + '_Mark_TEST.pdf'
    mark_url = r'\\esko\bg_data_marks_v010\dat\krest_for_bottom layer_layout.pdf'
    mm_to_pt = 25.4 / 72
    """
    Константа для преобразования мм в пункты PS
    """
    bpw = size_cut['width'] / mm_to_pt
    bph = size_cut['height'] / mm_to_pt
    scale_mark = 1

    mark_read = PdfFileReader(open(mark_url, 'rb'))
    mark = mark_read.getPage(0)
    mark_size_pt = {'width': float(mark.mediaBox.getWidth()), 'height': float(mark.mediaBox.getHeight())}

    print('mark', mark['/MediaBox'])
    print('mark_size', mark_size_pt)

    ## PyPDF2 работает с размерами в пунктах 1/72 ps дюйма в мм = 25,4/72
    big_page = PageObject.createBlankPage(None, bpw, bph)

    if len(colom_coord_list()) == 1 and len(row_coord_list()) > 1:
        tx = ((colom_coord_list()[0] + size_up['width']) / mm_to_pt) - mark_size_pt['width'] / 2
        tx += (offset_col() / 2) / mm_to_pt
        ty = (row_coord_list()[0] + size_up['height']) / mm_to_pt + mark_size_pt['height'] / 2
        ty += (offset_row()/2) / mm_to_pt
        big_page.mergeScaledTranslatedPage(mark, scale_mark, tx=tx, ty=bph - ty)

        tx = (colom_coord_list()[0] / mm_to_pt) - (mark_size_pt['width'] / 2)
        tx += (offset_col() / 2) / mm_to_pt
        ty = (row_coord_list()[0] + size_up['height']) / mm_to_pt + mark_size_pt['height'] / 2
        ty += (offset_row()/2) / mm_to_pt
        big_page.mergeScaledTranslatedPage(mark, scale_mark, tx=tx, ty=bph - ty)

    writer = PdfFileWriter()
    writer.addPage(big_page)
    with open(bigpage, 'wb') as f:
        writer.write(f)
        f.close()

    ######################################### END генерируем PDF ###########################################


if __name__ == "__main__":
    url_path = '\\\\Server-esko\\ae_base\CUT-TOOLS'
    listOfFiles = os.listdir(url_path)
    pattern = "*.xml"
    xml = []
    for entry in listOfFiles:
        if fnmatch.fnmatch(entry, pattern):
            if entry:
                xml.append(entry)
    print(xml)
    for file in xml:
        print('==  ' * 40)
        a = r'\\Server-esko\ae_base\CUT-TOOLS'
        create_pdf_marks_layers(url_xml='{}\{}'.format(a, file))
