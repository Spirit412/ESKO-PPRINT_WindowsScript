from PyPDF2 import PdfFileReader, PdfFileWriter
from PyPDF2.pdf import PageObject

# 1 user space unit is 1/72 inch
# 1/72 inch ~ 0.352 millimeters


bigpage = '\\\Server-esko\\ae_base\\TEMP-Shuttle-IN\\fon.pdf'
maket = '\\\Server-esko\\ae_base\\TEMP-Shuttle-IN\\maket.pdf'
outfile = '\\\Server-esko\\ae_base\\TEMP-Shuttle-IN\\output.pdf'
tx = 100
ty = 100
bpw = 700
bph = None
tx *= 0.352
ty *= 0.352
bpw = 700
bph = 500
scale=1

inMaket = PdfFileReader(open(maket, 'rb'))
min_page = inMaket.getPage(0)
# Большая страница вместит 4 старницы (2x2)
big_page = PageObject.createBlankPage(None, bpw, bph)
# mergeScaledTranslatedPage(page2, scale, tx, ty, expand=False)
# https://pythonhosted.org/PyPDF2/PageObject.html
big_page.mergeScaledTranslatedPage(inMaket.getPage(0), scale, tx, ty)

writer = PdfFileWriter()
writer.addPage(big_page)

with open(outfile, 'wb') as f:
    writer.write(f)
