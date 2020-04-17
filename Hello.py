from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.colors import PCMYKColor, PCMYKColorSep, CMYKColor, opaqueColor
from reportlab.pdfgen.canvas import Canvas
from random import randint
from reportlab.graphics.shapes import Drawing, Circle
from reportlab.graphics.charts.textlabels import Label

c = Canvas ("\\\SERVER-ESKO\\TEMP-Shuttle-IN\\hello2.pdf", (500*mm,500*mm))
c.setFillColor(PCMYKColorSep( randint(0, 9), 100.0, 91.0, 0.0, spotName='PANTONE 485 C',density=100))
c.rect(0*mm,0*mm,3.5*mm,3.5*mm, fill=True, stroke=False)
c.setFillColor(PCMYKColorSep( 1, spotName='PANTONE 485 C',density=50))
c.rect(3.5*mm,0*mm,3.5*mm,3.5*mm, fill=True, stroke=False)

d = (Drawing(200*mm, 100*mm))

# mark the origin of the label
d.add(Circle(100*mm,90*mm, 5*mm, fillColor=colors.green))
lab = Label()
lab.setOrigin(100*mm,90*mm)
lab.boxAnchor = 'ne'
lab.angle = 45
lab.dx = 0
lab.dy = -20
lab.boxStrokeColor = colors.green
lab.setText('Some Multi-Line Label')
d.add(lab, '')

c.save()
# canvas.save()


