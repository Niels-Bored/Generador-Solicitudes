import os
import io
import xlrd
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

current_folder = os.path.dirname (__file__)
parent_folder = os.path.dirname (current_folder)
files_folder = os.path.join (parent_folder, "files")
data = os.path.join (files_folder, f"Data.xlsx")
original_pdf = os.path.join (current_folder, f"solicitud.pdf")

def generatePDF(name, dni, motivo, fecha_solicitud, email, poblacion, ciudad, cp, telefono, situacion, prerrequisito, basica, automatizacion, redes, riesgo, quirofano, lampara, generadora):
    packet = io.BytesIO()
    # Fonts with epecific path
    pdfmetrics.registerFont(TTFont('times','times.ttf'))
    pdfmetrics.registerFont(TTFont('timesbd', 'timesbd.ttf'))

    c = canvas.Canvas(packet, landscape(letter))

    c.setFont('timesbd', 18)
    c.drawString(400-(len(name)/2)*7.5, 315, name)

    c.showPage()
    c.save()

    packet.seek(0)

    new_pdf = PdfFileReader(packet)
    
    existing_pdf = PdfFileReader(open(original_pdf, "rb"))
    output = PdfFileWriter()
    
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    page=existing_pdf.pages[1]
    output.add_page(page)
    
    new_pdf = os.path.join (files_folder, f"Solicitud {name}.pdf")
    output_stream = open(new_pdf, "wb")
    output.write(output_stream)
    output_stream.close()


  
wb = xlrd.open_workbook(data) 

hoja = wb.sheet_by_index(0) 
for i in range (2, hoja.nrows):
    print(hoja.cell_value(i, 1))
    print(hoja.cell_value(i, 9))
    print(hoja.cell_value(i, 10))
    print(hoja.cell_value(i, 11))
    print(hoja.cell_value(i, 12))
    print(hoja.cell_value(i, 13))
    print(hoja.cell_value(i, 14))
    print(hoja.cell_value(i, 15))
    print(hoja.cell_value(i, 16))

    fecha_solicitud=hoja.cell_value(i, 1)
    name=hoja.cell_value(i, 9)
    dni=hoja.cell_value(i, 10)
    email=hoja.cell_value(i, 11)
    poblacion=hoja.cell_value(i, 13)
    ciudad=hoja.cell_value(i, 14)
    cp=hoja.cell_value(i, 15)
    telefono=hoja.cell_value(i, 16)

    if(hoja.cell_value(i, 2)=="SI"):
        print("Inicial")
        motivo="inicial"
    if(hoja.cell_value(i, 3)=="SI"):
        print("Renovacion")
        motivo="renovacion"
    if(hoja.cell_value(i, 4)=="SI"):
        print("Experiencia")
        prerrequisito="experiencia"
    if(hoja.cell_value(i, 5)=="SI"):
        print("Formacion")         
        prerrequisito="formacion"
    if(hoja.cell_value(i, 6)=="SI"):
        print("Autonomo")
        situacion="autonomo"
    if(hoja.cell_value(i, 7)=="SI"):
        print("Ajena")
        situacion="ajena"
    if(hoja.cell_value(i, 8)=="SI"):
        print("No trabaja")         
        situacion="no trabaja" 
    if(hoja.cell_value(i, 17)=="SI"):
        print("Basica")
        basica=True
    else:
        print("No Basica")
        basica=False    
    if(hoja.cell_value(i, 18)=="SI"):
        print("Automatizacion")
        automatizacion=True
    else:
        print("No Automatizacion")
        automatizacion=False 
    if(hoja.cell_value(i, 19)=="SI"):
        print("Redes")
        redes=True
    else:
        print("No Redes")
        redes=False 
    if(hoja.cell_value(i, 20)=="SI"):
        print("Riesgo")
        riesgo=True
    else:
        print("No Riesgo")
        riesgo=False 
    if(hoja.cell_value(i, 21)=="SI"):
        print("Quirofano")
        quirofano=True
    else:
        print("No Quirofano")
        quirofano=False 
    if(hoja.cell_value(i, 22)=="SI"):
        print("Lampara")
        lampara=True
    else:
        print("No Lampara")
        lampara=False 
    if(hoja.cell_value(i, 23)=="SI"):
        print("Generadora")
        generadora=True
    else:
        print("No Generadora")
        generadora=False    
    print("_______________________________")
    generatePDF(name, dni, motivo, fecha_solicitud, email, poblacion, ciudad, cp, telefono, situacion, prerrequisito, basica, automatizacion, redes, riesgo, quirofano, lampara, generadora)
print("Documentos generados correctamente")    
input()