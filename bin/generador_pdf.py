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

def generatePDF(name, dni, motivo, fecha_solicitud, email, poblacion, ciudad, cp, telefono, situacion, prerrequisito, basica, automatizacion, redes, riesgo, quirofano, lampara, generadora, observaciones):
    packet = io.BytesIO()
    # Fonts with epecific path
    pdfmetrics.registerFont(TTFont('times','times.ttf'))
    pdfmetrics.registerFont(TTFont('timesbd', 'timesbd.ttf'))

    c = canvas.Canvas(packet, letter)

    #Página 1
    c.setFont('timesbd', 40)
    c.drawString(280-(len(name)/2)*16.66, 385, name)
    c.drawString(280-(len(dni)/2)*16.66, 335, dni)
    c.showPage()

    #Página 2
    c.setFont('timesbd', 14)
    c.drawString(190, 610, fecha_solicitud)
    c.drawString(210, 556, name)
    c.drawString(138, 528, dni)
    c.drawString(142, 502, email)
    c.drawString(160, 474, poblacion)
    c.drawString(374, 474, ciudad)
    c.drawString(180, 448, str(int(cp)))
    c.drawString(160, 422, str(int(telefono)))

    if(motivo=="inicial"):
        c.drawString(131, 641, "X")
    if(motivo=="renovacion"):
        c.drawString(293, 641, "X")

    if(basica):
        c.drawString(530, 235, "X")
    if(automatizacion):
        c.drawString(530, 213, "X")
    if(redes):
        c.drawString(530, 191, "X")
    if(riesgo):
        c.drawString(530, 169, "X")
    if(quirofano):
        c.drawString(530, 148, "X")
    if(lampara):
        c.drawString(530, 126, "X")
    if(generadora):
        c.drawString(530, 103, "X")
    
    c.setFont('timesbd', 6)

    if(situacion=="autonomo"):
        c.drawString(255, 395, "X")
    elif(situacion=="ajena"):    
        c.drawString(350, 395, "X")
    elif(situacion=="no trabaja" ):    
        c.drawString(441, 395, "X")

    if(prerrequisito=="experiencia"):
        c.drawString(113, 340, "X")
    elif(prerrequisito=="formacion"):
        c.drawString(113, 313, "X")

    c.showPage()

    #Página 3
    c.setFont('timesbd', 14)
    c.drawString(355, 293, dni)
    c.drawString(208, 107, observaciones[:50])
    c.drawString(53, 94, observaciones[51:128])
    c.drawString(53, 80, observaciones[128:207])
    c.showPage()

    #Página 4
    c.setFont('timesbd', 14)
    c.drawString(90, 718, name)
    c.drawString(342, 520, fecha_solicitud)
    c.drawString(342, 210, fecha_solicitud)
    c.showPage()

    #Página 5
    c.setFont('timesbd', 14)
    c.drawString(72, 736, name)
    c.drawString(392, 736, dni)

    
    c.showPage()
    c.save()

    packet.seek(0)

    new_pdf = PdfFileReader(packet)
    
    existing_pdf = PdfFileReader(open(original_pdf, "rb"))
    output = PdfFileWriter()
    
    #Primera Página Editada 1, 2, 18, 20
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)

    #Segunda Página Editada
    page = existing_pdf.pages[1]
    page.merge_page(new_pdf.pages[1])
    output.add_page(page)

    page=existing_pdf.pages[2]
    output.add_page(page)

    #Tercera Página Editada
    page = existing_pdf.pages[3]
    page.merge_page(new_pdf.pages[2])
    output.add_page(page)

    for i in range (4, 17):
        page=existing_pdf.pages[i]
        output.add_page(page)
    
    #Cuarta Página Editada
    page = existing_pdf.pages[17]
    page.merge_page(new_pdf.pages[3])
    output.add_page(page)

    page=existing_pdf.pages[18]
    output.add_page(page)

    #Quinta Página Editada
    page = existing_pdf.pages[19]
    page.merge_page(new_pdf.pages[4])
    output.add_page(page)

    for i in range (20, 22):
        page=existing_pdf.pages[i]
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
    print(hoja.cell_value(i, 24))

    fecha_segementada=hoja.cell_value(i, 1).split(" del ")
    fecha_solicitud=fecha_segementada[0]+"/"+fecha_segementada[1]+"/"+fecha_segementada[2]
    print(fecha_solicitud)
    name=hoja.cell_value(i, 9)
    dni=hoja.cell_value(i, 10)
    email=hoja.cell_value(i, 11)
    poblacion=hoja.cell_value(i, 13)
    ciudad=hoja.cell_value(i, 14)
    cp=hoja.cell_value(i, 15)
    telefono=hoja.cell_value(i, 16)
    observaciones=hoja.cell_value(i, 24)

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
    generatePDF(name, dni, motivo, fecha_solicitud, email, poblacion, ciudad, cp, telefono, situacion, prerrequisito, basica, automatizacion, redes, riesgo, quirofano, lampara, generadora, observaciones)
print("Documentos generados correctamente")    
input()