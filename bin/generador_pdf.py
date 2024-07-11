import os
import io
import xlrd
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime, timedelta

current_folder = os.path.dirname (__file__)
parent_folder = os.path.dirname (current_folder)
files_folder = os.path.join (parent_folder, "files")
data = os.path.join (files_folder, f"Data.xlsx")
original_pdf = os.path.join (current_folder, f"solicitud.pdf")

def generatePDF(name, dni, motivo, fecha_solicitud, email, poblacion, ciudad, cp, telefono, situacion, cumple, basica, automatizacion, redes, riesgo, quirofano, lampara, generadora, observaciones, iite, centro, autoriza):
    packet = io.BytesIO()
    # Fonts with epecific path
    pdfmetrics.registerFont(TTFont('times','times.ttf'))
    pdfmetrics.registerFont(TTFont('timesbd', 'timesbd.ttf'))

    c = canvas.Canvas(packet, letter)

    #Página 1
    c.setFont('timesbd', 28)
    c.drawString(290-(len(name)/2)*16.66, 385, name)
    c.setFont('timesbd', 40)
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

    if(cumple):
        c.setFont('timesbd', 12)
        c.drawString(113, 390, "X")
    
    c.setFont('timesbd', 7)
    
    if(basica):
        c.drawString(552, 217, "X")
    if(automatizacion):
        c.drawString(552, 197, "X")
    if(redes):
        c.drawString(552, 179, "X")
    if(riesgo):
        c.drawString(552, 165, "X")
    if(quirofano):
        c.drawString(552, 151, "X")
    if(lampara):
        c.drawString(552, 131, "X")
    if(generadora):
        c.drawString(552, 107, "X")
    if(iite):
        c.drawString(552, 87, "X")
    
    c.setFont('timesbd', 6)

    if(situacion=="autonomo"):
        c.drawString(333, 395, "X")
    if(situacion=="ajena"):    
        c.drawString(416, 395, "X")

    c.showPage()

    #Página 3
    c.setFont('timesbd', 13)
    c.drawString(350, 505, dni)
    c.drawString(54, 399, observaciones[:82])
    c.drawString(54, 385, observaciones[82:165])
    c.drawString(54, 371, observaciones[165:252])
    
    if(autoriza=="no trabaja"):
        c.setFont('timesbd', 7)
        c.drawString(21, 331, "X")
        c.setFont('timesbd', 12)
        c.drawString(180, 301, centro)
    c.showPage()

    #Página 4
    c.setFont('timesbd', 10)
    c.drawString(90, 718, name)
    c.drawString(342, 520, fecha_solicitud)
    c.drawString(342, 210, fecha_solicitud)
    c.showPage()

    #Página 5
    c.setFont('timesbd', 10)
    c.drawString(72, 736, name)
    c.drawString(392, 736, dni)

    c.setFont('timesbd', 10)
    if basica or automatizacion or redes or riesgo or quirofano or lampara or generadora:
        c.drawString(58, 675, "X")
    if iite:
        c.drawString(58, 577, "X")
    c.setFont('timesbd', 12)
    if(basica):
        c.drawString(49, 613, "X")
    if(automatizacion):
        c.drawString(130, 613, "X")
    if(redes):
        c.drawString(212, 613, "X")
    if(riesgo):
        c.drawString(293, 613, "X")
    if(quirofano):
        c.drawString(374, 613, "X")
    if(lampara):
        c.drawString(455, 613, "X")
    if(generadora):
        c.drawString(537, 613, "X")

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

    page=existing_pdf.pages[3]
    output.add_page(page)

    #Tercera Página Editada
    page = existing_pdf.pages[4]
    page.merge_page(new_pdf.pages[2])
    output.add_page(page)

    for i in range (5, 20):
        page=existing_pdf.pages[i]
        output.add_page(page)
    
    #Cuarta Página Editada
    page = existing_pdf.pages[20]
    page.merge_page(new_pdf.pages[3])
    output.add_page(page)

    page=existing_pdf.pages[21]
    output.add_page(page)

    #Quinta Página Editada
    page = existing_pdf.pages[22]
    page.merge_page(new_pdf.pages[4])
    output.add_page(page)

    for i in range (23, 26):
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
    print(hoja.cell_value(i, 17))
    print(hoja.cell_value(i, 26))

    #fecha_segementada=hoja.cell_value(i, 1).split(" del ")
    #fecha_solicitud=fecha_segementada[0]+"/"+fecha_segementada[1]+"/"+fecha_segementada[2]
    try:
        fecha_solicitud = datetime(1899, 12, 30) + timedelta(days=hoja.cell_value(i, 1))
        fecha_solicitud = str(fecha_solicitud).split(" ")[0]
        fecha_solicitud = fecha_solicitud.split("-")[2] + "/" + fecha_solicitud.split("-")[1] + "/" + fecha_solicitud.split("-")[0].replace("20", "")
    except:
        fecha_solicitud = hoja.cell_value(i, 1)

    print(fecha_solicitud)
    centro=hoja.cell_value(i, 9)
    name=hoja.cell_value(i, 10) +" "+ hoja.cell_value(i, 11)
    dni=hoja.cell_value(i, 12)
    email=hoja.cell_value(i, 13)
    poblacion=hoja.cell_value(i, 15)
    ciudad=hoja.cell_value(i, 16)
    cp=hoja.cell_value(i, 17)
    telefono=hoja.cell_value(i, 18)
    observaciones=hoja.cell_value(i, 27)

    if(hoja.cell_value(i, 2)=="SI"):
        print("Inicial")
        motivo="inicial"
    if(hoja.cell_value(i, 3)=="SI"):
        print("Renovacion")
        motivo="renovacion"
    if(hoja.cell_value(i, 4)=="X"):
        print("Cumple")
        cumple = True
    else:
        print("No Cumple")
        cumple = False
    if(hoja.cell_value(i, 6)=="SI"):
        print("Autonomo")
        situacion="autonomo"
    if(hoja.cell_value(i, 7)=="SI"):
        print("Ajena")
        situacion="ajena"
    if(hoja.cell_value(i, 8)=="Si"):
        print("Autoriza")         
        autoriza="no trabaja"
    else:
        print("No autoriza")         
        autoriza="no" 
    if(hoja.cell_value(i, 19)=="SI"):
        print("Basica")
        basica=True
    else:
        print("No Basica")
        basica=False    
    if(hoja.cell_value(i, 20)=="SI"):
        print("Automatizacion")
        automatizacion=True
    else:
        print("No Automatizacion")
        automatizacion=False 
    if(hoja.cell_value(i, 21)=="SI"):
        print("Redes")
        redes=True
    else:
        print("No Redes")
        redes=False 
    if(hoja.cell_value(i, 22)=="SI"):
        print("Riesgo")
        riesgo=True
    else:
        print("No Riesgo")
        riesgo=False 
    if(hoja.cell_value(i, 23)=="SI"):
        print("Quirofano")
        quirofano=True
    else:
        print("No Quirofano")
        quirofano=False 
    if(hoja.cell_value(i, 24)=="SI"):
        print("Lampara")
        lampara=True
    else:
        print("No Lampara")
        lampara=False 
    if(hoja.cell_value(i, 25)=="SI"):
        print("Generadora")
        generadora=True
    else:
        generadora=False   
    if(hoja.cell_value(i, 26)=="SI"):
        print("IITE")
        iite=True
    else:
        print("No IITE")
        iite=False 
    print("_______________________________")
    generatePDF(name, dni, motivo, fecha_solicitud, email, poblacion, ciudad, cp, telefono, situacion, cumple, basica, automatizacion, redes, riesgo, quirofano, lampara, generadora, observaciones, iite, centro, autoriza)
print("Documentos generados correctamente")    
input()