import xlrd
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os

# rename the name of your excel file
NAME_FILE_EXCEL="lista_principale.xls"

# change the number of the column on your excel file where the name of partecipants is stocked
COLUMN_NUMBER=1

# name of the sample (withe certificate)
source_pdf_name="A. Cert-Problem-Solver.pdf"

# path of the excel file
# ATTENTION, THE FOLDER MUST EXIST
main_list_location = r"C:\Users\folder_with_ecxel_file"


def create_pdf(list, scr_name, dst_name):
    for name in list:
        packet = io.BytesIO()
        # create a new PDF with Reportlab
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFontSize(45)

        # function for placing the name in the center of the white rectangular
        if len(name) < 20:
            can.drawString(230, 240, name.center(20))
        elif 20 <= len(name) <= 23:
            can.drawString(180, 240, name.center(23))
        elif len(name) > 23:
            can.setFontSize(37)
            can.drawString(180, 240, name.center(25))
        can.save()

        # move to the beginning of the StringIO buffer
        packet.seek(0)
        new_pdf = PdfFileReader(packet)
        # read your existing PDF
        existing_pdf = PdfFileReader(open(scr_name, "rb"))
        output = PdfFileWriter()
        # add the "watermark" (which is the new pdf) on the existing page
        page = existing_pdf.getPage(0)
        page.mergePage(new_pdf.getPage(0))
        output.addPage(page)
        # finally, write "output" to a real file
        new_pdf_name = name + ".pdf"

        final_name = os.path.join(dst_name, new_pdf_name)

        outputStream = open(final_name, "wb")
        output.write(outputStream)
        outputStream.close()


'''
___________________________ parsing the excel file________________________________________________
'''
'''
INSERT IN "NAME_FILE_EXCEL" THE EXACT NAME OF YOUR EXCEL FILE
'''
primary_list = xlrd.open_workbook(NAME_FILE_EXCEL)
worksheet_primary_list = primary_list.sheet_by_index(0)
primary_name_list = []


for i in range(COLUMN_NUMBER, worksheet_primary_list.nrows):
    primary_name_list.append(str(worksheet_primary_list.cell(i, 1)).replace("text:","")[1:-1])


'''
___________________________ writing the PDF file ________________________________________________
'''

create_pdf(primary_name_list,source_pdf_name,main_list_location,)


