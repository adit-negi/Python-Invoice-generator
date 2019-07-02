from PyPDF2 import PdfFileWriter, PdfFileReader
import io
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
import xlrd

loc = ("payment.xlsx")
wb = xlrd.open_workbook(loc)
PlanPurchaseSheet = wb.sheet_by_index(0)

total_plans_purchased = PlanPurchaseSheet.nrows
print(total_plans_purchased)
flag = 1
pk = 1
charecter_count = 97


for i in range(2, total_plans_purchased):
    if PlanPurchaseSheet.cell_value(i, 10) == "Payment Successful":
        Plan = PlanPurchaseSheet.cell_value(i, 4)
        Plan_type = int(Plan)
        if Plan_type == 120:

            packet = io.BytesIO()
            # create a new PDF with Reportlab
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont('Helvetica', 10)
            name = PlanPurchaseSheet.cell_value(i, 1)
            email = PlanPurchaseSheet.cell_value(i, 2)
            phone = str(PlanPurchaseSheet.cell_value(i, 3))
            Date_Time = PlanPurchaseSheet.cell_value(i, 5)
            Date_Time_str = str(Date_Time)
            date = Date_Time_str[1:11]

            if pk < 999:
                if pk < 10:
                    invoive_number = date[8:10] + date[5:7] + date[2:4] + \
                        "-DEL-"+chr(charecter_count).upper()+"00" + str(pk)

                elif pk < 100:
                    invoive_number = date[8:10] + date[5:7] + date[2:4] + \
                        "-DEL-"+chr(charecter_count).upper()+"0" + str(pk)

                else:
                    invoive_number = date[7:9] + date[4:6] + date[1:3] + \
                        "-DEL-"+chr(charecter_count).upper() + str(pk)

            if pk > 999:
                pk = 1
                charecter_count = charecter_count + 1
                invoive_number = date[7:9] + date[4:6] + date[1:3] + \
                    "-DEL-"+chr(charecter_count).upper()+"00" + str(pk)
            print(invoive_number)
            can.drawString(80, 420, name)
            can.drawString(
                80, 404, "Delhi New Delhi*, Delhi (DL - 07), India")
            can.drawString(80, 388, email)
            can.drawString(80, 372, name)
            can.drawString(80, 360, phone)
            can.drawString(475, 580.9, date)
            can.drawString(458, 506, invoive_number)
            can.save()

            pk = pk+1
            count = pk
            flag = flag + 1
            a = str(flag)

            # move to the beginning of the StringIO buffer
            packet.seek(0)
            new_pdf = PdfFileReader(packet)
            # read your existing PDF
            existing_pdf = PdfFileReader(open("1RideTemplate.pdf", "rb"))
            output = PdfFileWriter()
            # add the "watermark" (which is the new pdf) on the existing page
            page = existing_pdf.getPage(0)
            page.mergePage(new_pdf.getPage(0))
            output.addPage(page)
            # finally, write "output" to a real file
            outputStream = open(
                "Invoicesss_" + invoive_number + ".pdf", "wb")
            output.write(outputStream)
            outputStream.close()
        if Plan_type == 550:

            packet = io.BytesIO()
            # create a new PDF with Reportlab
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont('Helvetica', 10)
            name = PlanPurchaseSheet.cell_value(i, 1)
            email = PlanPurchaseSheet.cell_value(i, 2)
            phone = str(PlanPurchaseSheet.cell_value(i, 3))
            Date_Time = PlanPurchaseSheet.cell_value(i, 5)
            Date_Time_str = str(Date_Time)
            date = Date_Time_str[1:11]

            if pk < 999:
                if pk < 10:
                    invoive_number = date[8:10] + date[5:7] + date[2:4] + \
                        "-DEL-"+chr(charecter_count).upper()+"00" + str(pk)
                elif pk < 100:
                    invoive_number = date[8:10] + date[5:7] + date[2:4] + \
                        "-DEL-"+chr(charecter_count).upper()+"0" + str(pk)
                else:
                    invoive_number = date[7:9] + date[4:6] + date[1:3] + \
                        "-DEL-"+chr(charecter_count).upper() + str(pk)

            if pk > 999:
                pk = 1
                charecter_count = charecter_count + 1
                invoive_number = date[7:9] + date[4:6] + date[1:3] + \
                    "-DEL-"+chr(charecter_count).upper()+"00" + str(pk)
            print(invoive_number)
            can.drawString(80, 420, name)
            can.drawString(
                80, 404, "Delhi New Delhi*, Delhi (DL - 07), India")
            can.drawString(80, 388, email)
            can.drawString(80, 372, name)
            can.drawString(80, 360, phone)
            can.drawString(475, 580.9, date)
            can.drawString(458, 506, invoive_number)
            can.save()
            pk = pk+1
            count = pk
            flag = flag + 1
            a = str(flag)

            # move to the beginning of the StringIO buffer
            packet.seek(0)
            new_pdf = PdfFileReader(packet)
            # read your existing PDF
            existing_pdf = PdfFileReader(open("5RideTemplate.pdf", "rb"))
            output = PdfFileWriter()
            # add the "watermark" (which is the new pdf) on the existing page
            page = existing_pdf.getPage(0)
            page.mergePage(new_pdf.getPage(0))
            output.addPage(page)
            # finally, write "output" to a real file
            outputStream = open(
                "Invoicesss_" + invoive_number + ".pdf", "wb")
            output.write(outputStream)
            outputStream.close()
        if Plan_type == 1000:

            packet = io.BytesIO()
            # create a new PDF with Reportlab
            can = canvas.Canvas(packet, pagesize=letter)
            can.setFont('Helvetica', 10)
            name = PlanPurchaseSheet.cell_value(i, 1)
            email = PlanPurchaseSheet.cell_value(i, 2)
            phone = str(PlanPurchaseSheet.cell_value(i, 3))
            Date_Time = PlanPurchaseSheet.cell_value(i, 5)
            Date_Time_str = str(Date_Time)
            date = Date_Time_str[1:11]

            if pk < 999:
                if pk < 10:
                    invoive_number = date[8:10] + date[5:7] + date[2:4] + \
                        "-DEL-"+chr(charecter_count).upper()+"00" + str(pk)
                elif pk < 100:
                    invoive_number = date[8:10] + date[5:7] + date[2:4] + \
                        "-DEL-"+chr(charecter_count).upper()+"0" + str(pk)
                else:
                    invoive_number = date[7:9] + date[4:6] + date[1:3] + \
                        "-DEL-"+chr(charecter_count).upper() + str(pk)

            if pk > 999:
                pk = 1
                charecter_count = charecter_count + 1
                invoive_number = date[7:9] + date[4:6] + date[1:3] + \
                    "-DEL-"+chr(charecter_count).upper()+"00" + str(pk)
            print(invoive_number)
            can.drawString(80, 420, name)
            can.drawString(
                80, 404, "Delhi New Delhi*, Delhi (DL - 07), India")
            can.drawString(80, 388, email)
            can.drawString(80, 372, name)
            can.drawString(80, 360, phone)
            can.drawString(475, 580.9, date)
            can.drawString(458, 506, invoive_number)
            can.save()
            pk = pk+1
            count = pk
            flag = flag + 1
            a = str(flag)

            # move to the beginning of the StringIO buffer
            packet.seek(0)
            new_pdf = PdfFileReader(packet)
            # read your existing PDF
            existing_pdf = PdfFileReader(open("10RideTemplate.pdf", "rb"))
            output = PdfFileWriter()
            # add the "watermark" (which is the new pdf) on the existing page
            page = existing_pdf.getPage(0)
            page.mergePage(new_pdf.getPage(0))
            output.addPage(page)
            # finally, write "output" to a real file
            outputStream = open(
                "Invoicesss_" + invoive_number + ".pdf", "wb")
            output.write(outputStream)
            outputStream.close()
print(count)
