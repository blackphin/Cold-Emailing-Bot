import xlrd
import smtplib, xlrd
from time import sleep

a = 0
b = 0
email_address = "sales@mscorpres.in"
email_password = "Sales@2020"
subject = "Let us help you grow your Business and turn it towards profitability"
# url = "https://mca.gov.in/content/mca/global/en/data-and-reports/company-llp-info/incorporated-closed-month.html"

message_file = open(r"D:/OneDrive/Repositories/Cold-Emailing-Bot/message.txt", "r")
message = message_file.read()

wb = xlrd.open_workbook(r"C:\Users\shiva\Downloads\Book1.xls")
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

for x in range(sheet.ncols):
    column_name = sheet.cell_value(0, x)
    if (column_name.lower()).find("email") != -1 and a == 0:
        email_column = x
        a = 1
    if (column_name.lower()).find("name") != -1 and b == 0:
        name_column = x
        b = 1

with smtplib.SMTP("smtp.gmail.com") as connection:
    connection.starttls()
    connection.login(user=email_address, password=email_password)

    for i in range(1, sheet.nrows):
        to_email_address = sheet.cell_value(i, email_column)
        company_name = (sheet.cell_value(i, name_column)).title()
        connection.sendmail(
            from_addr=email_address,
            to_addrs=to_email_address,
            msg="Subject:"
            + subject
            + "\n\n"
            + "Hi! hope you are having a wonderful day at "
            + company_name
            + ",\n"
            + message,
        )
        print(str(i) + ". Email sent to " + company_name + " successfully")
        sleep(1)
