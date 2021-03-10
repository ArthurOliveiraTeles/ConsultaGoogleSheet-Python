"""@author: ArthurOliveiraTeles
Last modified date: 10/03/2021


For the purpose of this "Tunts" challenge, I will leave my spreadsheet complete with the values ​​assigned by my Python application.

This application of mine used Google's APIs and Google Sheets, through Google Cloud."""

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import math
import logging

# Function for tracking with log, tracks requests made and returns a status code. If it is 200, everything is OK.
# Creates a file that I called "mytest.log"
logging.basicConfig(filename="mytest.log", level=logging.DEBUG,
                    format="%(asctime)s %(levelname)s %(funcName)s => %(message)s")

# Here he takes the scope of the API and the authorization for the google sheet and drive
scope =["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

# Credential acquired by the JSON file
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

try:
    client = gspread.authorize(creds)
    
    # Accessing my sheet file    
    sheet = client.open("Cópia de Engenharia de Software - Desafio [Arthur Oliveira Teles]").sheet1

    # Storing all values ​​in the data variable
    data = sheet.get_all_records()
    logging.debug(f"Authorize credential: {client}", f"File access: {sheet}", "Get all records:OK")

except:
    logging.exception('Unknow exception')


# App code

count = 4 # Accountant must start at 4 because in the document, the 4th line starts students
while count <= (len(data)+1):
    # Accessing all rows in the spreadsheet
    row = sheet.row_values(count)
    print("-----All student grades-----")
    print(row[3], row[4], row[5])

    # Converting the values ​​of columns 3, 4 and 5 (notes) to float.
    n1 = float(row[3])
    n2 = float(row[4])
    n3 = float(row[5])

    # Calculation of average and rounding
    media = (n1 + n2 + n3)/3
    math.floor(media)

    # Calculation of the percentage of absences
    max_faltas = (25*60)/100 # pode apenas 15 faltas (translate to english later)

    if float(row[2]) > max_faltas:
        # Updating the cell of each student and their columns "Situation" and "Final Grade"
        # (row value, collumn, text)
        sheet.update_cell(count, 7, "Reprovado por falta")
        sheet.update_cell(count, 8, "0")
    else:
        if media > 70:
            sheet.update_cell(count, 7, "Aprovado")
            sheet.update_cell(count, 8, "0")

        if media < 50:
            sheet.update_cell(count, 7, "Reprovado por nota")
            sheet.update_cell(count, 8, "0")

        if 50 <= media < 70:
            sheet.update_cell(count, 7, "Exame Final")

            # Calculation of grade for final approval
            mediaFinal = (50*2) - media
            sheet.update_cell(count, 8, math.floor(mediaFinal))


    # It is important to add +1 to the counter at the end of the code, so that it skips the line.
    count += 1
