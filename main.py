import os
import re

from PIL import Image
from pdf2image import convert_from_path
from pytesseract import pytesseract
import xlsxwriter


def extract_text(_img):
    doc = convert_from_path(_img)
    path, fileName = os.path.split(_img)
    workbook = xlsxwriter.Workbook('Expenses01.xlsx')
    worksheet = workbook.add_worksheet()

    fileBaseName, fileExtension = os.path.splitext(fileName)
    pattern = r"\b\d{2}/\d{2}.*\d*,\d{2}\b"
    expense = r"\b\d{2}/\d{2} (ACHAT|VIREMENT.*À|PRELEVEMENT).*\d*,\d{2}\b"

    expenses = 0
    incomes = 0
    row = 0
    for page_number, page_data in enumerate(doc):
        txt = pytesseract.image_to_string(page_data, lang='fra').encode('utf-8')
        decoded = txt.decode('utf-8')
        for line in decoded.split("\n"):
            if re.fullmatch(pattern, line):
                date = line.split(" ")[0]
                amount = line.split(" ")[-1]
                if re.fullmatch(expense, line):
                    print(date + ": -" + amount)
                    expenses += float(amount.replace(',', '.'))
                    worksheet.write(row, 0, date)
                    worksheet.write(row, 1, float(amount.replace(',', '.')))
                    worksheet.write(row, 2, 'expense')
                else:
                    print(date + ": +" + amount)
                    incomes += float(amount.replace(',', '.'))
                    worksheet.write(row, 0, date)
                    worksheet.write(row, 1, float(amount.replace(',', '.')))
                    worksheet.write(row, 2, 'income')
                row += 1
    workbook.close()
    print("Incomes: " + str(incomes))
    print("Expenses: " + str(expenses))
    print("Balance: " + str(incomes - expenses))


if __name__ == '__main__':
    extract_text('/path/to/file.pdf')
