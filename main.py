import os
import re
from os import listdir
from os.path import isfile, join
from datetime import datetime

from PIL import Image
from pdf2image import convert_from_path
from pytesseract import pytesseract
import xlsxwriter


DATA_PATTERN    = r"\b\d{2}/\d{2}.*\d*,\d{2}\b"
EXPENSE_PATTERN = r"\b\d{2}/\d{2} (ACHAT|VIREMENT.*À|PRELEVEMENT).*\d*,\d{2}\b"


class BankStatementLine:
    def __init__(self, raw):
        self.date   = raw.split(" ")[0]
        self.amount = float(raw.split(" ")[-1].replace(',', '.'))
        self.type   = "expense" if re.fullmatch(EXPENSE_PATTERN, raw) else "income"

    def save_to_worksheet(self, row, worksheet):
        worksheet.write(row, 0, self.date)
        worksheet.write(row, 1, self.amount)
        worksheet.write(row, 2, self.type)

class BankStatementFile:
    def __init__(self, absolute_path, file_name):
        self.absolute_path = absolute_path
        self.file_name = file_name
        self.emission_date = datetime.strptime(
            file_name.split('_')[-1].split('.')[0],
            '%Y%d%m'
        ).date()

    def extract_data(self):
        doc = convert_from_path(self.absolute_path)
        workbook = xlsxwriter.Workbook('Expenses01.xlsx')
        worksheet = workbook.add_worksheet()

        for page_number, page_data in enumerate(doc):
            txt = pytesseract.image_to_string(page_data, lang='fra').encode('utf-8')
            decoded = txt.decode('utf-8')
            row = 0
            for line in decoded.split("\n"):
                if re.fullmatch(DATA_PATTERN, line):
                    extracted = BankStatementLine(line)
                    if extracted.type == "expense":
                        print(extracted.date + ": -" + str(extracted.amount))
                        extracted.save_to_worksheet(row, worksheet)
                    else:
                        print(extracted.date + ": " + str(extracted.amount))
                        extracted.save_to_worksheet(row, worksheet)
                    row += 1
        workbook.close()

def get_folder_content(root):
    files = [f for f in listdir(root) if isfile(join(root, f))]
    
    return [BankStatementFile(join(root, f), f) for f in files]


if __name__ == '__main__':
    content = get_folder_content('/path/to/file.pdf')
    for entry in content:
        entry.extract_data()
