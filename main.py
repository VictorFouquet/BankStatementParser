import os
import re
from os import listdir
from os.path import isfile, join
from datetime import datetime

from PIL import Image
from pdf2image import convert_from_path
from pytesseract import pytesseract
import xlsxwriter

NUMBER_PATTERN  = r"((?:\d{1,3}(?: \d{3})*|\d+),\d{2})"
DATA_PATTERN    = r"\b\d{2}/\d{2}.*((?:\d{1,3}(?: \d{3})*|\d+),\d{2})\b"
EXPENSE_PATTERN = r"\b\d{2}/\d{2} (ACHAT|VIREMENT.*À|PRELEVEMENT|CARTE|.*COMMISSION PAIEMENT|.*COTISATION TRI).*((?:\d{1,3}(?: \d{3})*|\d+),\d{2})\b"


class BankStatementLine:
    def __init__(self, raw):
        self.date   = raw.split(" ")[0]
        self.amount = float(re.findall(NUMBER_PATTERN, raw)[-1].replace(" ", "").replace(",", "."))
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
        self.total_incomes = -1
        self.total_expenses = -1

    def extract_totals(self, line):
        print(line)
        line = line.replace("Total des opérations", "").strip()
        pattern = r"((?:\d{1,3}(?: \d{3})*|\d+),\d{2})"
        match = re.findall(pattern, line)
        if match:
            print(match)
            num1, num2 = match
            self.total_expenses = float(num1.replace(" ", "").replace(",", "."))
            self.total_incomes = float(num2.replace(" ", "").replace(",", "."))

    def extract_data(self):
        doc = convert_from_path(self.absolute_path)
        lines = []
        stop = False
        search_for_total = False

        print("Extracting: " + self.file_name)
        for page_number, page_data in enumerate(doc):
            if stop:
                break
            print("Page: " + str(page_number + 1) + '/' + str(len(doc)))
            txt = pytesseract.image_to_string(page_data, lang='fra').encode('utf-8')
            decoded = txt.decode('utf-8')
            for line in decoded.split("\n"):
                print(line)
                if "Total des opérations" in line:
                    if re.fullmatch(r".*(\d* \d\d\d|\d*),\d\d (\d* \d\d\d|\d*),\d\d", line):
                        self.extract_totals(line)
                        stop = True
                        break
                    else:
                        search_for_total = True
                elif re.fullmatch(r"(\d* \d\d\d|\d*),\d\d (\d* \d\d\d|\d*),\d\d", line) and search_for_total:
                    self.extract_totals(line)
                    stop = True
                    break

                if re.fullmatch(DATA_PATTERN, line):
                    extracted = BankStatementLine(line)
                    lines.append(extracted)
        print("Total des opérations : " + str(self.total_expenses) + " " + str(self.total_incomes))
        print()
        return lines

class BankStatementConverter:
    def __init__(self, input_folder):
        self.input_folder = input_folder

    def extract_to_xlsx(self):
        files = sorted(self.get_folder_content(), key=lambda file: file.file_name)
        extracted = []

        for f in files:
            extracted.append(f.extract_data())
        
        workbook = xlsxwriter.Workbook('Expenses01.xlsx')
        worksheet = workbook.add_worksheet()
        row = 0
        incomes = 0
        expenses = 0
        
        for i in range(len(extracted)):
            f = extracted[i]
            file_incomes = 0
            file_expenses = 0
            for line in f:
                if line.type == "expense":
                    file_expenses += line.amount
                    line.save_to_worksheet(row, worksheet)
                else:
                    file_incomes += line.amount
                    line.save_to_worksheet(row, worksheet)
                row += 1
            incomes += file_incomes
            expenses += file_expenses
            print(f"Expected: {str(files[i].total_expenses)} / Actual: {str(file_expenses)}")
            print(f"Expected: {str(files[i].total_incomes)} / Actual: {str(file_incomes)}")
            if abs(files[i].total_expenses - file_expenses) > 0.01 or abs(files[i].total_incomes - file_incomes) > 0.01:
                print('INCONSITENCY IN FILE ' + files[i].file_name)
        print("Incomes: ", str(incomes))
        print("Expenses: ", str(expenses))
        print("Balance: " + str(incomes - expenses))
        workbook.close()
            
    def get_folder_content(self):
        files = [f for f in listdir(self.input_folder) if isfile(join(self.input_folder, f))]
        
        return [BankStatementFile(join(self.input_folder, f), f) for f in files]

if __name__ == '__main__':
    converter = BankStatementConverter('/path/to/bank/statements/folder')
    converter.extract_to_xlsx()
    # for entry in content:
    #     entry.extract_data()
