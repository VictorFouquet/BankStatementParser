import os
import re

from PIL import Image
from pdf2image import convert_from_path
from pytesseract import pytesseract


def extract_text(_img):
    doc = convert_from_path(_img)
    path, fileName = os.path.split(_img)
    
    fileBaseName, fileExtension = os.path.splitext(fileName)
    pattern = r"\b\d{2}/\d{2}.*\d*,\d{2}\b"
    expense = r"\b\d{2}/\d{2} (ACHAT|VIREMENT.*À|PRELEVEMENT).*\d*,\d{2}\b"

    for page_number, page_data in enumerate(doc):
        txt = pytesseract.image_to_string(page_data, lang='fra').encode('utf-8')
        decoded = txt.decode('utf-8')
        for line in decoded.split("\n"):
            if re.fullmatch(pattern, line):
                date = line.split(" ")[0]
                amount = line.split(" ")[-1]
                if re.fullmatch(expense, line):
                    print(date + ": -" + amount)
                else:
                    print(date + ": +" + amount)


if __name__ == '__main__':
    extract_text('/path/to/file.pdf')
