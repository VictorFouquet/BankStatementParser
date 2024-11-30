# Bank Statement Parser

Python script using tesseract to convert a bank statement pdf file to excel.

As the formatting of the handled bank statement is predefined, this script is not a generic utility, the parsing logic should be adapted to fit your needs.

Extracted data will contain only two columns, the date and the amount of the operations, amount being positive for incomes and negative for expenses.
