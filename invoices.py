import os
from openpyxl import Workbook

directory = 'pdf_invoices'
files = os.listdir(directory)
files_quantity = len(files)

if files_quantity == 0:
    raise Exception("No pdf files found in the directory")

wb = Workbook() # Workbook é o arquivo .xlsx
ws = wb.active # Worksheet é a aba de uma planilha
ws.title = 'Invoice Imports'

print(files_quantity)
print(ws.title)

# Estabeler nomes das colunas:
ws['A1'] = 'Invoice #'
ws['B1'] = 'Date'
ws['C1'] = 'File Name'
ws['D4'] = 'Status'

# Verificando quantas linhas de informações há:
last_empty_line = 1
while ws['A' + str(last_empty_line)] is not None:
    last_empty_line += 1


