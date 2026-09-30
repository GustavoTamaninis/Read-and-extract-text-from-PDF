import os
from openpyxl import Workbook
import pdfplumber
import re
from datetime import datetime

directory = 'pdf_invoices'
files = os.listdir(directory)
files_quantity = len(files)

if files_quantity == 0:
    raise Exception("No pdf files found in the directory")

wb = Workbook() # Workbook é o arquivo .xlsx
ws = wb.active # Worksheet é a aba de uma planilha
ws.title = 'Invoice Imports'

# Estabelecendo nomes das colunas:
ws['A1'] = 'Invoice #'
ws['B1'] = 'Date'
ws['C1'] = 'File Name'
ws['D1'] = 'Status'

# Verificando quantas linhas de informações há:
last_empty_line = 1
while ws['A' + str(last_empty_line)].value is not None:
    last_empty_line += 1

# Passando os dados para o arquivo Excel:
for file in files:
    with pdfplumber.open(directory + "/" + file) as pdf: # Abre o arquivo pdf inteiro;
        first_page = pdf.pages[0]
        pdf_text = first_page.extract_text() # extrai as informações requeridas

    invoice_num_re_pattern = r'INVOICE #(\d+)'
    invoice_date_re_pattern = r'DATE: (\d{2}/\d{2}/\d{4})'

    match_num = re.search(invoice_num_re_pattern, pdf_text) # Não armazena o texto, mas o grupo.
    match_date = re.search(invoice_date_re_pattern, pdf_text)

    if match_num:
        invoice_num = match_num.group(1) # Armazena o texto
        ws[f'A{last_empty_line}'] = invoice_num
    else:
        ws[f'A{last_empty_line}'] = "Couldn't find invoice number"

    if match_date:
        invoice_date = match_date.group(1)
        ws[f'B{last_empty_line}'] = invoice_date
    else:
        ws[f'B{last_empty_line}'] = "Couldn't find invoice date"

    ws[f'C{last_empty_line}'] = file

    ws[f'D{last_empty_line}'] = "Completed"

    last_empty_line += 1

full_now = str(datetime.now()).replace(':', '-')
dot_index = full_now.index(".")
now = full_now[:dot_index] # Armazena o datetime sem os milissegundos

wb.save(f"Invoices - {now}.xlsx")