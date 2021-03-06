﻿# -*- coding: utf-8 -*-
"""
Created on Fri Sep 23 23:33:34 2016

@author: Julio

v0.1:   Program TK to select files of all files in folder if Cancel pressed
        Program export to QIF from intermdiate List structs        
"""

from datetime import *
from dateutil.parser import parse
import locale

from tkinter import *
from tkinter import ttk
from tkinter import filedialog 
import os
import re
from string import Template
import bs4

#donde las variables de entrada son strings de 4, 4 y 10 caracteres respectivamente
def digitos_control(entidad, oficina, cuenta):
    def proc(digitos):
        if not digitos.isdigit() or len(digitos) != 10:
            raise ValueError('Debe ser número de 10 digitos: %s' % digitos)
        factores = [1, 2, 4, 8, 5, 10, 9, 7, 3, 6]
        resultado = 11 - sum(int(d)*f for d,f in zip(digitos, factores)) % 11
        if resultado == 10:  return 1
        if resultado == 11:  return 0
        return resultado
    return '%d%d' % (proc('00'+entidad+oficina), proc(cuenta))
    
def search_bank_account_number(file_text):
    """Search all the file for a sequence of 20 digits and check the DC field"""

    bank_account_format_in_spain = re.compile('\d{4}[ -]?\d{4}[ -]?\d{2,4}[ -]?\d{4,6}[ -]?\d{4}')
    account_code = bank_account_format_in_spain.search(file_text)
    if account_code:
        code_no_spaces = account_code.group(0).replace(' ','').replace('-','')
        if len(code_no_spaces) == 20:
            if digitos_control(code_no_spaces[0:4],code_no_spaces[4:8],code_no_spaces[10:]) == code_no_spaces[8:10]:
                return code_no_spaces

def parse_HTML_table_row_for_header(row):
    """Check the cells (<td>) objects in the Table row and find the columns where the transaction components are"""    
    
    date_header = re.compile('fecha', re.I)
    amount_header = re.compile('importe', re.I)
    description_header = re.compile('concepto', re.I)
    transaction_columns = {}
    cells = row.find_all('td')
    for index, table_cell in enumerate(cells):
        if table_cell.string:
            cell_contents = table_cell.string.strip(' \t')
            if date_header.search(cell_contents):
                transaction_columns['date'] = index
            elif amount_header.search(cell_contents):
                transaction_columns['amount'] = index
            elif description_header.search(cell_contents):
                transaction_columns['description'] = index
    if len(transaction_columns) == 3:
        return transaction_columns
    else:
        return False

def parse_HTML_table_row_for_a_transaction(row, columns):
    """Extract the transaction components from the columns indicated"""

    cells = row.find_all('td')
    transaction = ['','','', False]    # I will record every transaction in an array like [Date, Amount, Description, PeerAccount/SplitTransactions]
    for index, table_cell in enumerate(cells):
        if table_cell.string:
            cell_contents = table_cell.string.strip(' \t')
            if index == columns['date']:
                try:
                    transaction[0] = parse(cell_contents, dayfirst = True)
                except ValueError:
                    print('The Cell contents "' + cell_contents + '" could not be parsed as a date')
            elif index == columns['amount']:
                transaction[1] = locale.atof(cell_contents)
                #ValueError: could not convert string to float: 'Importe EUR'
            elif index == columns['description']:
                transaction[2] = cell_contents
    if isinstance(transaction[0], datetime):
        return transaction
    else:
        return False


qifAccountHdrTemplate = Template("""\
!Account
N$account_name
^
!Type:$account_type
""")

qifTransactionTemplate =  Template("""\
D$date
T$amount
M$description
""")

qifPairedAccountTemplate =  Template("""\
L$paired_account
""")

qifSplitEntryTemplate =  Template("""\
S$split_account
E$description
$$$amount
""")

accounts_names = ["Activo:Activo Circulante:COINC (Bankinter):Cuenta Corriente", "Pasivo:Tarjetas de Crédito:Santander Visa Cuenta Comun"]
accounts_types = ["Bank", "CCard"]
accounts_transactions = []

locale.setlocale(locale.LC_ALL, '')         # Used in locale to use the default locale, as returned in locale.getdefaultlocale()

root = Tk()
root.withdraw() #use to hide tkinter window

#Import list of Account Names and types from the original GNUCAsh file

currdir = r"D:\Test\QIF" 
filenames = filedialog.askopenfilename(filetypes = (("All Files","*.*"), ("Excel Files", ("*.xls","*.xlsx"))), 
                            multiple=True, parent=root, initialdir=currdir, 
                            title='Please select the Files to process or Cancel if you prefer to choose a working directory:')

if not filenames:
    target_directory = filedialog.askdirectory (parent=root, initialdir=currdir, mustexist=True,
                                 title='Please select the Folder with the files to be processed:')
    filenames = os.listdir (target_directory.replace('/','\\'))
    

for file_to_process in filenames:
    
    # Determine the type of file selected from the allowed types: Excel(HTML), Excel, PDF, HTML?, others?
    
    # If Excel(HTML) or HTML? parse with Beautiful Soup
    soup = bs4.BeautifulSoup(open(file_to_process.replace('/','\\')), 'html.parser')
    
    # For each file, identify which account the file corresponds to. Search for the CCC    
    file_plain_text = soup.get_text()
    print(search_bank_account_number(file_plain_text))
    
    transactions_table = []
    tables = soup.find_all("table")
    for table in tables:
        rows = table.find_all('tr')
        header_already_parsed = False
        for row in rows:
            if not header_already_parsed:
                columns = parse_HTML_table_row_for_header(row)
                if columns:
                    header_already_parsed = True
            else:
                transaction = parse_HTML_table_row_for_a_transaction(row, columns)
                if transaction:
                    transactions_table.append(transaction)
                    
    print(transactions_table)
    accounts_transactions.append(transactions_table)

output_file = open('FileToImport.qif', 'w')

for account, category, transactions in zip(accounts_names, accounts_types, accounts_transactions):
    output_file.write(qifAccountHdrTemplate.substitute(account_name=account, account_type= category))
    for transaction in transactions:
        output_file.write(qifTransactionTemplate.substitute(date=transaction[0].strftime("%Y/%m/%d"), amount= transaction[1], description=transaction[2]))
        if transaction[3]:
            if isinstance(transaction[3], str):
                output_file.write(qifPairedAccountTemplate.substitute(paired_account=transaction[3]))
            elif isinstance(transaction[3], list):
                for split_transaction in transaction[3]:
                    output_file.write(qifSplitEntryTemplate.substitute(split_account=split_transaction[0], \
                                        description= split_transaction[1], amount=split_transaction[2]))
        output_file.write("^\n")

output_file.close()