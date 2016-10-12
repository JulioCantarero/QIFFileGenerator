# -*- coding: utf-8 -*-
"""
Created on Fri Sep 23 23:33:34 2016

@author: Julio

v0.1:   Program TK to select files of all files in folder if Cancel pressed
        Program export to QIF from intermdiate List structs        
"""

from datetime import *
from dateutil.parser import parse
from tkinter import *
from tkinter import ttk
from tkinter import filedialog 
import os
import re
from string import Template
import bs4

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
accounts_transactions = [ [ ["2016/05/15","26,56", "Liquidación intereses  al 0,80 %  realizado por COINC", 
                                     [ ["Gastos:Impuestos:IRPF", "IMPUESTO Retención por Intereses del IRPF (19,0%)", "-6,22"],
                                       ["Ingresos:Intereses:Cuentas", "Intereses Brutos AL 0,80%", "32,78" ]
                                     ]
                            ],
                            [ "2016/06/30", "-111,1", "Prueba Ñoquis", False]
                          ]
                        ]
date_column = 0
amount_column = 2
description_column = 3

root = Tk()
root.withdraw() #use to hide tkinter window

currdir = r"D:\Test\QIF" 
filenames = filedialog.askopenfilename(filetypes = (("All Files","*.*"), ("Excel Files", ("*.xls","*.xlsx"))), 
                            multiple=True, parent=root, initialdir=currdir, 
                            title='Please select the Files to process or Cancel if you prefer to choose a working directory:')

if not filenames:
    target_directory = filedialog.askdirectory (parent=root, initialdir=currdir, mustexist=True,
                                 title='Please select the Folder with the files to be processed:')
    filenames = os.listdir (target_directory.replace('/','\\'))
    

for file_to_process in filenames:
    soup = bs4.BeautifulSoup(open(file_to_process.replace('/','\\')), 'html.parser')
    transactions_table = []
    tables = soup.findAll("table")
    for table in tables:
        rows = table.findAll('tr')
        for row in rows:
            cells = row.findAll('td')
            transaction = ['','','', False]
            for index, table_cell in enumerate(cells):
                if table_cell.string:
                    cell_contents = table_cell.string.strip(' \t')
                    if index == date_column:
                        try:
                            transaction[0] = parse(cell_contents, dayfirst = True)
                        except ValueError:
                            print('The Cell contents "' + cell_contents + '" could not be parsed as a date')
                    elif index == amount_column:
                        transaction[1] = cell_contents
                    elif index == description_column:
                        transaction[2] = cell_contents
            if isinstance(transaction[0], datetime):
                transactions_table.append(transaction)
print(transactions_table)
accounts_transactions[0] = transactions_table

#Import list of Account Names and types from the original GNUCAsh file

amount_format_regex = r'([-]?\d+(?:\.\d{3})*(?:\,\d+)*)'
date_format_regex = r'(\d{1,2})[-\s\/](\d{1,2})[-\s\/](\d{2,4})'

    
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