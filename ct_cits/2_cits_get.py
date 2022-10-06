import json
import xlsxwriter
from win32com.client.gencache import EnsureDispatch
import os
import re
import sqlite3
from pprint import pprint
import pandas as pd
import itertools
from itertools import groupby
from collections import defaultdict
from collections import Counter

# Get the Excel Application COM object
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\1.xlsx")
Sheets = wb.Sheets.Count
ws = wb.Worksheets(Sheets)

# Making connections to DBs

db = sqlite3.connect('cits.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()


def main():
    # Listing crew names and row numbers of the crew blocks
    row = 4
    L_block_rows, L_crew_names, L_fields = [], [], []
    while True:
        if 'ГНКТ №' in str(ws.Cells(row, 1).Value):
            L_block_rows.append(row)
            L_crew_names.append(ws.Cells(row, 1).Value)
            L_fields.append(ws.Cells(row, 5).Value)
        row += 1
        if 'ГНКТ №32' in str(ws.Cells(row, 1).Value):
            break

    # The function returns the trucks from all blocks
    def data_acquirer(srow, erow):
        L = []
        row = srow
        while True:
            L.append(ws.Cells(row, 6).Value)
            row += 1
            if row == erow:
                break
        return L

    # Listing block beginning and block ending rows to be thrown as params into data_acquirer fucn
    L_startIndex = L_block_rows[:-1]
    L_endIndex = L_block_rows[1:]
    
    # Listing all trucks by blocks by running data data_acquirer func with block beg and end rows
    L_units = []
    for j, k in zip(L_startIndex, L_endIndex):
        L_units.append(data_acquirer(j, k))
    
    wb.Close(True)
    xl.Quit()

    json.dump(L_units, open("L_ct_units.json", 'w'))
    json.dump(L_crew_names, open("L_ct_crews.json", 'w'))
    json.dump(L_fields, open("L_ct_fields.json", 'w'))
   
    
   
    
    
if __name__ == '__main__':
    main()