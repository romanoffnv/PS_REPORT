import time
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
import win32com
print(win32com.__gen_path__)

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
    L_block_rows, L_crews, L_fields, L_units = [], [], [], []
    while True:
        if 'ГНКТ №' in str(ws.Cells(row, 1).Value):
            L_block_rows.append(row)
            L_crews.append(ws.Cells(row, 1).Value)
            L_fields.append(ws.Cells(row, 5).Value)
        row += 1
        if 'ГНКТ №32' in str(ws.Cells(row, 1).Value):
            break

    # Equalizing L_crews and L_fields lengths to the length of units blocks
    L_crews = L_crews[:-1]
    L_fields = L_fields[:-1]
    
    # The function returns the trucks from all blocks
    def data_acquirer(srow, erow, col):
        L = []
        row = srow
        while True:
            L.append(ws.Cells(row, col).Value)
            row += 1
            if row == erow:
                break
        return L

    # Listing block beginning and block ending rows to be thrown as params into data_acquirer fucn
    L_startIndex = L_block_rows[:-1]
    L_endIndex = L_block_rows[1:]
    
    # Listing all trucks by blocks by running data data_acquirer func with block beg and end rows
    # L_units = []
    for j, k in zip(L_startIndex, L_endIndex):
        L_units.append(data_acquirer(j, k, 6))
        
    wb.Close(True)
    xl.Quit()

    
    # Listing lengths of block items by locations
    L_group_len = []
    for i in L_units:
        L_group_len.append(len(i))
    
    # Stretching crews over units blocks by multiplying crew name by block lengths
    L_crews = [(i + '**').split('**') * j for i, j in (zip(L_crews, L_group_len))]
    L_crews = list(itertools.chain.from_iterable(L_crews))
    L_crews = list(filter(None, L_crews))

    # Stretching fields over units blocks by multiplying crew name by block lengths
    L_fields = [str(x) for x in L_fields]
    L_fields = [(i + '**').split('**') * j for i, j in (zip(L_fields, L_group_len))]
    L_fields = list(itertools.chain.from_iterable(L_fields))
    L_fields = list(filter(None, L_fields))

    # Unpacking L_units
    L_units = list(itertools.chain.from_iterable(L_units))
    
    # Building df
    df = pd.DataFrame(zip(L_crews, L_units, L_fields), columns=['Crews', 'Units', 'Fields'])
    
    
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS Units_Locs_Raw")
    df.to_sql(name='Units_Locs_Raw', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
    
c