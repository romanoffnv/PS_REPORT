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
wb = xl.Workbooks.Open(f"{os.getcwd()}\\brm.xlsx")
Sheets = wb.Sheets.Count
ws = wb.Worksheets(Sheets)

# Making connections to DBs
db = sqlite3.connect('brm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

# Pandas
pd.set_option('display.max_rows', None)

def main():
    # Get frac blocks
    # Listing crew names and row numbers of the crew blocks
    row = 4
    L_block_rows, L_crew_names = [], []
    while True:
        if 'ГРП-' in str(ws.Cells(row, 1).Value):
            L_block_rows.append(row)
            L_crew_names.append(ws.Cells(row, 1).Value)
        row += 1
        if 'Ловильный сервис' in str(ws.Cells(row, 1).Value):
            break

    
    
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

    # Listing block beginning and block ending rows to be thrown as params into data_acquirer func
    L_startIndex = L_block_rows[:-1]
    L_endIndex = L_block_rows[1:]
    
    # Listing all trucks by blocks by running data data_acquirer func with block beg and end rows
    L_units, L_units2, L_locs = [], [], []
    for j, k in zip(L_startIndex, L_endIndex):
        L_units.append(data_acquirer(j, k, 3))
        L_units2.append(data_acquirer(j, k, 4))
        L_locs.append(data_acquirer(j, k, 11))
        
    
    L_group_len = []
    for i in L_units:
        L_group_len.append(len(i))
    
    
    L_crews = [(i + '**').split('**') * j for i, j in (zip(L_crew_names, L_group_len))]
    L_crews = list(itertools.chain.from_iterable(L_crews))
    L_crews = list(filter(None, L_crews))
    
    
    wb.Close(True)
    xl.Quit()
    
    L_units = list(itertools.chain.from_iterable(L_units))
    L_units2 = list(itertools.chain.from_iterable(L_units2))
    L_locs = list(itertools.chain.from_iterable(L_locs))
    
    # Build df for location picking
    df = pd.DataFrame(zip(L_crews, L_units, L_units2, L_locs), columns=['Crews', 'Units_1', 'Units_2', 'Locs'])
    print(df)
   
    # Post df to DB
    cursor.execute("DROP TABLE IF EXISTS Units_Locs_Raw")
    df.to_sql(name='Units_Locs_Raw', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()


if __name__ == '__main__':
    main()