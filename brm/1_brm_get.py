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
    # Get blocks by frac crews
    # Listing rows of the blocks by crews i.e. 'ГРП-1'
    print('Listing rows of the blocks by crews i.e. ГРП-1')
    row = 4
    L_block_rows, L_crew_names = [], []
    while True:
        if 'ГРП-' in str(ws.Cells(row, 1).Value):
            L_block_rows.append(row)
            L_crew_names.append(ws.Cells(row, 1).Value)
        row += 1
        if 'Ловильный сервис' in str(ws.Cells(row, 1).Value):
            L_block_rows.append(row) 
            break

    
    
    # The function returns the list of trucks from all frac crew blocks
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
    print('Listing block beginning and block ending rows')
    L_startIndex = L_block_rows[:]
    L_endIndex = L_block_rows[1:]
    
    
    # Listing all trucks and trailers by crew blocks by running data data_acquirer func with 
    # block beginning and end rows
    print('Listing all trucks and trailers by crew blocks')
    # L_plates - list for trucks, L_plates2 - list for trailers
    L_units, L_plates, L_plates2, L_locs = [], [], [], []
    for j, k in zip(L_startIndex, L_endIndex):
        L_units.append(data_acquirer(j, k, 2))
        L_plates.append(data_acquirer(j, k, 3))
        L_plates2.append(data_acquirer(j, k, 4))
        L_locs.append(data_acquirer(j, k, 11))
    
    
    
    # Listing lengths of block items by locations
    L_group_len = []
    for i in L_locs:
        L_group_len.append(len(i))
    
    # Stretching crews over locations blocks by multiplying crew name by block lengths
    print('Stretching crews over locations blocks')
    L_crews = [(i + '**').split('**') * j for i, j in (zip(L_crew_names, L_group_len))]
    L_crews = list(itertools.chain.from_iterable(L_crews))
    L_crews = list(filter(None, L_crews))
   
    
    print('Closing xls file')
    wb.Close(True)
    xl.Quit()
    
    # Unpacking block lists for trucks, trailers and locations
    print('Unpacking block lists for trucks, trailers and locations')
    L_units = list(itertools.chain.from_iterable(L_units))
    L_plates = list(itertools.chain.from_iterable(L_plates))
    L_plates2 = list(itertools.chain.from_iterable(L_plates2))
    L_locs = list(itertools.chain.from_iterable(L_locs))
    
    
    # Building df of raw data for to pick crew names and locations from later
    print('Building df of raw data for to pick crew names and locations from later')
    df = pd.DataFrame(zip(L_crews, L_units, L_plates, L_plates2, L_locs), columns=['Crews', 'Units', 'Plates_1', 'Plates_2', 'Locs'])
    
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS Units_Locs_Raw")
    df.to_sql(name='Units_Locs_Raw', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()

    print('1_brm_get is complete')
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))