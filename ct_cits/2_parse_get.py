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

# Get the Excel Application COM object
# xl = EnsureDispatch('Excel.Application')
# wb = xl.Workbooks.Open(f"{os.getcwd()}\\1.xlsx")
# Sheets = wb.Sheets.Count
# ws = wb.Worksheets(Sheets)

# Making connections to DBs

db = sqlite3.connect('cits.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()


def main():
     # Pulling Units_Locs_Raw.db into lists
    L_crews = cursor.execute("SELECT Crews FROM Units_Locs_Raw").fetchall()
    L_units = cursor.execute("SELECT Units FROM Units_Locs_Raw").fetchall()
    L_locs = cursor.execute("SELECT Fields FROM Units_Locs_Raw").fetchall()
    
    # Cleaning L_units
    
    L_cleanwords = ['Цель работ:', 'профессия', 'Бурильщик', 'Пом.бур', 'Маш-т', '\n', 'гос№',
                  '№', '-', '\.']
    for i in L_cleanwords:
        L_units = [re.sub(i, ' ', x).strip() for x in L_units if x != None]
    L_units = [x for x in L_units if x != '']
    
    # Splitting merged cells
    L_units = [re.sub('\s+', ' ', x) for x in L_units]
    L_units = [re.sub(';', ',', x) for x in L_units]
    L_units = [re.sub('\+', ',', x) for x in L_units]
    L_units = [re.sub('в пути', ',', x) for x in L_units]
    # L_units = [re.sub('86 86', '86 86,', x) for x in L_units]
    L_units = [re.sub('86', '86,', x) for x in L_units if '2386' not in x]
    L_units = [re.sub('трал', 'трал ', x) for x in L_units]
    
    L_units = [x.split(',') for x in L_units]
    L_units = list(itertools.chain.from_iterable(L_units))
    
    pprint(L_units)
    pprint(len(L_units))
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))