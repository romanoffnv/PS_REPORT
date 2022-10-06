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
    L_block_rows, L_crew_names = [], []
    while True:
        if 'ГНКТ №' in str(ws.Cells(row, 1).Value):
            L_block_rows.append(row)
            L_crew_names.append(ws.Cells(row, 1).Value)
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

    
    def list_generator(L):
        L_ct = []
        
        for i in L:
            if type(i) == str:
                L_ct.append(i.replace(';', ','))
        
        L_ct = [x.replace(':', ',') for x in L_ct]
        
        
        # Splits
        
        L_ct = [x.replace(' 86- ', ' 86 split ') for x in L_ct ]
        L_ct = [x.replace(',', ' split ') for x in L_ct]
        L_ct_compressed = [x.split('split') for x in L_ct]
        L_ct_compressed = list(itertools.chain.from_iterable(L_ct_compressed))
        L_ct_compressed = [x.split('+') for x in L_ct_compressed]
        L_ct_compressed = list(itertools.chain.from_iterable(L_ct_compressed))
        L_ct_compressed = [x.split('RUS') for x in L_ct_compressed]
        L_ct_compressed = list(itertools.chain.from_iterable(L_ct_compressed))
       
        
        return L_ct_compressed
    
    # Splitting data in strings into the lists(which get extracted) into the list by running list_generator func
    L_units_temp = []
    for i in L_units:
        L_units_temp.append((list_generator(i)))
    L_units = [x for x in L_units_temp]
    L_units_temp.clear()
    
    
    def list_cleaner(i):
        L_ct_clean = [x.strip() for x in i]
        
       
        
        # Replacing crappy unit names into omnicomm smth
        D_replacers = json.load(open('D_replacers.json'))
        for k, v in D_replacers.items():
            L_ct_clean = [x.replace(k, v) for x in L_ct_clean]
        
        
        # removing trash 
        D_patterns = json.load(open('D_patterns.json'))
        for k, v in D_patterns.items():
            L_ct_clean = [re.sub(k, v, x) for x in L_ct_clean]
        
        # extracting plates from brackets
        L_ct_clean = [x.split('(') for x in L_ct_clean]
        L_ct_clean = list(itertools.chain.from_iterable(L_ct_clean)) 
        
        # removing spaces
        L_ct_clean = [''.join(re.sub('\s+', ' ', x)).strip() for x in L_ct_clean]

        # Removing items that don't have numbers (i.e. plates)
        pattern_D = re.compile(r'\d')
        L_ct_clean = [x for x in L_ct_clean if re.findall(pattern_D, str(x))]
        

        # The ultimate list should contain items of 3 types: 100% - МЗКТ УУ 0775 86, 80% - Автокран 766, 50% - 232
        return L_ct_clean
    
    # Running list_cleaner func to clean up trash
    for i in L_units:
        L_units_temp.append((list_cleaner(i)))
    L_units = [x for x in L_units_temp]
    L_units_temp.clear()
    
   
    # pprint(L_units)
    # print(len(L_units))
    
    # Mulitplying units by crews
    L_crews = [(k + '**').split('**') * len(v) for k, v in zip(L_crew_names, L_units)]
    L_crews = list(itertools.chain.from_iterable(L_crews))
    L_crews = list(filter(None, L_crews))
    
    # Merging units lists into one list
    L_units = list(itertools.chain.from_iterable(L_units))
       
    # Post into db
    
    cursor.execute("DROP TABLE IF EXISTS Trucks")
    cursor.execute("""
	CREATE TABLE IF NOT EXISTS Trucks(
		Dept text,
	    Unit text)

              """)
    cursor.executemany("INSERT INTO Trucks VALUES (?, ?)", zip(L_crews, L_units))

    db.commit()
    db.close()
    
    pprint(L_units)
if __name__ == '__main__':
    main()