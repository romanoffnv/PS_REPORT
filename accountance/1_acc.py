import xlsxwriter
from win32com.client.gencache import EnsureDispatch
import os
import re
from pprint import pprint
import pandas as pd
from functools import reduce
import itertools
import sqlite3
import win32com
print(win32com.__gen_path__)

# Excel connection  
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\acc.xls")
ws1 = wb.Worksheets(1)

# Pandas
pd.set_option('display.max_rows', None)

# db connections
db = sqlite3.connect('accountance.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

def main():
    # Getting all data from the column 1 (mols and items) and filtering mols out of it
        row = 14
        L_mols_scratch, L_items = [], []
        while True:
            if ws1.Cells(row, 2).Font.Bold == True and ("вич" in ws1.Cells(row, 2).Value or "вна" in ws1.Cells(row, 2).Value):
                L_mols_scratch.append(ws1.Cells(row, 2).Value)
                L_items.append('****')
            elif ws1.Cells(row, 2).Font.Bold != True:
                L_items.append(ws1.Cells(row, 2).Value)
            if ws1.Cells(row, 2).Value == None:
                break
            row += 1
        
        # Extracting plates from paranthesis
        L_items = [x.split('(') for x in L_items if x != None]
        L_items = list(itertools.chain.from_iterable(L_items))
        
        L_items = L_items[1:]
        L_items.append('****')
        
        L_counts = []
        counter = 0
        
        for i in L_items:
            if i == '****':
                L_items.remove(i)
                L_counts.append(counter)
                counter = 0
            counter += 1    
                
        
        sumL_counts = reduce(lambda x, y: x + y, L_counts)
        
      
        # print(f'this is mols len {len(L_mols_scratch)}')
        # print(f'this is items len {len(L_items)}')
        # print(f'this is sum of items  {sumL_counts}')
        # print(f'this is the list of ranges {L_counts}')
        # the last item of the L_counts list is 24, when should be 23 
        # print(f'this is the length of ranges list {len(L_counts)}')
        
        wb.Close(True)
        xl.Quit()
        # Correlating the number of mols to match the number of items
        # after mols_scratch, items and counts lists are collected, L_mols should be populated by
        # mulitplying each mols_scratch element by counts
        L_mols = [(i + '**').split('**') * j for i, j in (zip(L_mols_scratch, L_counts))]
        L_mols = list(itertools.chain.from_iterable(L_mols))
        L_mols = list(filter(None, L_mols))
       
       
        # Post the data acquired into the db as accountance_1
        cursor.execute("DROP TABLE IF EXISTS accountance_1;")
        cursor.execute("""
                    CREATE TABLE IF NOT EXISTS accountance_1(
                    Responsible text,
                    Item text
    
                                )
                    """)
        cursor.executemany("INSERT INTO accountance_1 VALUES (?, ?)", zip(L_mols, L_items))
        cursor.execute("DELETE FROM accountance_1 WHERE Item IS NULL OR trim(Item) = ''")
    
        db.commit()
        db.close()


        


if __name__ == '__main__':
    main()
