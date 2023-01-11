import time
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
wb = xl.Workbooks.Open(f"{os.getcwd()}\\acc.xlsx")
ws1 = wb.Worksheets(1)

# Pandas
pd.set_option('display.max_rows', None)

# db connections
db = sqlite3.connect('data.db')
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
                
        
               
        wb.Close(True)
        xl.Quit()
        # Correlating the number of mols to match the number of items
        # after mols_scratch, items and counts lists are collected, L_mols should be populated by
        # mulitplying each mols_scratch element by counts
        L_mols = [(i + '**').split('**') * j for i, j in (zip(L_mols_scratch, L_counts))]
        L_mols = list(itertools.chain.from_iterable(L_mols))
        L_mols = list(filter(None, L_mols))

        # Assigning indexed for the items
        # L_indexes = []
        # for i in range(0, len(L_mols)):
        #     L_indexes.append(i)
       
        # print(L_indexes)
        L_plates = []
        for i in L_items:
            L_plates.append('')

        df = pd.DataFrame(zip(L_mols, L_items, L_plates), columns=['Mols', 'Units', 'Plates'])
        df = df.dropna(how='any', subset=['Units'], thresh=1)
        print(df)
        # Posting df to DB as raw accountance_1
        print('Posting df to DB')
        cursor.execute("DROP TABLE IF EXISTS accountance_1")
        df.to_sql(name='accountance_1', con=db, if_exists='replace', index=False)
        db.commit()

        df = df[df["Units"].str.contains("Гамма плотномер") == False]
        # Posting df to DB as cleaned for "Гамма плотномер" (because it interferes with the trailer plates)as accountance_2
        print('Posting df to DB')
        cursor.execute("DROP TABLE IF EXISTS accountance_2")
        df.to_sql(name='accountance_2', con=db, if_exists='replace', index=False)
        db.commit()
        db.close()
        

if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
