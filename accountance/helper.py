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
    L_total_accountance = []
    row = 14
    col = 2
    while True:
        L_total_accountance.append(ws1.Cells(row, col).Value)
        row += 1
        if ws1.Cells(row, col).Value == None:
            break
    pprint(L_total_accountance)
    pprint(len(L_total_accountance))

    wb.Close(True)
    xl.Quit()

    L_total_accountance2 = [x for x in L_total_accountance]
      # Post the data acquired into the db as accountance_1
    cursor.execute("DROP TABLE IF EXISTS accountance_all;")
    cursor.execute("""
                    CREATE TABLE IF NOT EXISTS accountance_all(
                    Item1 text,
                    Item2 text
    
                                )
                    """)
    cursor.executemany("INSERT INTO accountance_all VALUES (?, ?)", zip(L_total_accountance, L_total_accountance2))
    
    
    db.commit()
    db.close()

if __name__ == '__main__':
    main()