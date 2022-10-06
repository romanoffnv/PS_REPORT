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
wb = xl.Workbooks.Open(f"{os.getcwd()}\\1. СВОДКА СЛУЖБ ГНКТ ООО ПАКЕР СЕРВИС за Август 2022г..xlsx")
ws = wb.Worksheets('31.08.2022')

# Pandas
pd.set_option('display.max_rows', None)

# db connections
db = sqlite3.connect('total_match.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

def main():
    L1 = []
    row = 1
    col = 6
    while True:
        L1.append(ws.Cells(row, col).Value)
        row += 1
        if row == 270:
            break
    L2 = [x for x in L1]

    cursor.execute("DROP TABLE IF EXISTS items_cits;")
    cursor.execute("""
                        CREATE TABLE IF NOT EXISTS items_cits(
                        Item1 text,
                        Item2 text
                                )
                           """)
    cursor.executemany("INSERT INTO items_cits VALUES (?, ?)", zip(L1, L2))
    # refreshing database
    db.commit()
    # closing database
    db.close()


if __name__ == '__main__':
    main()