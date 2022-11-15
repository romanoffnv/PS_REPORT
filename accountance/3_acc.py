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



# Making connections to DBs
db = sqlite3.connect('arby.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

# Pandas
pd.set_option('display.max_rows', None)

def main():
    xl = pd.ExcelFile('gen_report.xls')
    
    pprint(xl.sheet_names)  # see all sheet names
    def sheets_parser(service):
        df = xl.parse(service)
        df = df.drop(range(0, 11))
        L_units = df['Unnamed: 3']
        L_plates = df['Unnamed: 9']
        L_comments = df['Unnamed: 15']
        df = pd.DataFrame(zip(L_units, L_plates, L_comments))
        return df
    
    df1 = sheets_parser('ГНКТ')
    # df2 = xl.parse('ГРП')
    pprint(df1)
    # df3 = xl.parse('ТР.Служба')
    # pprint(df1)

if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
