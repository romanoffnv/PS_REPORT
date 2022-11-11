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

# # Get the Excel Application COM object
# xl = EnsureDispatch('Excel.Application')
# wb = xl.Workbooks.Open(f"{os.getcwd()}\\arby_ct.xlsx")
# Sheets = wb.Sheets('Общая')
# ws = wb.Worksheets(Sheets)

# Making connections to DBs
db = sqlite3.connect('brm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

# Pandas
pd.set_option('display.max_rows', None)

def main():
    # Functions
    def xlsreader(xls):
        df = pd.read_excel(xls)
        return df
    def replacer(L_plates):
        L_replacers = ['№', '\s+']
        for i in L_replacers:
            L_plates = [''.join(re.sub(str(i), '', x)).strip() for x in L_plates]
        return L_plates
    def driversprepper(L_drivers):
        L_drivers = [''.join(x).strip() for x in L_drivers]
        return L_drivers
    def dubspacker(L_plates, L_drivers):
        dc = defaultdict(list)
        for i in range(len(L_plates)):
            item = L_plates[i]
            dc[item].append(L_drivers[i])
        dc = dict(zip(dc.keys(), map(set, dc.values())))
            
        L_plates, L_drivers  = zip(*dc.items())
        L_drivers = [', '.join(x) for x in L_drivers]
        
        df = pd.DataFrame(zip(L_drivers, L_plates), columns =['Driver', 'Plate'])
        return df
    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions_long = [126, 156, 158, 174, 186, 188, 196, 797]
        L_regions_short = ['01', '02', '03', '04', '05', '06', '07', '09']
        for i in L_regions_long:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 9 else x for x in plates]
        for i in L_regions_short:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        for i in range(10, 100):
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates

    # ===================================================================================
    
    # Reading xls
    df = xlsreader("arby_ct.xlsx")
    
    # Listing columns
    L_plates = df['Гос. №'].tolist()
    L_drivers = df['Ответственный'].tolist()

    # Making replacements
    L_plates = replacer(L_plates)
    
    # Preping drivers
    L_drivers = driversprepper(L_drivers)
    
    # Packing dubs
    df_ct = dubspacker(L_plates, L_drivers)
    
    # Reading xls
    df = xlsreader("arby_trans.xlsx")
    
    # Listing columns
    L_plates = df['Гос №'].tolist()
    L_drivers = df['Ответственный'].tolist()

     # Making replacements
    L_plates = replacer(L_plates)
    
    # Preping drivers
    L_drivers = driversprepper(L_drivers)
    
    # Packing dubs
    df_trans = dubspacker(L_plates, L_drivers)
    df = pd.merge(df_ct, df_trans, how="outer")

    L_plates = df['Plate'].tolist()
    L_PI = transform_plates(L_plates)
    pprint(L_PI)
    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))