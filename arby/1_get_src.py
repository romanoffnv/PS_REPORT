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
        
        df = pd.DataFrame(zip(L_drivers, L_plates), columns =['Drivers', 'Plates'])
        return df
    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions_long = [126, 156, 158, 174, 186, 188, 196, 197, 797]
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
    def list_getter(df_type, arg):
        L = df_type[arg].tolist()
        return L
    def df_maker(*args):
        if len(args) == 2:
            df = pd.DataFrame(zip(*args), columns = ['Plates', 'Drivers'])
        elif len(args) == 3:
            df = pd.DataFrame(zip(*args), columns = ['Plates', 'PI', 'Drivers'])
        return df
    # ===================================================================================
    
    # Reading xls
    df_ct = xlsreader("arby_ct.xlsx")
    df_trans = xlsreader("arby_trans.xlsx")
    
    # Listing columns, making df for ct
    L_plates = list_getter(df_ct, 'Гос. №')
    L_drivers = list_getter(df_ct, 'Ответственный')
    df1 = df_maker(L_plates, L_drivers)
    
    # Listing columns, making df for trans
    L_plates = list_getter(df_trans, 'Гос №')
    L_drivers = list_getter(df_trans, 'Ответственный')
    df2 = df_maker(L_plates, L_drivers)
    
    # Merging dfs
    df = pd.merge(df1, df2, how="outer")
    
    # Listing plates and drivers from total df 
    L_plates = list_getter(df, 'Plates')
    L_drivers = list_getter(df, 'Drivers')
    
    # Making replacements
    L_plates = replacer(L_plates)
    
    # Prepping drivers
    L_drivers = driversprepper(L_drivers)
    
    # Packing dubs
    df = dubspacker(L_plates, L_drivers)
    
    # Listing plates and drivers after dubs were packed
    L_plates = list_getter(df, 'Plates')
    L_drivers = list_getter(df, 'Drivers')
    
    # Transforming plates into 123abc format 
    L_PI = transform_plates(L_plates)
    df = df_maker(L_plates, L_PI, L_drivers)
    
    # Post df to DB
    cursor.execute("DROP TABLE IF EXISTS ct_drivers")
    df.to_sql(name='ct_drivers', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))