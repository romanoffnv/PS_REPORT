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
    # pprint(xl.sheet_names)  # see all sheet names
    
    def sheets_parser(service):
        df = xl.parse(service)
        df = df.drop(range(0, 11))
        L_units = df['Unnamed: 3']
        L_plates = df['Unnamed: 9']
        L_comments = df['Unnamed: 15']
        df = pd.DataFrame(zip(L_units, L_plates, L_comments), columns=['Units', 'Plates', 'Comments'])
        return df
    
    df_ct = sheets_parser('ГНКТ')
    df_fr = sheets_parser('ГРП')
    df = pd.merge(df_ct, df_fr, how="outer")

    df_trans = sheets_parser('ТР.Служба')
    df = pd.merge(df, df_trans, how="outer")
    df = df.dropna(how='any', subset=['Comments'], thresh=1)

    pprint(df.describe())
    
    # List 'Units', 'Plates', 'Comments'
    L_units = df['Units'].tolist()
    L_plates = df['Plates'].tolist()
    L_comments = df['Comments'].tolist()
    L_plates = [str(x).strip() for x in L_plates]
    L_plates = [re.sub('\s+', '', x) for x in L_plates]
    L_comments = [str(x).strip() for x in L_comments]
    
    # Pull items from units if 'Нет данных'
    L_plates_temp = []
    for i, j in zip(L_units, L_plates):
        if j == 'Нетданных' or j == None:
            L_plates_temp.append(i)
        else:
            L_plates_temp.append(j)
   
   # Fishing out plates by regex from long sentences
    def regexBomber(x, L_units):
        
        L_plates_temp = []
        for i in L_units:
            if re.findall(x, str(i)):
                L_plates_temp.append(''.join(re.findall(x, str(i))))
            else:
                L_plates_temp.append(i)
                # print(i)
    
        L_units = [str(x).strip() for x in L_plates_temp]
        L_plates_temp.clear() 
            
        return L_units

    
    L_plates = regexBomber(re.compile('\s\D\d+\D{2}\s\d+'), L_plates_temp)
    L_plates = regexBomber(re.compile('\s\D{1}\s*\d+\s*\D{2}\s*\d+'), L_plates)
    L_plates = regexBomber(re.compile('\s\d+\s*\D{2}\s*\d+'), L_plates)
    L_plates = regexBomber(re.compile('\s\инв.№\s*\d+'), L_plates)
    L_plates = regexBomber(re.compile('\skz\s\D\d+\s\d+'), L_plates)
    L_plates = regexBomber(re.compile('\s\D{2}\s*\d+\s*\d+'), L_plates)
    # L_plates = regexBomber(re.compile('\s\d+'), L_plates)
    
    D = dict(zip(L_plates_temp, L_plates))
    pprint(D)
    
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
    
    # Repatch problem items after conversion to match with gen_PI
    D_replacers = {
        '637386ух': '6373ух'
    }
    
    for k, v in D_replacers.items():
        L_plates = [x.replace(k, v) for x in L_plates]
    
    L_PI = transform_plates(L_plates) 
    df = pd.DataFrame(zip(L_units, L_plates, L_PI, L_comments), columns=['Units', 'Plates', 'PI', 'Comments'])
    # pprint(df)

    
    
    # Form the comments column by matching
    # Merge comments column to gen db
    
    # Collecting crews and locs
    # L_crws, L_lcs = [], []
    # L_plates_unmatched = []
    # for i in L_plates:
    #     if cursor.execute(f"SELECT Units FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall():
    #         L_crws.append(cursor.execute(f"SELECT Crews FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall())
    #         L_lcs.append(cursor.execute(f"SELECT Fields FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall())
    #     else:
    #         L_plates_unmatched.append(i)
            
    
    # # Unpacking nested lists
    # L_crws = [', '.join(map(str, x)) for x in L_crws]
    # L_lcs = [', '.join(map(str, x)) for x in L_lcs]
    
    # df = pd.DataFrame(zip(L_crws, L_units, L_plates, L_lcs), columns=['Crews', 'Units', 'Plates', 'Locations'])

if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
