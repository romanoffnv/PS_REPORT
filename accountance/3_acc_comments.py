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
    # pprint(L_plates_temp)
    # pprint(len(L_plates_temp))
    
    # Fish out plates with regex from newly acquired items
    # Slicing vehicles column(list) into plate, index, literal cols
    plates1 = re.compile("[А-Яа-я]*\d+[А-Яа-я]{2}\s*\d+")
    plates2 = re.compile("[А-Яа-я]{2}\d+\s\d+")
    plates3 = re.compile("\w{2}\s\D\d+\s\d{2}")
    plates4 = re.compile("\W\d+\s\d+\.\d+")
    plates5 = re.compile("\ДЭС.*")
    plates6 = re.compile("\дэс.*")
    plates7 = re.compile("\D{2}\s\d+\s\d+")
    
    # Crutches
    plates8 = re.compile("GH120530SM")
    plates9 = re.compile("ПГУ-ОЗРД  113")
    # Jereh Маз насос инв №
    plates10 = re.compile("091217")
    
    
    # Derivating plates from vehicles
    L_plates = [''.join(re.findall(plates1, x)) or 
                     ''.join(re.findall(plates2, x)) or 
                     ''.join(re.findall(plates3, x)) or 
                     ''.join(re.findall(plates4, x)) or
                     ''.join(re.findall(plates5, x)) or
                     ''.join(re.findall(plates6, x)) or
                     ''.join(re.findall(plates7, x)) or 
                     ''.join(re.findall(plates8, x)) or
                     ''.join(re.findall(plates9, x)) or
                     ''.join(re.findall(plates10, x)) 
                     if re.findall(plates1, x) else x for x in L_plates_temp]

   
    # Patch items that can't be regexed with dict
    D_replacers = {
        'Установка Колтюбинговая МЗКТ 6373 УХ 86': '6373 УХ 86'
    }
    
    for k, v in D_replacers.items():
        L_plates = [x.replace(k, v) for x in L_plates]

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
    df = pd.DataFrame(zip(L_units, L_plates, L_PI, L_comments))
    pprint(df)

    
    
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
