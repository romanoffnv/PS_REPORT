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
db = sqlite3.connect('data.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
cnx = sqlite3.connect('data.db')


def main():
    # Pulling lists (working, common for merging) from db
    L_units = cursor.execute("SELECT Units FROM om_parse_status").fetchall()
    
    
    # Db into df
    
    df1 = pd.read_sql_query("SELECT * FROM om_parse_status", cnx)
    
   

    # Fishing out plates by regex from long sentences
    L_plates_temp = []
    def plate_fisher(regex, L_units):
        for i in L_units:
            if 'ДЭС' in i:
                L_plates_temp.append(i)
            else:
                if re.findall(regex, str(i)):
                    L_plates_temp.append(''.join(re.findall(regex, str(i))))
                else:
                    L_plates_temp.append(i)
                # print(i)
    
        L_units = [str(x).strip() for x in L_plates_temp]
        L_plates_temp.clear() 
            
        return L_units

    L_regex = [
            '\s\D{2}\s*\d{2}\s*\d{2}\s*\d+', #ВВ  4553 86, # АН 78 96 82, ВВ  4553 86
            '\s\D\d+\D{2}\s\d+', #Mitsubishi L200 а581тв 156
            '\s\D\s*\d{3}\s*\D{2}\s*\d+', #Е 898 СВ 186, У 039 ВК186
            '\(\d+\s*\D+\s*\d+\)', #(7250ах86)
            '\s\D\s\d{4}\s+\d+', #H 0762  07
            '\s\d{4}\s\D{2}\s+\d+', #7713 НХ 77
            '\s\d{4}\D{2}\s\d+', #3824ат 86
            '\s\d{4}\D{2}\d+', #гос.№ 6241АУ86
            '\s\D{2}\-\D+\-\d+', #CT-DV-141, CT-CTU-1000
            '\s\D{3}\-\d+', #HFU-2000
            '\№\s\d+', #№ 0079
            '\s\D\s*\d{3}\s*\D{2}\s*\d+', #runs again to choose bw paranthesis and outside par Е 898 СВ 186
            '\s\d+\D{2}\s*\d+', #0775уу 86
            '\skz\s\D\d+\s\d+', #kz н0762 07
            '\s\D\d+\s\d+\.\d+', #С008 02.2014
            

            '\s\d{5}', #Блендер контейнерный 05566
        ]

    L_plates = plate_fisher(re.compile(L_regex[0]), L_units)

    for regex in L_regex:
        L_plates = plate_fisher(re.compile(regex), L_plates)
        
    L_cleaners = ['№', '(', ')', '.']
    for i in L_cleaners:
        L_plates = [x.replace(i, '') for x in L_plates]

    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions_long = [126, 156, 158, 174, 186, 188, 196, 797]
        L_regions_short = ['01', '02', '03', '04', '05', '06', '07', '09']
        for i in L_regions_long:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 9 else x for x in plates]
        for i in L_regions_short:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or len(x) == 9 else x for x in plates]
        for i in range(10, 100):
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or len(x) == 9 else x for x in plates]
        
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates
    
    
    L_plates = [re.sub('\s+', '', x) for x in L_plates]
    L_plates_ind = transform_plates(L_plates) 
    
    for i in L_plates_ind:
        if len(i) != 6:
            print(i)
    
    
    df2 = pd.DataFrame(zip(L_plates, L_plates_ind), columns = ['Plates', 'PI'])
    
    # Merge dfs by columns 
    df = df1.join(df2, how = 'left')
    print(df)

    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS om_parse_plates")
    df.to_sql(name='om_parse_plates', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))