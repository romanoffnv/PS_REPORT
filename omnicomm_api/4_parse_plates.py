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
db = sqlite3.connect('omnicomm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
cnx = sqlite3.connect('omnicomm.db')


def main():
    # Pulling lists (working, common for merging) from db
    L_units = cursor.execute("SELECT Units FROM parse_status").fetchall()
    
    
    # Db into df
    
    df1 = pd.read_sql_query("SELECT * FROM parse_status", cnx)
    
    # Fishing out plates by regex from long sentences
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
                    for x in L_units]

    
    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions = [186, 86, 797, 116, '02', '07',89, 82, 78, 54, 77, 126, 188, 88, 174, 74, 158, 196, 156, 56, 76, 23]
        
        for i in L_regions:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) > 7 else x for x in plates]
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates
    

    L_plates_ind = transform_plates(L_plates) 
    
    # Fixing diesel stations from dict
    D_om_diesels = json.load(open('D_om_diesels.json'))
    for k, v in D_om_diesels.items():
        L_plates_ind = [''.join(x.replace(k, v)).strip() for x in L_plates_ind]
    
    
    df2 = pd.DataFrame(zip(L_plates, L_plates_ind), columns = ['Plates', 'PI'])
    
    # Merge dfs by columns 
    df = df1.join(df2, how = 'left')

    print(df)
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS parse_plates")
    df.to_sql(name='parse_plates', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))