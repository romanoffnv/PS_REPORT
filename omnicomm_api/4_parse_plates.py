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

    
    L_plates = plate_fisher(re.compile('\s\D\d+\D{2}\s\d+'), L_units)
    L_plates = plate_fisher(re.compile('\(\d{4}\D{2}\d{2}\)'), L_plates) #(7250ах86)
    L_plates = plate_fisher(re.compile('\D{2}\d{4}\s\d{2}'), L_plates) #s/n № 1000004402
    L_plates = plate_fisher(re.compile('s/n\s\№\s\d+'), L_plates) #НВД №1 ВВ8684 86
    L_plates = plate_fisher(re.compile('\s\D{1}\s*\d+\s*\D{2}\s*\d+'), L_plates)
    L_plates = plate_fisher(re.compile('\s\d+\s*\D{2}\s*\d+'), L_plates)
    L_plates = plate_fisher(re.compile('\s\инв.№\s*\d+'), L_plates)
    L_plates = plate_fisher(re.compile('\skz\s\D\d+\s\d+'), L_plates)
    L_plates = plate_fisher(re.compile('\s\D{2}\s*\d{4}\s*\d+'), L_plates)
    L_plates = plate_fisher(re.compile('\d{4}\D{2}\s\d+'), L_plates) #(8804ах 86)

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
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        for i in range(10, 100):
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates
    
    
    L_plates = [re.sub('\s+', '', x) for x in L_plates]
    L_plates_ind = transform_plates(L_plates) 
    
    
    
    df2 = pd.DataFrame(zip(L_plates, L_plates_ind), columns = ['Plates', 'PI'])
    
    # Merge dfs by columns 
    df = df1.join(df2, how = 'left')

    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS parse_plates")
    df.to_sql(name='parse_plates', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))