import json
import time
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


# Making connections to DBs
db = sqlite3.connect('data.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

# Pandas
pd.set_option('display.max_rows', None)

def main():
    # Pulling truck and trailer plates from db
    L_plates = cursor.execute("SELECT Plates_1 FROM brm_fixed").fetchall()
    L_plates2 = cursor.execute("SELECT Plates_2 FROM brm_fixed").fetchall()
    
    # Stringifying plates not to deal with NoneType problem
    L_plates = [str(x) for x in L_plates]
    L_plates2 = [str(x) for x in L_plates2]
    
     # Splitting merged truck plates
    L_plates = [''.join(re.sub('/', '', x)).strip() for x in L_plates ]
    L_plates = [''.join(re.sub('86', '86**', x)).strip() if x.endswith('86') and len(x) > 9 else x for x in L_plates ]
    L_plates = [x.split('**') if x != None else 'None' for x in L_plates]
    L_plates = list(itertools.chain.from_iterable(L_plates))
    # Cleaning truck plates off Nones
    L_plates = [x for x in L_plates if x != '']
    L_plates = [x for x in L_plates if x != 'None']
    
    # Splitting merged trailer plates
    L_plates2 = [''.join(re.sub('/', '', x)).strip() for x in L_plates2 ]
    L_plates2 = [''.join(re.sub('86', '86**', x)).strip() if x.endswith('86') and len(x) > 9 else x for x in L_plates2 ]
    L_plates2 = [x.split('**') if x != None else 'None' for x in L_plates2]
    L_plates2 = list(itertools.chain.from_iterable(L_plates2))
    # Cleaning trailer plates off Nones
    L_plates2 = [x for x in L_plates2 if x != '']
    L_plates2 = [x for x in L_plates2 if x != 'None']
   
    
    
    # Introducing del marker for items containing " (e.g. '3пл-5"кл-5,0"'), deleting items with del marker
    L_plates2 = [''.join(re.sub('\"', 'del', x)).strip() for x in L_plates2 if x != None]
    L_plates2 = ['' if 'del' in x else x for x in L_plates2]
    
    
    # Removing items that don't have numbers (i.e. not plates, e.g. 'автоцистерна'), or are the blank ones after the del removal
    pattern_D = re.compile(r'\d')
    L_plates2 = [x for x in L_plates2 if re.findall(pattern_D, str(x))]
   
    # merge truck and trailer plates
    L_plates3 = L_plates + L_plates2
    
 
    L_crws, L_unts, L_lcs = [], [], []
    for i in L_plates3:
        if cursor.execute(f"SELECT Crews FROM brm_fixed WHERE Plates_1 like '%{i}%'").fetchall():
            L_crws.append(cursor.execute(f"SELECT Crews FROM brm_fixed WHERE Plates_1 like '%{i}%'").fetchall())
            L_unts.append(cursor.execute(f"SELECT Units FROM brm_fixed WHERE Plates_1 like '%{i}%'").fetchall())
            L_lcs.append(cursor.execute(f"SELECT Locs FROM brm_fixed WHERE Plates_1 like '%{i}%'").fetchall())
        elif cursor.execute(f"SELECT Crews FROM brm_fixed WHERE Plates_2 like '%{i}%'").fetchall():
            L_crws.append(cursor.execute(f"SELECT Crews FROM brm_fixed WHERE Plates_2 like '%{i}%'").fetchall())
            L_unts.append(cursor.execute(f"SELECT Units FROM brm_fixed WHERE Plates_2 like '%{i}%'").fetchall())
            L_lcs.append(cursor.execute(f"SELECT Locs FROM brm_fixed WHERE Plates_2 like '%{i}%'").fetchall())

    
    L_crws_temp = []
    for i in L_crws:
        L_crws_temp.append(set(i))
    
    
    L_crws = [', '.join(map(str, x)) for x in L_crws_temp]
    L_unts = [', '.join(map(str, x)) for x in L_unts]
    L_lcs = [', '.join(map(str, x)) for x in L_lcs]
    
    
    # Turn plates into 123abc type
    def transform_plates(plates):
        plates = [re.sub('\s+', '', x) for x in plates]
        L_regions_long = [102, 126, 156, 158, 174, 186, 188, 196, 797]
        L_regions_short = ['01', '02', '03', '04', '05', '06', '07', '09']
        # Trucks
        for i in L_regions_long:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 9 else x for x in plates]
        # Trailers
        for i in L_regions_short:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 else x for x in plates]
        for i in range(10, 100):
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 else x for x in plates]
        
        # kzh Trailers
        plates = [x.removesuffix(str('07')).strip() if 'kzн' in str(x) and len(x) == 9 else x for x in plates]
        
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates

    
    
    L_plates_ind = transform_plates(L_plates3) 
    
    # Fixing PI to match to Omnicomm
    D_match_om = {
        '120530ghsm': '1205gh',
    }

    for k, v in D_match_om.items():
        L_plates_ind = [''.join(re.sub(k, v, x)).strip() if x != None and len(x) != 6 else x for x in L_plates_ind ]

    df = pd.DataFrame(zip(L_crws, L_unts, L_plates3, L_plates_ind, L_lcs), columns=['Crews', 'Units', 'Plates', 'Plate_index', 'Locs'])
    df = df.drop_duplicates(subset='Plate_index', keep="first")
    
    # df.sort_values('Crews')
    
    # Post df to DB
    cursor.execute("DROP TABLE IF EXISTS brm_final")
    df.to_sql(name='brm_final', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()



   
   
    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))

