from asyncio.windows_events import NULL
import json
from numpy import NaN
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
db = sqlite3.connect('brm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

# Pandas
pd.set_option('display.max_rows', None)

def main():
    # Pulling Units_Locs_Raw.db into lists
    L_crews = cursor.execute("SELECT Crews FROM Units_Locs_Raw").fetchall()
    L_units = cursor.execute("SELECT Units FROM Units_Locs_Raw").fetchall()
    L_plates = cursor.execute("SELECT Plates_1 FROM Units_Locs_Raw").fetchall()
    L_plates2 = cursor.execute("SELECT Plates_2 FROM Units_Locs_Raw").fetchall()
    L_locs = cursor.execute("SELECT Locs FROM Units_Locs_Raw").fetchall()
    
    
    
    # Fixing truck plates 
    # Crutch
    D_plates_fix = {
        '618':'в618нт186',
        '992':'в992сх 186',
        'Н 0762':'kz н0762 07',
        '751.0':' В751КА 186',
        '865.0':' в865мт186',
        '730.0':' в730ка186',
        '717.0':' в717ка186',
        '368.0':' в368мр186',
        '743.0':' в368мр186',
    }

    for k, v in D_plates_fix.items():
        L_plates = [''.join(re.sub(k, v, x)).strip() if x != None and len(x) < 7 else x for x in L_plates ]
    
    
    # Cleaning trailers
    # Crutch
    D_plates2_fix = {
        '403 Аренда с ЮТС': '',   
        'Насос ВД': '',   
        'АДПМ': '',   
        'АЦН-17': '',   
        'Площадка': '',   
        'Блендер': '',   
    }

    for k, v in D_plates2_fix.items():
        L_plates2 = [''.join(re.sub(k, v, x)).strip() if x != None else x for x in L_plates2 ]

    
    # Bringing plates into A123MT186 format
    # Removing spaces in plates 
    L_plates = [''.join(re.sub('\s+', '', x)) if x != None else x for x in L_plates ]
    L_plates2 = [''.join(re.sub('\s+', '', x)) if x != None else x for x in L_plates2]
 
    
    # Building Data Frame
    df = pd.DataFrame(zip(L_crews, L_units, L_plates, L_plates2, L_locs), columns=['Crews', 'Units', 'Plates_1', 'Plates_2', 'Locs'])
    # Drop rows with nulls in all columns other than 'Crews'
    df = df.dropna( how='any', subset=['Plates_1', 'Plates_2', 'Locs'], thresh=1)
    
    # Posting df to DB
    cursor.execute("DROP TABLE IF EXISTS Units_Locs_Fixed")
    df.to_sql(name='Units_Locs_Fixed', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
if __name__ == '__main__':
    main()