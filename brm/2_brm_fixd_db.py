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
db = sqlite3.connect('brm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

# Pandas
pd.set_option('display.max_rows', None)

def main():
    # Pulling Units_Locs_Raw.db into lists
    L_crews = cursor.execute("SELECT Crews FROM Units_Locs_Raw").fetchall()
    L_plates = cursor.execute("SELECT Units_1 FROM Units_Locs_Raw").fetchall()
    L_plates2 = cursor.execute("SELECT Units_2 FROM Units_Locs_Raw").fetchall()
    L_locs = cursor.execute("SELECT Locs FROM Units_Locs_Raw").fetchall()
    
    # Fixing truck plates that don't have literals
    # Crutch
    D_plates_fix = {
        '618':'в618нт186',
        '992':'в992сх 186',
        'Н 0762':'kz н0762 07',
        '751.0':' В751КА 186',
        '865.0':' в865мт186',
        '730.0':' в730ка186',
    }

    for k, v in D_plates_fix.items():
        L_plates = [''.join(re.sub(k, v, x)).strip() if x != None and len(x) < 7 else x for x in L_plates ]
    
    # Could be smth for fixing trailer plates too, as D_plates2_fix
    
    # Bringing plates into A123MT186 format
    # Removing spaces in plates 
    L_plates = [''.join(re.sub('\s+', '', x)) if x != None else x for x in L_plates ]
    L_plates2 = [''.join(re.sub('\s+', '', x)) if x != None else x for x in L_plates2]

    # Stripping locations
    L_locs = [str(x).strip() for x in L_locs if x != None]
    
    # Building Data Frame
    df = pd.DataFrame(zip(L_crews, L_plates, L_plates2, L_locs), columns=['Crews', 'Units_1', 'Units_2', 'Locs'])
    
    # Posting df to DB
    cursor.execute("DROP TABLE IF EXISTS Units_Locs_Raw")
    df.to_sql(name='Units_Locs_Raw', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
if __name__ == '__main__':
    main()