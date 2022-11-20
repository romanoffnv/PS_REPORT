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


# Making connections to DBs
# connection to brm.db
db_drv = sqlite3.connect('drivers.db')
db_drv.row_factory = lambda cursor, row: row[0]
cursor_drv = db_drv.cursor()

# connection to match.db
db_match = sqlite3.connect('match.db')
db_match.row_factory = lambda cursor, row: row[0]
cursor = db_match.cursor()

cnx_drv = sqlite3.connect('drivers.db')
cnx_match = sqlite3.connect('match.db')

def main():
    df_drv = pd.read_sql_query("SELECT * FROM frac_drivers", cnx_drv)
    df_match = pd.read_sql_query("SELECT * FROM arby_to_match", cnx_match)

    # Destructuring df_drv
    def drv_destructurer():
        L0 = df_drv['Drivers'].tolist()
        L1 = df_drv['PI'].tolist()
        return L0, L1
    
    L_all_drv = drv_destructurer()
    
    # Destructuring match df 
    L_PI = df_match['PI_gen'].tolist()
    
    # Get matched plates
    def matcher(L_values):
        D = dict(zip(L_all_drv[1], L_values))
        L = []
        L_mm = []
        for i in L_PI:
            if i in D.keys():
                L.append(D.get(i))
                L_mm.append(D.get(i))
            else:
                L.append('-')
            
        return L
            
    L_drivers = matcher(L_all_drv[0])
    df_drivers = pd.DataFrame(L_drivers, columns=['Drivers_frac'])
    df = df_match.join(df_drivers, how = 'left')
   
    cols = df.columns.tolist()
    pprint(cols)
    cols = ['Groups_om',
            'Units_om',
            'id_om',
            'Status_om',
            'Plates_om',
            'PI_om',
            'Locs_om',
            'Crews_ct',
            'Units_ct',
            'Plates_ct',
            'Drivers_ct',
            'Locs_ct',
            'Crews_brm',
            'Units_brm',
            'Plates_brm',
            'Drivers_frac',
            'Locs_brm',
            'PI_gen',
            'Mols',
            'Acc_comments',
            ]
    df = df[cols]
    pprint(cols)
    
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS cunt_to_match")
    df.to_sql(name='cunt_to_match', con=db_match, if_exists='replace', index=False)
    db_match.commit()
    db_match.close()

if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))