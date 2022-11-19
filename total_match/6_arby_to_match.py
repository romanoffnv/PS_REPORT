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
db_arb = sqlite3.connect('arby.db')
db_arb.row_factory = lambda cursor, row: row[0]
cursor_arb = db_arb.cursor()

# connection to match.db
db_match = sqlite3.connect('match.db')
db_match.row_factory = lambda cursor, row: row[0]
cursor = db_match.cursor()

cnx_arb = sqlite3.connect('arby.db')
cnx_match = sqlite3.connect('match.db')

def main():
    df_arb = pd.read_sql_query("SELECT * FROM ct_drivers", cnx_arb)
    df_match = pd.read_sql_query("SELECT * FROM acc_com_to_match", cnx_match)

    # Destructuring df_arb
    def cits_destructurer():
        L0 = df_arb['Drivers'].tolist()
        L1 = df_arb['PI'].tolist()
        return L0, L1
    
    L_all_arb = cits_destructurer()
    
    # Destructuring match df 
    L_PI = df_match['PI_om'].tolist()
    
    # Get matched plates
    def matcher(L_values):
        D = dict(zip(L_all_arb[1], L_values))
        L = []
        L_mm = []
        for i in L_PI:
            if i in D.keys():
                L.append(D.get(i))
                L_mm.append(D.get(i))
            else:
                L.append('-')
            
        return L
            
    L_drivers = matcher(L_all_arb[0])
    df_drivers = pd.DataFrame(L_drivers, columns=['Drivers_ct'])
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
            'Locs_brm',
            'Loctions_om',
            'PI_gen',
            'Mols',
            'Acc_comments',
            ]
    df = df[cols]
    pprint(cols)
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS arby_to_match")
    df.to_sql(name='arby_to_match', con=db_match, if_exists='replace', index=False)
    db_match.commit()
    db_match.close()
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))