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
db_acc = sqlite3.connect('accountance.db')
db_acc.row_factory = lambda cursor, row: row[0]
cursor_acc = db_acc.cursor()

# connection to match.db
db_match = sqlite3.connect('match.db')
db_match.row_factory = lambda cursor, row: row[0]
cursor = db_match.cursor()

cnx_acc = sqlite3.connect('accountance.db')
cnx_match = sqlite3.connect('match.db')

def main():
    df_acc = pd.read_sql_query("SELECT * FROM accountance_2", cnx_acc)
    df_match = pd.read_sql_query("SELECT * FROM om_cits_brm", cnx_match)

    # Destructuring df_acc
    def cits_destructurer():
        L0 = df_acc['Mols'].tolist()
        L1 = df_acc['PI'].tolist()
        return L0, L1
    
    L_all_acc = cits_destructurer()
    
    # Destructuring match df 
    def match_destructurer():
        # Omnicomm
        L0 = df_match['Groups_om'].tolist()
        L1 = df_match['Units_om'].tolist()
        L2 = df_match['id_om'].tolist()
        L3 = df_match['Status_om'].tolist()
        L4 = df_match['Plates_om'].tolist()
        L5 = df_match['PI_om'].tolist()
        L6 = df_match['Locs_om'].tolist()
        # Cits
        L7 = df_match['Crews_ct'].tolist()
        L8 = df_match['Units_ct'].tolist()
        L9 = df_match['Plates_ct'].tolist()
        L10 = df_match['Locs_ct'].tolist()
        # Brm
        L11 = df_match['Crews_brm'].tolist()
        L12 = df_match['Units_brm'].tolist()
        L13 = df_match['Plates_brm'].tolist()
        L14 = df_match['Locs_brm'].tolist()
        return L0, L1, L2, L3, L4, L5, L6, L7, L8, L9, L10, L11, L12, L13, L14
    
    L_all_match = match_destructurer()
    
    # Get matched plates
    def matcher(L_values):
        D = dict(zip(L_all_acc[1], L_values))
        L = []
        L_mm = []
        for i in L_all_match[5]:
            if i in D.keys():
                L.append(D.get(i))
                L_mm.append(D.get(i))
            else:
                L.append('-')
            
        return L
        # return L_mm
    
    L_mols_acc = matcher(L_all_acc[0])
    df_mols = pd.DataFrame(L_mols_acc, columns=['Mols'])
    df = df_match.join(df_mols, how = 'left')
    
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS om_ct_fr_acc")
    df.to_sql(name='om_ct_fr_acc', con=db_match, if_exists='replace', index=False)
    db_match.commit()
    db_match.close()

    writer = pd.ExcelWriter('DB.xlsx', engine='xlsxwriter')
    # Write each dataframe to a different worksheet.
    # data.index += 1
    df.to_excel(writer, index = True, header=True)
    writer.save()
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))