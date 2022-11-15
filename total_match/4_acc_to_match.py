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
    df_match = pd.read_sql_query("SELECT * FROM gen_PI", cnx_match)

    # Listing accountance
    L_mols = df_acc['Mols'].tolist()
    L_acc_PI = df_acc['PI'].tolist()

    # Listing match
    L_gen_PI = df_match['PI_gen'].tolist()
    

     # Get matched plates
    def matcher(L_values):
        D = dict(zip(L_acc_PI, L_values))
        L = []
        L_mm = []
        for i in L_gen_PI:
            if i in D.keys():
                L.append(D.get(i))
                L_mm.append(D.get(i))
            else:
                L.append('-')
            
        return L
    
    
    L_mols_acc = matcher(L_mols)
    
    df_mols = pd.DataFrame(L_mols_acc, columns=['Mols'])
    df = df_match.join(df_mols, how = 'left')

    
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS acc_to_match")
    df.to_sql(name='acc_to_match', con=db_match, if_exists='replace', index=False)
    db_match.commit()
    db_match.close()

   
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))