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
db_match = sqlite3.connect('match.db')
db_match.row_factory = lambda cursor, row: row[0]
cursor_match = db_match.cursor()

db_acc = sqlite3.connect('accountance.db')
db_acc.row_factory = lambda cursor, row: row[0]
cursor_acc = db_acc.cursor()

cnx_match = sqlite3.connect('match.db')
cnx_acc = sqlite3.connect('accountance.db')
df_match = pd.read_sql_query("SELECT * FROM acc_to_match", cnx_match)
df_acc = pd.read_sql_query("SELECT * FROM accountance_3", cnx_acc)



# Pandas
pd.set_option('display.max_rows', None)
pd.options.display.width = 1200
pd.options.display.max_colwidth = 30
pd.options.display.max_columns = 30


def main():
    L_gen_PI = df_match['PI_gen'].tolist()
    L_acc_PI = df_acc['PI'].tolist()
    L_acc_com = df_acc['Comments'].tolist()

    # pprint(L_gen_PI)
    # pprint(len(L_gen_PI))
    # pprint(len(L_acc_PI))
    # pprint(len(L_acc_com))

    D = dict(zip(L_acc_PI, L_acc_com))
    pprint(D)
    L_acc_comments = []
    for i in L_gen_PI:
        L_acc_comments.append(D.get(i))
    
    df = pd.DataFrame(zip(L_acc_comments), columns=['Acc_comments'])
    df = df_match.join(df, how = 'left')
    
    
    # Posting df to DB
    print('Posting df to DB')
    cursor_match.execute("DROP TABLE IF EXISTS acc_com_to_match")
    df.to_sql(name='acc_com_to_match', con=db_match, if_exists='replace', index=False)
    db_match.commit()
    db_match.close()

if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))