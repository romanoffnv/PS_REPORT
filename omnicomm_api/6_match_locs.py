from collections import defaultdict
import time
import xlsxwriter
from multiprocessing.sharedctypes import Value
import pandas as pd
import os
import sqlite3
import re
from pprint import pprint
from win32com.client.gencache import EnsureDispatch
import win32com
print(win32com.__gen_path__)



# Making connections to db
db = sqlite3.connect('omnicomm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
cnx = sqlite3.connect('omnicomm.db')

def main():
    df1 = pd.read_sql_query("SELECT * FROM parse_plates", cnx)
    
    L_units = cursor.execute(f"SELECT Units FROM parse_plates").fetchall()
   
    L_locs_temp = []
    for i in L_units:
        if cursor.execute(f"SELECT Units FROM parse_locs WHERE Units like '%{i}%'").fetchall():
            L_locs_temp.append(cursor.execute(f"SELECT Locations FROM parse_locs WHERE Units like '%{i}%'").fetchall())
        else:
            L_locs_temp.append('нет данных')

     # Unpacking nested lists
    L_locs = [', '.join(map(str, x)) if isinstance(x, list) else x for x in L_locs_temp]
    df2 = pd.DataFrame(L_locs, columns=['Locations'])
    
    # Merge dfs by columns 
    df = df1.join(df2, how = 'left')
    pprint(df)
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS Final_om")
    df.to_sql(name='Final_om', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))