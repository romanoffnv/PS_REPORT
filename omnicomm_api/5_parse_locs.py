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
db = sqlite3.connect('data.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
cnx = sqlite3.connect('data.db')

def main():
    df = pd.read_excel('geozonesreport.xlsx')
    df = df[9:]
    df = df.drop_duplicates(subset='Unnamed: 0', keep="first")
    L_units = df['Unnamed: 0'].tolist()
    L_locs = df['Unnamed: 1'].tolist()
    df = pd.DataFrame(zip(L_units, L_locs), columns=['Units', 'Locs'])
    df1 = pd.read_sql_query("SELECT * FROM om_parse_plates", cnx)

    cursor.execute("DROP TABLE IF EXISTS om_parse_locs")
    df.to_sql(name='om_parse_locs', con=db, if_exists='replace', index=False)
    db.commit()
    
    L_units = cursor.execute(f"SELECT Units FROM om_parse_plates").fetchall()
   
    L_locs_temp = []
    for i in L_units:
        if cursor.execute(f"SELECT Units FROM om_parse_locs WHERE Units like '%{i}%'").fetchall():
            L_locs_temp.append(cursor.execute(f"SELECT Locs FROM om_parse_locs WHERE Units like '%{i}%'").fetchall())
        else:
            L_locs_temp.append('-')

    # Unpacking nested lists
    L_locs = [', '.join(map(str, x)) if isinstance(x, list) else x for x in L_locs_temp]
    
    df2 = pd.DataFrame(L_locs, columns=['Locs'])
    df = df1.join(df2, how = 'left')
    pprint(df)

    cursor.execute("DROP TABLE IF EXISTS om_final")
    df.to_sql(name='om_final', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))