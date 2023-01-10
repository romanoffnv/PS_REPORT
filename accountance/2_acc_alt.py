import time
import xlsxwriter
from win32com.client.gencache import EnsureDispatch
import os
import re
from pprint import pprint
import pandas as pd
import numpy as np
from functools import reduce
import itertools
import sqlite3
import win32com
print(win32com.__gen_path__)


# Pandas
pd.set_option('display.max_rows', None)

# db connections
db = sqlite3.connect('data.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
cnx = sqlite3.connect('data.db')

pd.set_option('display.max_rows', None)

def main():
    # Get Unit col from accountance_1.db as list
    L_units = cursor.execute("SELECT Units FROM accountance_1").fetchall()
    

    # Split strings in Unit list by the keywords and get as L_plates list
    def splitter(split, L):
      L = [str(x) for x in L]  
      L = [x.split(split) for x in L]
      L = list(itertools.chain.from_iterable(L))

      return L

    L_keywords = ['г/н', 
                'Truck', 
                '43118', 
                'г.н.', 'гн', '(', 'г/р', ';', '43118', 'Гос.№', ',', 'зав.', 
                # 'мод.', 
                'зав', 
                ')', 'ст ', 'Г/н', 
                'АЦН'
                ]
    

    # Getting pilot splitted list of plates by sending first keyword from the list and L_units
    L_plates = splitter(L_keywords[0], L_units)
    # Keep on splitting by sending keywords and pilot plates list
    for i in L_keywords:
        L_plates = splitter(str(i), L_plates)
    
    # Filter out strings that:
    # have more or less letters than in a real plate except for diesel stations
    L_plates = [''.join(x).strip() for x in L_plates if 'изель' in x or (sum(map(str.isalpha, x)) < 4 and sum(map(str.isalpha, x)) > 1)]
    # are shorter than 6 characters
    L_plates = [x for x in L_plates if len(x) > 6]
    # have one of the keys
    L_keys = ['г.в.', 'л.с.', 'VIN', 'НД', 'Квт', 'час', 'ит', 'Gr', 'dpi', 
              'ф/з', 'FHD', 'до', '.', '-', '=', 'Ш', 'Mb', 'лот', 'HI', 'г', 'кВт', 'ST', 'TTR']
    for i in L_keys:
        L_plates = [x for x in L_plates if i not in x]
    # regexed as follows
    L_reg = ['\ЕМС\s\d{3}', #ЕМС 600
             '\d+\х\d.*', #8000х2500 мм, 6000х2450х2600
             ]
    for i in L_reg:
        L_plates = [re.sub(i, '', x) for x in L_plates]

    L_plates = set(L_plates)
    L_plates = list(L_plates)
    L_plates = [x for x in L_plates if x != '']
    # Fishing out serial nums
    L_plates = [''.join(re.findall('\№.*', x)) if re.findall('\№.*', x) else x for x in L_plates]

    
    
    # Fishing mols and units by L_plates iterable from accountance_1 db
    L_mols, L_units_temp, L_plates_unmatched = [], [], []
    for i in L_plates:
        if cursor.execute(f"SELECT Units FROM accountance_1 WHERE Units like '%{i}%'").fetchall():
            L_mols.append(cursor.execute(f"SELECT Mols FROM accountance_1 WHERE Units like '%{i}%'").fetchall())
            L_units_temp.append(cursor.execute(f"SELECT Units FROM accountance_1 WHERE Units like '%{i}%'").fetchall())
        else:
            
            L_plates_unmatched.append(i)

    
    
    L_mols = [', '.join(map(str, x)) for x in L_mols]
    L_units_temp = [', '.join(map(str, x)) for x in L_units_temp]
    L_units = [x for x in L_units_temp]
   
    df1 = pd.read_sql_query("SELECT * FROM accountance_1", cnx)
    df2 = pd.DataFrame(zip(L_mols, L_units, L_plates), columns=['Mols', 'Units', 'Plates'])
    df = pd.merge(df1, df2, how="outer")
    # df = df.drop_duplicates(subset='Units', keep="last")
    pprint(df)
         
    
    #  Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS accountance_2")
    df.to_sql(name='accountance_2', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
    

    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
