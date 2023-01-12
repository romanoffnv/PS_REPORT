import time
import collections
import xlsxwriter
from win32com.client.gencache import EnsureDispatch
import sys
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
path = '/Users/roman/OneDrive/Рабочий стол/SANDBOX/PS_REPORT'
file = os.path.join(path, 'data.db')
db = sqlite3.connect(file)
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
cnx = sqlite3.connect(file)

pd.set_option('display.max_rows', None)

def main():
    # Get Unit col from accountance_2.db as list
    L_units = cursor.execute("SELECT Units FROM accountance_2").fetchall()
    
    # ******************************* REUSABLE FUNCTIONS ****************************************
    
    # STRING SPLITTER
    # Landing splitting marks
    def mark_landing(s, L_units):
        L = [re.sub(s, '*split*', x) for x in L_units]
        return L
    
    # Split strings in Unit list by the keywords and get as L_plates list
    def splitter(L):
        L = [x.split('*split*') for x in L]
        L = list(itertools.chain.from_iterable(L))
    

        return L

    # ******************************* FUNCTION CALL PARAMS ****************************************
    
    # Sending keywords to mark_landing func
    L_keywords = ['г/н', 'Truck', '43118', 'г.н.', 'гн', '\(', 'г/р', ';', ',', 'Гос.№', 'зав.',
                'зав', '№', '\)', 'ст ', 'Г/н', 'АЦН', 'электростанция', 'дизельный']

    for i in L_keywords:
        L_units = mark_landing(str(i), L_units)
    
    # Getting splitted list by marks as L_plates
    L_plates  = splitter(L_units)
   
    # ********************************************************************************************
    
    def plate_validator(L_plates):
        L_literals = []
        L_numeric = []
        for i in L_plates:
            L_literals.append(sum(map(str.isalpha, i)) < 4 and sum(map(str.isalpha, i)) > 1)
            L_numeric.append(sum(map(str.isnumeric, i)) < 7 and sum(map(str.isnumeric, i)) > 5)
            
        df = pd.DataFrame(zip(L_plates, L_literals, L_numeric))
        return df
        # have more or less letters than in a real plate 
        # L_plates = [''.join(x).strip() for x in L_plates if (sum(map(str.isalpha, x)) < 4 and sum(map(str.isalpha, x)) > 1)]
    
    run = plate_validator(L_plates)
    pprint(run)
    # are shorter than 6 characters
    # L_plates = [x for x in L_plates if len(x) > 6]
    
    # # have one of the keys
    # L_keys = ['г.в.', 'л.с.', 'VIN', 'НД', 'Квт', 'кВт', 'час', 'ит', 'Gr', 'dpi', 
    #           'ф/з', 'FHD', 'до', '.', '-', '=', 'Ш', 'Mb', 'лот', 'HI', 'г', 'кВт', 'ST', 'TTR', 's/n', 'сер№', 'S/N']
    # for i in L_keys:
    #     L_plates = [x for x in L_plates if i not in x]
    # # regexed as follows
    # L_reg = ['\ЕМС\s\d{3}', #ЕМС 600
    #          '\d+\х\d.*', #8000х2500 мм, 6000х2450х2600
    #          '\d{2}\.\d\s\мм', #50.8 мм
    #         ]
    # for i in L_reg:
    #     L_plates = [re.sub(i, '', x) for x in L_plates]

    # # Removing empty strings
    # L_plates = [str(x) for x in L_plates if x != ''] 
    
    # # Fishing mols and units by L_plates iterable from accountance_2 db
    # L_mols, L_units_temp, L_plates_unmatched = [], [], []
    # for i in L_plates:
    #     if cursor.execute(f"SELECT Units FROM accountance_2 WHERE Units like '%{i}%'").fetchall():
    #         L_mols.append(cursor.execute(f"SELECT Mols FROM accountance_2 WHERE Units like '%{i}%'").fetchall())
    #         L_units_temp.append(cursor.execute(f"SELECT Units FROM accountance_2 WHERE Units like '%{i}%'").fetchall())
    #     else:
            
    #         L_plates_unmatched.append(i)

    # L_mols = [', '.join(map(str, x)) for x in L_mols]
    # L_units_temp = [', '.join(map(str, x)) for x in L_units_temp]
    # L_units = [x for x in L_units_temp]
   
    # # Slicing dubbed Mols by comma combined while searching db like {i}
    # L_mols_temp = []
    # for i in L_mols:
    #     ind = str(i).find(',')
    #     if ind != -1:
    #         L_mols_temp.append(i[:ind])
    #     else:
    #         L_mols_temp.append(i)
    # L_mols = [x for x in L_mols_temp]
    
    # df1 = pd.read_sql_query("SELECT * FROM accountance_1", cnx)
    # df2 = pd.DataFrame(zip(L_mols, L_units, L_plates), columns=['Mols', 'Units', 'Plates'])
    # df = pd.merge(df1, df2, how="outer")
    # df = df.drop_duplicates(subset='Units', keep="last")
    # pprint(df)

    # #  Posting df to DB
    # print('Posting df to DB')
    # cursor.execute("DROP TABLE IF EXISTS accountance_3")
    # df.to_sql(name='accountance_3', con=db, if_exists='replace', index=False)
    # db.commit()
    # db.close()
    
    

    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
