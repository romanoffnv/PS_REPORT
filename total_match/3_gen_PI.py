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

# connection to match.db
db_match = sqlite3.connect('match.db')
db_match.row_factory = lambda cursor, row: row[0]
cursor = db_match.cursor()

cnx_match = sqlite3.connect('match.db')

def main():
    df = pd.read_sql_query("SELECT * FROM brm_to_om", cnx_match)
    
    L_PI_om = df['PI_om'].tolist()
    L_plates_ct = df['Plates_ct'].tolist()
    L_plates_brm = df['Plates_brm'].tolist()
    
    L_PI_gen = []
    for i, j, k in zip(L_PI_om, L_plates_ct, L_plates_brm):
        if i == '-':
            if j == '-':
                L_PI_gen.append(k)
            else:
                L_PI_gen.append(j)
        else:
            L_PI_gen.append(i)
    
    L_PI_gen = [re.sub('\s+', '', x) if len(x) < 12 and 'ДЭС' not in x else x for x in L_PI_gen]


    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions_long = [116, 126, 156, 158, 174, 186, 188, 196, 797]
        L_regions_short = ['01', '02', '03', '04', '05', '06', '07', '09']
        for i in L_regions_long:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 9 else x for x in plates]
        for i in L_regions_short:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        for i in range(10, 100):
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates
    
    
    L_PI_gen = transform_plates(L_PI_gen) 

    # Fixing diesel stations from dict
    D_om_diesels = json.load(open('D_om_diesels.json'))
    for k, v in D_om_diesels.items():
        L_PI_gen = [''.join(x.replace(k, v)).strip() for x in L_PI_gen]
    
    df2 = pd.DataFrame(L_PI_gen, columns=['PI_gen'])
    df = df.join(df2, how = 'left')
    pprint(df)
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS gen_PI")
    df.to_sql(name='gen_PI', con=db_match, if_exists='replace', index=False)
    db_match.commit()
    db_match.close()

if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))