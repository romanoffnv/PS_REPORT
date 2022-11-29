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
    df_mark = pd.read_excel("mark.xlsx")
    print(df_mark)
    # Get Бригада, СПТ, Гос.номер, Данные по закреплению МОЛ (бухгалтерия), Данные закрепления службы ГНКТ / ГРП (Водители), Примечание
    # Destructuring df_drv
    def mark_destructurer():
        # Turn plates into 123abc type
        def transform_plates(plates):
            plates = [str(x) for x in plates]
            plates = [re.sub('\s+', '', x) for x in plates]
            L_regions_long = [126, 156, 158, 174, 186, 188, 196, 797]
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
        
        L0 = df_mark['Crews'].tolist()
        L1 = df_mark['Units'].tolist()
        L2 = df_mark['Plates'].tolist()
        L3 = df_mark['Mols'].tolist()
        L4 = df_mark['Drivers'].tolist()
        L5 = df_mark['Acc_comments'].tolist()
        L6 = transform_plates(L2)
        return L0, L1, L2, L3, L4, L5, L6
    
    L_all_mark = mark_destructurer()
    
    df_match = pd.read_sql_query("SELECT * FROM cunt_to_match", cnx_match)
     # Destructuring match df 
    L_PI = df_match['PI_gen'].tolist()
    
    # # Get matched plates
    # def matcher(L_values):
    #     D = dict(zip(L_all_mark[5], L_values))
    #     L = []
    #     L_mm = []
    #     for i in L_PI:
    #         if i in D.keys():
    #             L.append(D.get(i))
    #             L_mm.append(D.get(i))
    #         else:
    #             L.append('-')
            
    #     return L
  
    
   
            
    # L_crews = matcher(L_all_drv[0])
    # df_drivers = pd.DataFrame(L_drivers, columns=['Drivers_frac'])
    # df = df_match.join(df_drivers, how = 'left')

    #  Get unmatched plates
    def dismatcher(L_values):
        D = dict(zip(L_all_mark[6], L_values))
        
        L = []
        for k, v in D.items():
            if k not in L_PI:
                L.append(v)
                
                
        return L
    
    
    L_crews = dismatcher(L_all_mark[0])
    L_units = dismatcher(L_all_mark[1])
    L_plates = dismatcher(L_all_mark[2])
    L_mols = dismatcher(L_all_mark[3])
    L_drivers = dismatcher(L_all_mark[4])
    L_comments = dismatcher(L_all_mark[5])
    
    df = pd.DataFrame(zip(L_crews, L_units, L_plates, L_mols, L_drivers, L_comments))
    pprint(df)
    
    writer = pd.ExcelWriter('mark2.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index = True, header=True)
    writer.save()

if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))