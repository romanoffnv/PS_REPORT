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
db = sqlite3.connect('brm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

db_om = sqlite3.connect('omnicomm.db')
db_om.row_factory = lambda cursor, row: row[0]
cursor_om = db_om.cursor()

# Pandas
pd.set_option('display.max_rows', None)

def main():
    # Pulling truck and trailer plates from db
    L_plates = cursor.execute("SELECT Plates_1 FROM Units_Locs_Fixed").fetchall()
    L_plates2 = cursor.execute("SELECT Plates_2 FROM Units_Locs_Fixed").fetchall()
    
    # Stringifying plates not to deal with NoneType problem
    L_plates = [str(x) for x in L_plates]
    L_plates2 = [str(x) for x in L_plates2]
    
     # Splitting merged truck plates
    L_plates = [''.join(re.sub('/', '', x)).strip() for x in L_plates ]
    L_plates = [''.join(re.sub('86', '86**', x)).strip() if x.endswith('86') and len(x) > 9 else x for x in L_plates ]
    L_plates = [x.split('**') if x != None else 'None' for x in L_plates]
    L_plates = list(itertools.chain.from_iterable(L_plates))
    # Cleaning truck plates off Nones
    L_plates = [x for x in L_plates if x != '']
    L_plates = [x for x in L_plates if x != 'None']
    
    # Splitting merged trailer plates
    L_plates2 = [''.join(re.sub('/', '', x)).strip() for x in L_plates2 ]
    L_plates2 = [''.join(re.sub('86', '86**', x)).strip() if x.endswith('86') and len(x) > 9 else x for x in L_plates2 ]
    L_plates2 = [x.split('**') if x != None else 'None' for x in L_plates2]
    L_plates2 = list(itertools.chain.from_iterable(L_plates2))
    # Cleaning trailer plates off Nones
    L_plates2 = [x for x in L_plates2 if x != '']
    L_plates2 = [x for x in L_plates2 if x != 'None']
   
    
    
    # Introducing del marker for items containing " (e.g. '3пл-5"кл-5,0"'), deleting items with del marker
    L_plates2 = [''.join(re.sub('\"', 'del', x)).strip() for x in L_plates2 if x != None]
    L_plates2 = ['' if 'del' in x else x for x in L_plates2]
    
    
    # Removing items that don't have numbers (i.e. not plates, e.g. 'автоцистерна'), or are the blank ones after the del removal
    pattern_D = re.compile(r'\d')
    L_plates2 = [x for x in L_plates2 if re.findall(pattern_D, str(x))]
   
    # merge truck and trailer plates
    L_plates3 = L_plates + L_plates2
    
 
    L_crws, L_unts, L_lcs = [], [], []
    for i in L_plates3:
        if cursor.execute(f"SELECT Crews FROM Units_Locs_Fixed WHERE Plates_1 like '%{i}%'").fetchall():
            L_crws.append(cursor.execute(f"SELECT Crews FROM Units_Locs_Fixed WHERE Plates_1 like '%{i}%'").fetchall())
            L_unts.append(cursor.execute(f"SELECT Units FROM Units_Locs_Fixed WHERE Plates_1 like '%{i}%'").fetchall())
            L_lcs.append(cursor.execute(f"SELECT Locs FROM Units_Locs_Fixed WHERE Plates_1 like '%{i}%'").fetchall())
        elif cursor.execute(f"SELECT Crews FROM Units_Locs_Fixed WHERE Plates_2 like '%{i}%'").fetchall():
            L_crws.append(cursor.execute(f"SELECT Crews FROM Units_Locs_Fixed WHERE Plates_2 like '%{i}%'").fetchall())
            L_unts.append(cursor.execute(f"SELECT Units FROM Units_Locs_Fixed WHERE Plates_2 like '%{i}%'").fetchall())
            L_lcs.append(cursor.execute(f"SELECT Locs FROM Units_Locs_Fixed WHERE Plates_2 like '%{i}%'").fetchall())

    
    L_crws_temp = []
    for i in L_crws:
        L_crws_temp.append(set(i))
    
    
    L_crws = [', '.join(map(str, x)) for x in L_crws_temp]
    L_unts = [', '.join(map(str, x)) for x in L_unts]
    L_lcs = [', '.join(map(str, x)) for x in L_lcs]
    
    
    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions_long = [186, 797,  126, 188,  174, 158, 196, 156]
        L_regions_short = [86, 96, '02', '07', 82, 78, 54, 52, 77, 88, 89, 74, 76]
        for i in L_regions_long:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 9 else x for x in plates]
        for i in L_regions_short:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        

        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates
    
    
    L_plates_ind = transform_plates(L_plates3) 
    
    
    df = pd.DataFrame(zip(L_crws, L_unts, L_plates3, L_plates_ind, L_lcs), columns=['Crews', 'Units', 'Plates', 'Plate_index', 'Locs'])
    df = df.drop_duplicates(subset='Plate_index', keep="first")
    
    # df.sort_values('Crews')
    
    # Post df to DB
    cursor.execute("DROP TABLE IF EXISTS Units_Locs")
    df.to_sql(name='Units_Locs', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()



   
   
    
    # # Fixing discrepancies in plates
    # D_brm_descrepancy = json.load(open('D_brm_descrepancy.json'))    
    # for k, v in D_brm_descrepancy.items():
    #     L_plates_ind = [x.replace(k, v) for x in L_plates_ind]

    # # Match CD to omnicomm, if matches pull unit name from omnicomm into L_unit
    # L_om_units = cursor_om.execute("SELECT Vehicle FROM final_DB").fetchall()
    # L_om_index = cursor_om.execute("SELECT Plate_index FROM final_DB").fetchall()

    
    # L_plates_ind = set(L_plates_ind)
    
    # L_matched_plates = [x for x in L_plates_ind if x in L_om_index]
    # L_unmatched_plates = [x for x in L_plates_ind if x not in L_om_index]

    # # pprint(L_unmatched_plates)
   
    # # Matching plates to Omnicomm
    # df = pd.DataFrame(zip(L_om_units, L_om_index), columns=['Units', 'Plate_index'])
    # df = df.loc[df.Plate_index.isin(L_matched_plates)]
    # df = df.drop_duplicates(subset=['Plate_index'], keep='first')
   
    # # Derivating matched units from matched plates
    # L_matched_units = df.loc[:, 'Units']
    # L_matched_index = df.loc[:, 'Plate_index']
    
    # # Verifying matched and unmatched against total by plates
    # # pprint(f'Number of all trucks by plates: {len(L_plates_ind)}')
    # # pprint(f'Number of trucks matched to Omnicomm: {len(L_matched_index)}')
    # # pprint(f'Number of trucks unmatched to Omnicomm: {len(L_unmatched_plates)}')
    # # pprint(f'Sum of matched and unmatched equals to all trucks: {len(L_plates_ind) == (len(L_matched_plates) + len(L_unmatched_plates))}')
    
    # # Checking if a unit is missing in D_unmatched_trucks database dictionary @ ditc_brm
    # D_unmatched_trucks = json.load(open('D_brm_unmatchedTrucks.json'))
    # for i in L_unmatched_plates:
    #     if i not in D_unmatched_trucks:
    #         print(f'Vehicles not in D_unmached_db {i} or maybe should be in D_descrepancy_fix, please add manually into dict_brm.py')
    
    
        
    # # Derivating numeric only plate indeces to fish locations Units_Locs_Raw db
    # def numeric_maker(L):
    #     L = [''.join(re.findall(r'\d+', x)).lower() for x in L if x != None]
    #     return L
    # L_matched_numeric = numeric_maker(L_matched_index)
    # L_unmatched_numeric = numeric_maker(L_matched_index)
    
    
    # # Building dataframes of unmatched units and plates
   
    # L_unmatched_units, L_unmatched_plates = [], []
    # for k, v in D_unmatched_trucks.items():
    #     if k in L_unmatched_plates:
    #         L_unmatched_units.append(v)
    #         L_unmatched_plates.append(k) 
    
    # # Derivating numeric only plate indeces to fish locations Units_Locs_Raw db
    # def numeric_maker(L):
    #     L = [''.join(re.findall(r'\d+', x)).lower() for x in L if x != None]
    #     return L
    # L_matched_numeric = numeric_maker(L_matched_index)
    # L_unmatched_numeric = numeric_maker(L_unmatched_plates)

    # df_unmatched = pd.DataFrame(zip(L_unmatched_units, L_unmatched_plates, L_unmatched_numeric), columns=['Units', 'Plate_index', 'Numeric'])
    # df_matched = pd.DataFrame(zip(L_matched_units, L_matched_index, L_matched_numeric), columns=['Units', 'Plate_index', 'Numeric']) 
    
    # # Turn plates into 123abc type
    # def region_cutter(plates):
    #     L_regions = [186, 86, 797, '02', '07', 82, 78, 54, 52, 77, 126, 188, 88, 89, 174, 74, 158, 196, 156, 76]
        
    #     for i in L_regions:
    #         plates = [x.removesuffix(str(i)).strip() if x != None and len(x) > 7 else x for x in plates]
    #     return plates
    # L_plates_cut = region_cutter(L_plates3)
    
   
    
    
    

    # pprint(len(L_plates3))
    
    
if __name__ == '__main__':
    main()

