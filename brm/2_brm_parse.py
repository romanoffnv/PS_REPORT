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

# Get the Excel Application COM object
# xl = EnsureDispatch('Excel.Application')
# wb = xl.Workbooks.Open(f"{os.getcwd()}\\brm.xlsx")
# Sheets = wb.Sheets.Count
# ws = wb.Worksheets(Sheets)

# Making connections to DBs
db_brm = sqlite3.connect('brm.db')
db_brm.row_factory = lambda cursor, row: row[0]
cursor_brm = db_brm.cursor()

db_om = sqlite3.connect('omnicomm.db')
db_om.row_factory = lambda cursor, row: row[0]
cursor_om = db_om.cursor()

# Pandas
pd.set_option('display.max_rows', None)

def main():
    L_crews = cursor_brm.execute("SELECT Crews FROM Units_Locs_Raw").fetchall()
    L_plates = cursor_brm.execute("SELECT Units_1 FROM Units_Locs_Raw").fetchall()
    L_plates2 = cursor_brm.execute("SELECT Units_2 FROM Units_Locs_Raw").fetchall()
    L_locs = cursor_brm.execute("SELECT Locs FROM Units_Locs_Raw").fetchall()
    
    
    # Clean col C 
    D_replacers = {
        '\s+': ' ',
        '/': '',
        '186': '186**'
    }
    D_replacers2 = {
        '\s+': ' ',
        '\"': 'del',
        'Аренда с ЮТС': '',
        'АЦН-17': '',
    }
    
    # Crutch
    D_truck_fix = {
        '618':'618внт',
        '865':'865вмт',
    }

    for k, v in D_replacers.items():
        L_plates = [''.join(re.sub(k, v, x)).strip() if x != None else '' for x in L_plates ]

    L_plates = [x.split('**') for x in L_plates]
    L_plates = list(itertools.chain.from_iterable(L_plates))
    L_plates = [x.strip() for x in L_plates if x != '']

    L_crews1 = []
    for i in L_plates:
        # print(i)
        L_crews1.append(cursor_brm.execute(f"SELECT Crews FROM Units_Locs_Raw WHERE Units_1 like '%{i}%'").fetchall())
    L_crews1 = list(itertools.chain.from_iterable(L_crews1))
    L_crews1 = [x.strip() for x in L_crews1 if x != '']

    df = pd.DataFrame(zip(L_crews1, L_plates))
    pprint(df)
    # pprint(L_crews1)
    # pprint(len(L_crews1))

    # print(len(L_crews))
    # print(len(L_plates))
    # print(len(L_plates2))
    # print(len(L_locs))
    
    # Clean col D
    L_splits = ['\n', 'ВД', '/']
    for i in L_splits: 
        L_plates2 = [x.split(i) for x in L_plates2 if x != None]
        L_plates2 = list(itertools.chain.from_iterable(L_plates2))
    
    for k, v in D_replacers2.items():
        L_plates2 = [''.join(re.sub(k, v, x)).strip() for x in L_plates2 if x != None]
    L_plates2 = ['' if 'del' in x else x for x in L_plates2]
    # Crutch
    L_plates2 = ['' if x == '403' else x for x in L_plates2]
    L_plates2 = [x.strip() for x in L_plates2  if x != '']

     # Removing items that don't have numbers (i.e. plates)
    pattern_D = re.compile(r'\d')
    L_plates2 = [x for x in L_plates2 if re.findall(pattern_D, str(x))]

    # merge cols into CD
    L_plates3 = L_plates + L_plates2
    L_plates = [x for x in L_plates3]
    L_plates = [x.replace('.0', '') for x in L_plates]
    
    # pprint(L_plates)
    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions = [186, 86, 797, '02', '07', 82, 78, 54, 52, 77, 126, 188, 88, 89, 174, 74, 158, 196, 156, 76]
        
        for i in L_regions:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) > 7 else x for x in plates]
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates
    
    
    L_plates_ind = transform_plates(L_plates) 
   
    for k, v in D_truck_fix.items():
        L_plates_ind = [x.replace(k, v) for x in L_plates_ind]
    # Removing bare indeces like '743вмт' which are duplicates for
    L_plates_ind = [x for x in L_plates_ind if len(x) != 3]
    
    # Fixing discrepancies in plates
    D_brm_descrepancy = json.load(open('D_brm_descrepancy.json'))    
    for k, v in D_brm_descrepancy.items():
        L_plates_ind = [x.replace(k, v) for x in L_plates_ind]

    # Match CD to omnicomm, if matches pull unit name from omnicomm into L_unit
    L_om_units = cursor_om.execute("SELECT Vehicle FROM final_DB").fetchall()
    L_om_index = cursor_om.execute("SELECT Plate_index FROM final_DB").fetchall()

    
    L_plates_ind = set(L_plates_ind)
    

    L_matched_plates = [x for x in L_plates_ind if x in L_om_index]
    L_unmatched_plates = [x for x in L_plates_ind if x not in L_om_index]

    # Matching plates to Omnicomm
    df = pd.DataFrame(zip(L_om_units, L_om_index), columns=['Units', 'Plate_index'])
    df = df.loc[df.Plate_index.isin(L_matched_plates)]
    df = df.drop_duplicates(subset=['Plate_index'], keep='first')
   
    # Derivating matched units from matched plates
    L_matched_units = df.loc[:, 'Units']
    L_matched_index = df.loc[:, 'Plate_index']
    
    # Verifying matched and unmatched against total by plates
    # pprint(f'Number of all trucks by plates: {len(L_plates_ind)}')
    # pprint(f'Number of trucks matched to Omnicomm: {len(L_matched_index)}')
    # pprint(f'Number of trucks unmatched to Omnicomm: {len(L_unmatched_plates)}')
    # pprint(f'Sum of matched and unmatched equals to all trucks: {len(L_plates_ind) == (len(L_matched_plates) + len(L_unmatched_plates))}')
    
    # Checking if a unit is missing in D_unmatched_trucks database dictionary @ ditc_brm
    D_unmatched_trucks = json.load(open('D_brm_unmatchedTrucks.json'))
    for i in L_unmatched_plates:
        if i not in D_unmatched_trucks:
            print(f'Vehicles not in D_unmached_db {i} or maybe should be in D_descrepancy_fix, please add manually into dict_brm.py')
    
    
    # Building dataframes of matched and unmatched units and plates
   
    L1, L2 = [], []
    for k, v in D_unmatched_trucks.items():
        if k in L_unmatched_plates:
            L1.append(v)
            L2.append(k) 
    df_unmatched = pd.DataFrame(zip(L1, L2), columns=['Units', 'Plate_index'])
    
    
    # Derivating numeric only plate indeces to fish locations Units_Locs_Raw db
    def numeric_maker(L):
        L = [''.join(re.findall(r'\d+', x)).lower() for x in L if x != None]
        return L
    L_matched_numeric = numeric_maker(L_matched_index)
    L_unmatched_numeric = numeric_maker(df_unmatched.loc[:, 'Plate_index'])
    
    # db to df
    df_matched = pd.DataFrame(zip(L_matched_units, L_matched_index, L_matched_numeric), columns=['Units', 'Plate_index', 'Numeric']) 
    # pprint(df_matched)

    L_crws, L_lcs = [], []
    for i in L_matched_numeric:
        # print(i)
        L_crws.append(cursor_brm.execute(f"SELECT Crews FROM Units_Locs_Raw WHERE Units_1 OR Units_1 like '%{i}%'").fetchall())
        # L_lcs.append(cursor_brm.execute(f"SELECT Locs FROM Units_Locs_Raw WHERE Units_1 like '%{i}%'").fetchall())
    # pprint(L_crws)
    # df = pd.DataFrame(zip(L_crws, L_matched_units, L_matched_index, L_lcs))
    # pprint(df)
    # pprint(len(df))
    # L_plates = cursor_brm.execute("SELECT Units_1 FROM Units_Locs_Raw").fetchall()
    # L_plates2 = cursor_brm.execute("SELECT Units_2 FROM Units_Locs_Raw").fetchall()
    # pprint(len(L_crews))
    # pprint(len(L_plates))
    # pprint(len(L_plates2))
    # pprint(len(L_locs))
    
    # L_c, L_p, L_l = [], [], []
    # for i in L_matched_numeric:
    #     for c, p, l in zip(L_crews, L_plates, L_locs):
    #         if i in str(p):
    #             L_c.append(c) 
    #             L_p.append(i)
    #             L_l.append(l)
    #             break

    
    # L_c2, L_p2, L_l2 = [], [], []
    # for i in L_matched_numeric:
    #     for c, p, l in zip(L_crews, L_plates2, L_locs):
    #         if i in str(p):
    #             L_c2.append(c) 
    #             L_p2.append(i)
    #             L_l2.append(l)
    #             break
            
    # pprint(L_l2)
  
    # L_crews = L_c + L_c2
    # L_plates = L_p + L_p2
    # L_locs = L_l + L_l2
    # pprint(len(L_crews))
    # pprint(len(L_plates))
    # pprint(len(L_locs))

    # ddf = pd.DataFrame(zip(L_crews, L_plates, L_locs), columns=['Crews', 'Plate_index', 'Loc'])
    # ddf = ddf.drop_duplicates(subset=['Plate_index'], keep='first')
    # pprint(ddf)
    # pprint(ddf.describe())
if __name__ == '__main__':
    main()

