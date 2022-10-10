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
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\brm.xlsx")
Sheets = wb.Sheets.Count
ws = wb.Worksheets(Sheets)

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
    L_plates = cursor_brm.execute("SELECT Units_1 FROM Units_Locs_Raw").fetchall()
    L_plates2 = cursor_brm.execute("SELECT Units_2 FROM Units_Locs_Raw").fetchall()
    
    
    
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
        L_plates = [''.join(re.sub(k, v, x)).strip() for x in L_plates if x != None]
    
    L_plates = [x.split('**') for x in L_plates]
    L_plates = list(itertools.chain.from_iterable(L_plates))
    L_plates = [x.strip() for x in L_plates if x != '']
    
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
        L_regions = [186, 86, 797, '02', '07', 82, 78, 54, 77, 126, 188, 88, 89, 174, 74, 158, 196, 156, 76]
        
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
    
    D_descrepancy_fix = {
        '435аср': '435акр',
        '408акс': '408авм',
        '680екм': '680екн',
    }

    for k, v in D_descrepancy_fix.items():
        L_plates_ind = [x.replace(k, v) for x in L_plates_ind]

    # Match CD to omnicomm, if matches pull unit name from omnicomm into L_unit
    L_om_units = cursor_om.execute("SELECT Vehicle FROM final_DB").fetchall()
    L_om_index = cursor_om.execute("SELECT Plate_index FROM final_DB").fetchall()

    
    L_plates_ind = set(L_plates_ind)
    # pprint(L_plates_ind)
    # pprint(len(L_plates_ind))

    L_matched_plates = [x for x in L_plates_ind if x in L_om_index]
    L_unmatched_plates = [x for x in L_plates_ind if x not in L_om_index]
    df = pd.DataFrame(zip(L_om_units, L_om_index), columns=['Units', 'Plate_index'])
    df = df.loc[df.Plate_index.isin(L_matched_plates)]
    df = df.drop_duplicates(subset=['Plate_index'], keep='first')
    
    L_matched_units = df.loc[:, 'Units']
    
    D_unmached_names = {
        '7708нх': 'Прицеп',
        '889вмт': 'Тойота hillux',
        '9917ах': 'Прицеп АЦ-17',
        '125как': 'МАЗ АЦ 20м3 НК-Транс',
        '7403ук': 'Прицеп АЦ 10м3',
        '4683ат': 'Полуприцеп',
        '4897ат': 'Полуприцеп',
        '1974вв': 'Прицеп АЦ-17',
        '2847ау': 'Полуприцеп',
        '1823ан': 'Насос ВД',
        '367вмм': 'Тойота hillux',
        '9910ах': 'Площадка',
        '0762н': 'Насос ВД',
        '6562ах': 'Кран-манипулятор',
        '479ммс': 'Шевроле Нива',
        '7458ас': 'Автоцестерна 20м3-полуприцеп',
        '971внх': 'Полуприцеп (Химка)',
        '2130ах': 'Площадка',
        '531хху': 'МАЗ АЦ 20м3 НК-Транс',
        '2872вв': 'Полуприцеп',
        '4697ат': 'Площадка',
        '0877ва': 'Прицеп АЦ-17',
        '813рос': 'Автокран МАЗ ИП Рыжков',
        '686вмт':,
        '6415ах':,
        '1474ах':,
        '120530ghsm':,
        '513сне':,
        '0897ва':,
        '004aae':,
        '740онв':,
        '2861вв':,
        '2862вв':,
        '341внм':,
        '7717нх':,
        '324авр':,
        '508ауа':,
        '365вмм':,
        '0842ва':,
        '2841вв':,
        '7232ау':,
        '6561ах':,
        '3105ат':,
        '2696ат':,
        '250кох':,
        '746кох':,
        '693вмт':,
        '492аое':,
        '5403ва':,
        '564вмм':,
        '898вмм':,
        '9892ах':,
        '7231ау':,
        '395вмт':,
        '0879ва':,
        '2870вв':,
        '184оас':,
        '692ахр':,
        '5824вв':,
        '4740ат':,
        '0909ва':,
        '6420ах':,
        '8098ах':,
        '352вмм':,
        '0932ат':,
        '6339ау':,
        '562вмм':,
        '896вмт':
            }
    pprint(L_unmatched_plates)
    pprint(len(L_unmatched_plates))
    
    # pprint(len(L_matched_plates))
    # pprint(len(L_matched_units))
    # print(len(df))
    

        
        
    # pprint(L)
    # pprint(len(L))
        
    # df = pd.DataFrame(zip(L_om_units, L_om_index), columns=['Units', 'Plate_index'])
    # df = df.loc[df.Plate_index.isin(L_plates_ind)]
    # print(df)
    # print(df.describe())
    # Get whatever is unmatched manually into a dict for names (crutch)
    # Roll CD to subtitute untmached with names in L_unit
    # Multiply by frac crew
    # Populate L_loc by finding in location picking df
if __name__ == '__main__':
    main()