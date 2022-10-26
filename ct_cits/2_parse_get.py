from sys import prefix
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

db = sqlite3.connect('cits.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()


def main():
     # Pulling Units_Locs_Raw.db into lists
    L_units = cursor.execute("SELECT Units FROM Units_Locs_Raw").fetchall()
    
    
    # Cleaning L_units
    L_cleanwords = ['Цель работ:', 'профессия', 'Бурильщик', 'Пом.бур', 'Маш-т', '\n', 'гос№', 'гос',
                  '№', '\.', 'Вагоны:']
    for i in L_cleanwords:
        L_units = [re.sub(i, ' ', x).strip() for x in L_units if x != None]
    L_units = [x for x in L_units if x != '']
    
    # Splitting merged cells
     # Inserting comma after the following items to split by later
    D_replacers = {
        '\s+': ' ', 
        ';': ',', 
        '\+': ',', 
        'в пути': ',',
        'трал': 'трал ',
        ' 86-': ' 86,',
        'RUS': ',', 
        # Setting del marker
        'катушка': 'del',
    }
    for k, v in D_replacers.items():
        L_units = [''.join(re.sub(k, v, x)).strip() for x in L_units ]
   
    # Special: Inserting comma after region to split plates, exluding leading 86 (i.e. УГА АР 8629 86, not УГА АР 86,29 86)    
    L_units = [re.sub(' 86', ' 86,', x) if not '86.*' else x for x in L_units]
    
    L_units = [x.split(',') for x in L_units]
    L_units = list(itertools.chain.from_iterable(L_units))
    L_units = [str(x).strip() for x in L_units if x != '']
    L_units = [x for x in L_units if x != '']
    
    # Replacing cits abbreviations with items names of Omnicomm
    D_ct_replacers = {
                'С/Т': 'Камаз ',
                'С/С': 'Камаз ',
                'В/А': 'Камаз ',
                'В/О': 'Верхее оборудование ',
                'А/К': 'Автокран ',
                'В/Б': 'Вакуумбочка ',
                'п/п': 'Полуприцеп ',
                'НТ': 'МЗКТ ',
                'НКА': 'НКА ',
                'ПКА ': 'Жидко-азотный агрегат ',
                'трал': 'Полуприцеп ',
                'УНБ': 'Камаз УНБ',
                'KW': 'KENWORTH',
                'F/S': 'Камаз ',
            }    
    for k, v in D_ct_replacers.items():
        L_units = [''.join(re.sub(k, v, x)).strip() for x in L_units ]
  
    # Doing the rest of the splits to extract whatever is in paranthesis or slashed diesel stations
    def splitter(L, a, b, c):
        
        L = [re.sub(a, b, x) if c in str(x) else x for x in L]    
        L = [x.split(',') for x in L]
        L = list(itertools.chain.from_iterable(L))
        return L
    
    L_units = splitter(L_units, '\(', '\,(', 'ПС')    
    L_units = splitter(L_units, '\)', '\),', 'ПС')    
    L_units = splitter(L_units, '/', ',', 'ДЭС')    
    

    # Removing items that don't have numbers (i.e. not plates, e.g. 'автоцистерна'), or are the blank ones after the del removal
    pattern_D = re.compile(r'\d')
    L_units = [''.join(x).strip() for x in L_units if re.findall(pattern_D, str(x))]
    # Removing items with del marker
    L_units = ['' if 'del' in str(x) else x for x in L_units]
    L_units = [x for x in L_units if x != '']
    L_units = [re.sub(r'\\', '', x) for x in L_units]

    print(f'initial {len(L_units)}')
    L_units_original = [x for x in L_units]

    # Listing plate numbers from L_units
    def plate_ripper(L, regex):
        L = [re.findall(regex, x) for x in L if re.findall(regex, x) ]
        L = [''.join([str(y) for y in x]) if isinstance(x, list) else x for x in L]
        return L
        
    # Fishing plates
    L_plates1 = plate_ripper(L_units, '\w{2}\s*\d{4}\s*\d+')
    L_plates2 = plate_ripper(L_units, '\w{1}\s*\d{3}\s*\w{2}\s*\d+')
    
    # with Crutch /
    L_plates3 = plate_ripper(L_units, '\d{4}\s*\w{2}\s*\/*\d+')
        
    # Crutch
    L_plates4 = plate_ripper(L_units,'86\s\D{2}\s\d{4}') 
    L_plates5 = plate_ripper(L_units,'\D{2}\s\d{2}\-\d{2}\s\d+')
    
               
    # Diesel stations
    L_plates6 = plate_ripper(L_units,'\ДЭС\s\АД\s\d{2}\s\D\s\d{3}') 
    L_plates7 = plate_ripper(L_units,'\инв\s*\d{4}')
    

    # Bitten Niva Crutch
    L_plates8 = plate_ripper(L_units,'Е134КК')
    
    L_plates = L_plates1 + L_plates2 + L_plates3 + L_plates4 + L_plates5 + L_plates6 + L_plates7 + L_plates8
    L_plates = list(set(L_plates))
    
    
    # Popping collected plates from units, leaving items prepped for 4, 3 digit plates rip off
    def plate_popper(num, L_plates, L_units):
        while True:
            for i in L_plates:
                for j in L_units: 
                    if i in j:
                        ind = L_units.index(j)
                        L_units.pop(ind)
            num += 1
            if num == 10:
                break
        return L_units

    L_units = plate_popper(0, L_plates, L_units)
    L_plates_4d =  plate_ripper(L_units, '\d{4}')
    L_units = plate_popper(0, L_plates_4d, L_units)
    L_plates_3d =  plate_ripper(L_units, '\d{3}')
    L_units = plate_popper(0, L_plates_3d, L_units)
    
    L_plates = L_plates + L_plates_4d + L_plates_3d
    L_plates = list(set(L_plates))
    
    
   
    # Collecting crews and locs
    L_crws, L_lcs, L_unmatched = [], [], []
    
    for i in L_plates:
        if cursor.execute(f"SELECT Units FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall():
            pass
        else:
            L_unmatched.append(i)
    
    # pprint(f'checkin plates: {len(L_plates)}')
    L_unmatched_4d =  plate_ripper(L_unmatched, '\d{4}')
    L_unmatched = plate_popper(0, L_unmatched_4d, L_unmatched)
    L_unmatched_3d =  plate_ripper(L_unmatched, '\d{3}')
    L_unmatched = plate_popper(0, L_unmatched_3d, L_unmatched)
    
    L_unmatched_all = L_unmatched_4d + L_unmatched_3d
    L_plates = L_plates + L_unmatched_all
    
    
    for i in L_plates:
        if cursor.execute(f"SELECT Units FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall():
            print(i)
            
            L_crws.append(cursor.execute(f"SELECT Crews FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall())
            L_lcs.append(cursor.execute(f"SELECT Fields FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall())
        # else:
            
        #     L_unmatched.append(i + ' tracer')
    
    # Recollecting units
    # for i in L_plates:
    #     if cursor.execute(f"SELECT Crews FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall():
    #         L_crws.append(cursor.execute(f"SELECT Crews FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall())
    #         L_lcs.append(cursor.execute(f"SELECT Fields FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall())
    #     else:
    #         L_unmatched.append(i)
    
    L_units.clear()
    for i in L_plates:
        for j in L_units_original:
            if i in j:
                L_units.append(j)
                break
    
    
    
    # pprint(len(L_crws))
    # pprint(len(L_lcs))
    # pprint(len(L_units))
    # pprint(len(L_plates))
    # df = pd.DataFrame(zip(L_crws, L_units, L_plates, L_lcs), columns=['Crews', 'Units', 'Plates', 'Locations'])
    # pprint(df)
    
    
    # Posting df to DB
    
    # print('Posting df to DB')
    # cursor.execute("DROP TABLE IF EXISTS Units_Locs_Parsed")
    # df.to_sql(name='Units_Locs_Parsed', con=db, if_exists='replace', index=False)
    # db.commit()
    # db.close()
    
    
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))

