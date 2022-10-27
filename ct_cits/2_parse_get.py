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
    def units_preserver(L, regex):
        L = [x for x in L if re.findall(regex, x) ]
        L = [''.join([str(y) for y in x]) if isinstance(x, list) else x for x in L]
        return L

    # Fishing plates
    L_plates1 = plate_ripper(L_units, '\w{2}\s*\d{4}\s*\d+')
    L_units1 = units_preserver(L_units, '\w{2}\s*\d{4}\s*\d+')

    L_plates2 = plate_ripper(L_units, '\w{1}\s*\d{3}\s*\w{2}\s*\d+')
    L_units2 = units_preserver(L_units, '\w{1}\s*\d{3}\s*\w{2}\s*\d+')
    
    
    # with Crutch /
    L_plates3 = plate_ripper(L_units, '\d{4}\s*\w{2}\s*\/*\d+')
    L_units3 = units_preserver(L_units, '\d{4}\s*\w{2}\s*\/*\d+')
    
    # Crutch
    L_plates4 = plate_ripper(L_units,'86\s\D{2}\s\d{4}')
    L_units4 = units_preserver(L_units, '86\s\D{2}\s\d{4}') 
    
    L_plates5 = plate_ripper(L_units,'\D{2}\s\d{2}\-\d{2}\s\d+')
    L_units5 = units_preserver(L_units, '\D{2}\s\d{2}\-\d{2}\s\d+')
    
    # Diesel stations
    L_plates6 = plate_ripper(L_units,'\ДЭС\s\АД\s\d{2}\s\D\s\d{3}')
    L_units6 = units_preserver(L_units, '\ДЭС\s\АД\s\d{2}\s\D\s\d{3}') 
    L_plates7 = plate_ripper(L_units,'\инв\s*\d{4}')
    L_units7 = units_preserver(L_units, '\инв\s*\d{4}') 
    
    # Bitten Niva Crutch
    L_plates8 = plate_ripper(L_units,'Е134КК')
    L_units8 = units_preserver(L_units, 'Е134КК') 

    L_plates = L_plates1 + L_plates2 + L_plates3 + L_plates4 + L_plates5 + L_plates6 + L_plates7 + L_plates8
    L_units_good = L_units1 + L_units2 + L_units3 + L_units4 + L_units5 + L_units6 + L_units7 + L_units8
    
   
    # Function that pops plates with certain params from units list to avoid regex interference
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

    # Removing all conditioned plates from units (not to interfere for 4 digit non-conditioned plates rip off)
    L_units = plate_popper(0, L_plates, L_units)
    # Ripping off 4 digit non-conditioned plates
    L_plates_4d =  plate_ripper(L_units, '\d{4}')
    L_units_4d = units_preserver(L_units, '\d{4}')
     
    # Removing 4 digit non-cond plates not to interfere with pulling 3 digit nc plates
    L_units = plate_popper(0, L_plates_4d, L_units)
    # Ripping 3 digit nc plates
    L_plates_3d =  plate_ripper(L_units, '\d{3}')
    L_units_3d = units_preserver(L_units, '\d{3}')
   
    # Popping 3 digit nc plates, leaving basically trash
    L_units = plate_popper(0, L_plates_3d, L_units)
    
    
    # Merging all types of plates, removing dubs
    L_plates = L_plates + L_plates_4d + L_plates_3d
    L_units = L_units_good + L_units_4d + L_units_3d
    
    df = pd.DataFrame(zip(L_units, L_plates), columns=['Units', 'Plates'])
    
    # Posting df to DB to be able to collect units later
    # print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS Units_Plates")
    df.to_sql(name='Units_Plates', con=db, if_exists='replace', index=False)
    db.commit()
    

    # Collecting crews and locs
    L_crws, L_lcs, L_matched, L_unmatched = [], [], [] , []
    
    # Checking if plates would match to db, sending them to matched and unmatched lists, Collecting units for unmatched plates
    L_units_matched, L_units_umatched = [], []
    for i in L_plates:
        if cursor.execute(f"SELECT Units FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall():
            L_matched.append(i)
            L_units_matched.append(cursor.execute(f"SELECT Units FROM Units_Plates WHERE Plates like '%{i}%'").fetchall())
        else:
            L_unmatched.append(i)
            L_units_umatched.append(cursor.execute(f"SELECT Units FROM Units_Plates WHERE Plates like '%{i}%'").fetchall())
    
    
    # L_matched 366
    # L_unmatched 4
    # L_unmatched ['Н 397 КС 86', 'ДЭС АД 30 Т 400', 'инв 2219', 'инв 0002']
    
    
    L_unmatched_4d =  plate_ripper(L_unmatched, '\d{4}')
    # L_units_4d = units_preserver(L_units, '\d{3}')
    # L_unmatched_4d - ['2219', '0002']
    # L_unmatched ['инв 2219', 'инв 0002', 'Н 397 КС 86', 'ДЭС АД 30 Т 400']
    
    L_unmatched = plate_popper(0, L_unmatched_4d, L_unmatched)
    # pprint(L_unmatched) - ['ДЭС АД 30 Т 400', 'Н 397 КС 86']
    
    L_unmatched_3d =  plate_ripper(L_unmatched, '\d{3}')
    # pprint(L_unmatched_3d) - ['400', '397']
    
    L_unmatched = plate_popper(0, L_unmatched_3d, L_unmatched)
    # pprint(L_unmatched) - []
    
    L_unmatched_all = L_unmatched_4d + L_unmatched_3d
    # pprint(L_unmatched_all) - ['0002', '2219', '397', '400']
    L_plates = L_matched + L_unmatched_all
    L_units = L_units_matched + L_units_umatched
    
    L_units = [', '.join(map(str, x)) for x in L_units]
    
    
    for i in L_plates:
        if cursor.execute(f"SELECT Units FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall():
            L_crws.append(cursor.execute(f"SELECT Crews FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall())
            L_lcs.append(cursor.execute(f"SELECT Fields FROM Units_Locs_Raw WHERE Units like '%{i}%'").fetchall())
        else:
            L_unmatched.append(i)
    
    # Unpacking nested lists
    L_crws = [', '.join(map(str, x)) for x in L_crws]
    L_lcs = [', '.join(map(str, x)) for x in L_lcs]
    
    df = pd.DataFrame(zip(L_crws, L_units, L_plates, L_lcs), columns=['Crews', 'Units', 'Plates', 'Locations'])
    # pprint(df)
    
    
    # Posting df to DB
    
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS Units_Locs_Parsed")
    df.to_sql(name='Units_Locs_Parsed', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
    
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))

