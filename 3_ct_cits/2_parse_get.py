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

db = sqlite3.connect('data.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()


def main():
    # Pulling cits_get.db into lists
    L_units = cursor.execute("SELECT Units FROM cits_get").fetchall()
    
    # Cleaning L_units
    L_cleanwords = ['Цель работ:', 'профессия', 'Бурильщик', 'Пом.бур', 'Маш-т', '\n', 'гос№', 'гос',
                    # '№', 
                    '\.', 'Вагоны:']
    for i in L_cleanwords:
        L_units = [re.sub(i, ' ', x).strip() for x in L_units if x != None]
    L_units = [x for x in L_units if x != '']
    
    
    # Splitting merged cells
    # Inserting comma after the following items to split by later
    D_replacers = {
        # '\s+': ' ', 
        ';': ',', 
        '\+': ',', 
        'в пути': ',',
        'трал': 'трал ',
        ' 86-': ' 86,',
        ' 186': ' 186,',
        'RUS': ',', 
        # Setting del marker
        'катушка': 'del',
    }
    for k, v in D_replacers.items():
        L_units = [''.join(re.sub(k, v, x)).strip() for x in L_units ]
   
    # Special: Inserting comma after region to split plates, exluding leading 86 (i.e. УГА АР 8629 86, not УГА АР 86,29 86)    
    L_units = [re.sub(' 86', ' 86,', x) if not '86.*' else x for x in L_units]
    
    
    # Spltting merged strings by comma, unpacking lists, stripping, removing '' items
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
  
    # This function splits merged strings by params
    def splitter(L, a, b, c):
        
        L = [re.sub(a, b, x) if c in str(x) else x for x in L]    
        L = [x.split(',') for x in L]
        L = list(itertools.chain.from_iterable(L))
        return L
    
    # Splitting merged strings by sending to splitter function with params
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

    
    # This function exctracts plate numbers from L_units
    def plate_ripper(L, regex):
        # listing plates by regex
        L = [re.findall(regex, x) for x in L if re.findall(regex, x) ]
        # extracting nested lists
        L = [''.join([str(y) for y in x]) if isinstance(x, list) else x for x in L]
        return L

    # This function keeps L_units items after they have been ripped off plates, to match units vs plates
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

    # Merging lists of all ripped off conditioned plates
    L_plates_conditioned = L_plates1 + L_plates2 + L_plates3 + L_plates4 + L_plates5 + L_plates6 + L_plates7 + L_plates8
    # Merging lists of all conditioned units been ripped off for plates
    L_units_conditioned = L_units1 + L_units2 + L_units3 + L_units4 + L_units5 + L_units6 + L_units7 + L_units8
   
    # Function that pops plates with certain params from units list to avoid regex interference
    def plate_popper(num, L_plates_conditioned, L_units):
        while True:
            for i in L_plates_conditioned:
                for j in L_units: 
                    if i in j:
                        ind = L_units.index(j)
                        L_units.pop(ind)
            num += 1
            if num == 10:
                break
        return L_units

    # Removing all conditioned plates from units (not to interfere for 4 digit non-conditioned plates rip off)
    L_units = plate_popper(0, L_plates_conditioned, L_units)
    
    # Ripping off 4 digit non-conditioned plates
    L_plates_4d =  plate_ripper(L_units, '\d{4}')
    # Keeping units to match 4 digit plates
    L_units_4d = units_preserver(L_units, '\d{4}')
    
    # Removing 4 digit non-cond plates not to interfere with pulling 3 digit plates
    L_units = plate_popper(0, L_plates_4d, L_units)

    # Ripping 3 digit plates
    L_plates_3d =  plate_ripper(L_units, '\d{3}')
    # Keeping units to match 3 digit plates
    L_units_3d = units_preserver(L_units, '\d{3}')
   
    # Popping 3 digit plates, leaving basically trash like ['УН 38мм', 'АН№3']
    L_units = plate_popper(0, L_plates_3d, L_units)
    
    # Adding extra space to 3d numbers to avoid multiple match as a substring of 4d (ie 371 in 3714)
    L_plates_3d = [x + '  ' for x in L_plates_3d]

    # Merging all types of plates
    L_plates = L_plates_conditioned + L_plates_4d + L_plates_3d
    L_units = L_units_conditioned + L_units_4d + L_units_3d
    # pprint(len(L_plates))
    # pprint(len(L_units))
    # df = pd.DataFrame(zip(L_units, L_plates))
    # print(df)

    # Building df, dropping dubs
    df = pd.DataFrame(zip(L_units, L_plates), columns=['Units', 'Plates'])
    df = df.drop_duplicates(subset='Plates', keep="first")
    
    # Posting df to DB 'Units_Plates' to be able to collect units later
    cursor.execute("DROP TABLE IF EXISTS Units_Plates")
    cursor.execute("DROP TABLE IF EXISTS cits_Units_Plates")
    df.to_sql(name='cits_Units_Plates', con=db, if_exists='replace', index=False)
    db.commit()
    
    # Listing plates from verified df (matched with units, no dubs)
    L_plates = df['Plates'].tolist()
    L_units = df['Units'].tolist()

    # Stripping plates back to normal
    L_plates = [str(x).strip() for x in L_plates]
        
    # Collecting crews and locs
    L_crws, L_lcs = [], []
    L_plates_unmatched = []
    for i in L_plates:
        if cursor.execute(f"SELECT Units FROM cits_get WHERE Units like '%{i}%'").fetchall():
            L_crws.append(cursor.execute(f"SELECT Crews FROM cits_get WHERE Units like '%{i}%'").fetchall())
            L_lcs.append(cursor.execute(f"SELECT Fields FROM cits_get WHERE Units like '%{i}%'").fetchall())
        else:
            L_plates_unmatched.append(i)
            
    
    # Unpacking nested lists
    L_crws = [', '.join(map(str, x)) for x in L_crws]
    L_lcs = [', '.join(map(str, x)) for x in L_lcs]
    
    df = pd.DataFrame(zip(L_crws, L_units, L_plates, L_lcs), columns=['Crews', 'Units', 'Plates', 'Locations'])
    
    print(df)
    
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS cits_parse")
    df.to_sql(name='cits_parse', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
    
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))

