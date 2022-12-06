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
        # '\)': '), ', 
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
    
    # Fishing out plates by regex from long sentences
    def plate_ripper(L_units):
        def plate_fisher(regex, L_units):
            L_plates_temp = []
            for i in L_units:
                if 'ДЭС' in i:
                    L_plates_temp.append(i)
                else:
                    if re.findall(regex, str(i)):
                        L_plates_temp.append(''.join(re.findall(regex, str(i))))
                    else:
                        L_plates_temp.append(i)
                    # print(i)
        
            L_units = [str(x).strip() for x in L_plates_temp]
            L_plates_temp.clear() 
                
            return L_units
        L_regex = [
            '\s\D{2}\s*\d{2}\-*\s*\d{2}\s*\d+', #ВВ  4553 86, # АН 78 96 82, ВВ  4553 86
            '\s*\D\s*\d{3}\s*\D{2}\s*\d+', #Е 898 СВ 186, У 039 ВК186
            '\s\d{4}\s*\D{2}\s*\d+', #7441УР 86, УХ 3130 86
            '\s*\D{2}\-*\s*\d{4}\s*\d+', #АТ-246786, 
            '\s*\d{4}\s*\D{2}\s*\d+', # 7317 УН 86
            '\s*\d{2}\s*\D{2}\s*\d{4}', #86 УМ 8475
            '\-\d{3}', # -445
            '\№\s*\d+', # инв№0002, инв№2219
        ]
        L_plates = plate_fisher(re.compile(L_regex[0]), L_units) 
        
        for regex in L_regex:
            L_plates = plate_fisher(re.compile(regex), L_plates)
        return L_plates

    L_plates = plate_ripper(L_units)
    
    L_plates_temp = []
    for i in L_plates:
        if 'олуприцеп' in i:
            L_plates_temp.append(''.join(re.findall('\d{4}', str(i))))
        elif 'мкость' in i:
            L_plates_temp.append(''.join(re.findall('\d{4}', str(i))))
        else:
            L_plates_temp.append(i)

    
    L_plates = [x for x in L_plates_temp]
    
    # Getting crews and locs 
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
    
    # pprint(L_plates_unmatched)
    # pprint(len(L_plates_unmatched))

    # Cleaning trash off the plates
    L_cleans = ['(', 'ПС', ')', '-']
    for i in L_cleans:
        L_plates = [x.replace(i, '') for x in L_plates]

    
    # Fixing plates
    # Converting unconditioned plates into conditioned ones thru the manually supported dict
    D_ct_plates = json.load(open('D_ct_plates.json'))    
    for k, v in D_ct_plates.items():
        for j in L_plates:
            if k == j:
                ind = L_plates.index(j)
                L_plates[ind] = v

    # Derivating PI
    def transform_plates(plates):
        plates = [re.sub('\s+', '', x) for x in plates]
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

    L_cits_PI = transform_plates(L_plates)
    # Making filter mask by counting number of digits in plates (to filter out stuff like УН 38мм)
    L_mask = []
    for i in L_cits_PI:
        L_mask.append(sum(c.isdigit() for c in i))
    # Building df
    df = pd.DataFrame(zip(L_crws, L_units, L_plates, L_cits_PI, L_lcs, L_mask), columns=['Crews', 'Units', 'Plates', 'PI', 'Locs', 'Mask'])
    df = df[df['Mask'] > 2]
    df = df.drop(['Mask'], axis=1)
    
    pprint(df)
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS cits_final")
    df.to_sql(name='cits_final', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
    
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))