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
    L_crews = cursor.execute("SELECT Crews FROM Units_Locs_Raw").fetchall()
    L_units = cursor.execute("SELECT Units FROM Units_Locs_Raw").fetchall()
    L_locs = cursor.execute("SELECT Fields FROM Units_Locs_Raw").fetchall()
    
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

    
    # Listing plate numbers from L_units
    L_plates = [x for x in L_units]
    
    def plate_ripper(L, regex):
        L = [re.findall(regex, x) if re.findall(regex, x) else x for x in L]
        L = [''.join([str(y) for y in x]) if isinstance(x, list) else x for x in L]
        return L
        
    L_plates =  plate_ripper(L_units, '\D{2}\s\d+\s\d+')
    L_regex = [ 
                # Plates w/regions
                '\w{2}\s*\d{4}\s*\d+', 
                '\w{1}\s*\d{3}\s*\w{2}\s*\d+', '\d{4}\w{2}\s\d+', '\d{4}\s\w{2}\s\d+', '\d{4}\s\w{2}\s/\d+', 
                '\w{1}\s\d{3}\s\w{2}\s\d+', '\d{4}\w{2}\d+', 
                '\w{1}\d{3}\w{2}\d+', 
                '\d{4}\s\w{2}\d+', 
                # Crutch
                '86\s\D{2}\s\d{4}', '\D{2}\s\d{2}\-\d{2}\s\d+',
               
                # # Diesel stations
                '\ДЭС\s\АД\s\d{2}\s\D\s\d{3}', '\инв\s*\d{4}',
                # For digit numerals
                
                ]
    for i in L_regex:
        L_plates =  plate_ripper(L_plates, i)

    
    pprint(L_plates)
    pprint(len(L_plates))

    # for i in L_plates:
    #     if sum(c.isdigit() for c in i) == 4:
    #         print(i)
        
            
    # L_plates1 = []
    # for i in L_plates:
    #     for j in L_regex:
    #         if re.findall(j, i):
    #             i = L_plates.index(i)
    #             L_plates1.append(L_plates.pop(i))
    #             break
    # pprint(L_plates1)
    # pprint(len(L_plates1))
    # pprint(L_plates)
    # pprint(len(L_plates))
    
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))