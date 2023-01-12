import time
import collections
import xlsxwriter
from win32com.client.gencache import EnsureDispatch
import sys
import os
import re
from pprint import pprint
import pandas as pd
import numpy as np
from functools import reduce
import itertools
import sqlite3
import win32com
print(win32com.__gen_path__)



# Pandas
pd.set_option('display.max_rows', None)

# db connections
path = '/Users/roman/OneDrive/Рабочий стол/SANDBOX/PS_REPORT'
file = os.path.join(path, 'data.db')
db = sqlite3.connect(file)
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
# cnx = sqlite3.connect('data.db')

pd.set_option('display.max_rows', None)

def main():
    # Get Unit col from accountance_2.db as list
    L_units = cursor.execute("SELECT Units FROM accountance_2").fetchall()
    
    # ******************************* REUSABLE FUNCTIONS ****************************************
    
    # STRING SPLITTER
    # Landing splitting marks
    def mark_landing(s, L_units):
        L = [re.sub(s, '*split*', x) for x in L_units]
        return L
    
    # Split strings in Unit list by the keywords and get as L_plates list
    def splitter(L):
        L = [x.split('*split*') for x in L]
        L = list(itertools.chain.from_iterable(L))
    

        return L

    # ******************************* FUNCTION CALL PARAMS ****************************************
    
    # Sending keywords to mark_landing func
    L_keywords = ['г/н', 'Truck', '43118', 'г.н.', 'гн', '\(', 'г/р', ';', ',', 'Гос.№', 'зав.',
                'зав', '№', '\)', 'ст ', 'Г/н', 'АЦН', 'электростанция', 'дизельный']

    for i in L_keywords:
        L_units = mark_landing(str(i), L_units)
    
    # Getting splitted list by marks as L_plates
    L_plates  = splitter(L_units)
   
    # ********************************************************************************************
    
    def plate_validator(L_plates):
        L_literals = []
        L_numeric = []
        L_illegal, L_illegal2, L_illegal3, L_illegal_lits, L_illegal_signs, L_illegal_regex = [], [], [], [], [], []
        L_plates_alt = []
        L_length = []
        for i in L_plates:
            L_plates_alt.append(''.join(re.sub('\s+', '', i)).strip())
            
            L_literals.append(sum(map(str.isalpha, i)) < 4 and sum(map(str.isalpha, i)) > 1)
            L_numeric.append(sum(map(str.isnumeric, i)) < 7 and sum(map(str.isnumeric, i)) > 2)
            L_illegal.append(''.join(re.findall('[Й, й, Ц, ц, Г, г, Ш, ш, Щ, щ, З, з, Ф, ф, Ы, ы, П, п, Л, л, Д, д, Ж, ж, Э, э, Я, я, Ч, ч, Ь, ь, Ъ, ъ, Б, б, Ю, ю]', i)).strip())
            L_illegal2.append(''.join(re.findall('[=, /, ", \., :, Gr, dpi, Mb]', i)).strip())
            L_illegal3.append(''.join(re.findall('\d{3}\х\d{3}\s*\мм', i)).strip())
           
            
        for i, j, k in zip(L_illegal, L_illegal2, L_illegal3):
            L_illegal_lits.append(i == '')
            L_illegal_signs.append(j == '')
            L_illegal_regex.append(k == '')
        
        for i in L_plates_alt:
            if len(i) < 6:
                L_length.append(False)
            else:
                L_length.append(True)
                
        df = pd.DataFrame(zip(L_plates, L_literals, L_numeric, L_illegal_lits, L_illegal_signs, L_length, L_illegal_regex), 
                          columns = ['Plates', 'Literals', 'Numeric', 'Illegal_lits', 'Illegal_signs', 'Length', 'L_illegal_regex'])
        df = df[df['Literals'] == True]
        df = df[df['Numeric'] == True]
        df = df[df['Illegal_lits'] == True]
        df = df[df['Illegal_signs'] == True]
        df = df[df['Length'] == True]
        return df
        # have more or less letters than in a real plate 
        # L_plates = [''.join(x).strip() for x in L_plates if (sum(map(str.isalpha, x)) < 4 and sum(map(str.isalpha, x)) > 1)]
    
    df = plate_validator(L_plates)
    
    pprint(df)
    pprint(df.describe())
    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))