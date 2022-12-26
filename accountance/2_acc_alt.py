import time
import xlsxwriter
from win32com.client.gencache import EnsureDispatch
import os
import re
from pprint import pprint
import pandas as pd
from functools import reduce
import itertools
import sqlite3
import win32com
print(win32com.__gen_path__)


# Pandas
pd.set_option('display.max_rows', None)

# db connections
db = sqlite3.connect('data.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
pd.set_option('display.max_rows', None)

def main():
    # Get Unit col from accountance_1.db as list
    L_units = cursor.execute("SELECT Unit FROM accountance_1").fetchall()

    # Split strings in Unit list by the keywords and get as L_plates list
    def splitter(split, L):
      L = [str(x) for x in L]  
      L = [x.split(split) for x in L]
      L = list(itertools.chain.from_iterable(L))

      return L

    L_keywords = ['г/н', 
                '43118', 'г.н.'
                'г.н.', 'гн', '(', 'г/р', ';', '43118', 'Гос.№', ',', 'зав.', 'мод.', 'зав', '№', ')', 'ст ', 'Г/н', 
                'АЦН'
                ]
    # Getting pilot splitted list of plates by sending first keyword from the list and L_units
    L_plates = splitter(L_keywords[0], L_units)
    # Keep on splitting by sending keywords and pilot plates list
    for i in L_keywords:
        L_plates = splitter(str(i), L_plates)
    
    # Filter out strings that:
    # have more or less letters than in a real plate
    L_plates = [''.join(x).strip() for x in L_plates if 'изель' in x or (sum(map(str.isalpha, x)) < 4 and sum(map(str.isalpha, x)) > 1)]
    # are shorter than 6 characters
    L_plates = [x for x in L_plates if len(x) > 6]
    # have one of the keys
    L_keys = ['г.в.', 'л.с.', 'VIN', 'НД', 'Квт', 'час', 'ит', 'Gr', 'dpi', 'ф/з', 'FHD', 'до']
    for i in L_keys:
        L_plates = [x for x in L_plates if i not in x]

    # Extract plates from strings by regex into separate lists to avoid regex overlapping
    L_plates_original = [x for x in L_plates]
    def plate_ripper(regex, L):
        L_temp = []
        for i in L:
            if re.findall(regex, str(i)): 
                L_temp.append(''.join(re.findall(regex, str(i))))
            else:
                L_temp.append('')
        L = [x.strip() for x in L_temp]
        return L
    
    
    # Every regex has its own list, with blanks if has no match
    L_plates1 = plate_ripper('\D\s*\d{3}\s*\D{2}\s*\d+', L_plates) #А 779 ЕН 186, А782ОН 186
    L_plates2 = plate_ripper('\D{2}\s\d{4}\s{2}\d+', L_plates) #АХ 6576  86
    L_plates3 = plate_ripper('\d{4}\s\D{2}\s\d+', L_plates) #6654 УС 76
    L_plates4 = plate_ripper('\D{2}\s*\d{4}\s*\d+', L_plates) #АУ2845 86
    L_plates5 = plate_ripper('\d{2}\s*\D{2}\s*\d{4}', L_plates) #86 УК 4804
    L_plates6 = plate_ripper('\d+\s*\d+\s*\D{2}\s*\d+', L_plates) #0288  УВ 86, 06 41 УВ 86
    L_plates7 = plate_ripper('\D{2}\s*\d{2}\s*\d{4}', L_plates) #УВ 86 0594
    L_plates8 = plate_ripper('\D{2}\s\d{4}', L_plates) #АХ 9399
    # Regex conflict!!!, prior regex cuts S/N 0015286 into /N 0015
    L_plates9 = plate_ripper('\D{2}\s\d{4}', L_plates) #S/N 0015286
    
    # Merge obtained regex lists into the one
    def listmerger(L1, L2):
        L = []
        for x, y in zip(L1, L2):
            if x != '':
                L.append(x)
            elif x == '' or len(x) < len(y):
                L.append(y)    
        return L

    # Clearing initial L_plates list to reuse it as a pilot list for merging
    L_plates.clear()
    
    # Collecting regexed lists names into array
    L_regs = [
        L_plates1, L_plates2, 
        L_plates3, L_plates4, L_plates5, L_plates6, L_plates7, L_plates8, 
        L_plates9]
    # Getting pilot list by merging the first two lists
    L_plates = listmerger(L_regs[0], L_regs[1])
    # Sending pilot list and lists from array into listmerger to keep on merging
    for i in L_regs:
        L_plates = listmerger(L_plates, i)
    
    
    df = pd.DataFrame(zip(L_plates_original, L_plates))
    print(df)
    

    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
