# This file grabs db as accountance_1
# Wraps it in df
# Filters df out by keywords so only relavant equipments stays
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

# Excel connection  
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\acc.xls")
ws1 = wb.Worksheets(1)

# Pandas
pd.set_option('display.max_rows', None)

# db connections
db = sqlite3.connect('accountance.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
pd.set_option('display.max_rows', None)

def main():
    # exp 1
    # Get accountance_1 db cols as lists L_mols and L_items
    L_mols = cursor.execute("SELECT Responsible FROM accountance_1").fetchall()
    L_items = cursor.execute("SELECT Item FROM accountance_1").fetchall()
    
    # Build dataframe
    data = pd.DataFrame(zip(L_mols, L_items), columns=['Responsible', 'Item'])
    
    # Populate list of keywords for df filtration
    L_items_filter = ['г/н', 'гос.№', 'гос№', 'гос. №', 'Truck', 'VIN', 'Насосная установка', 
                    'Mercedes', 'KENWORTH', 'Передвижная паровая установка', 'ППУ', 
                    'Полуприцеп', 'прицеп', 'тягач', 'Кран', 'Гидратационная установка', 'Автоцистерна', 'смеситель',
                    'блендер', 'КАМАЗ', 'Камаз', ' гн ']
    
    # Filter df by keywords
    data = data[data['Item'].str.contains('|'.join(L_items_filter))]
    
    # Destructure df into lists
    L_items = data.loc[:, 'Item'].tolist()
    L_mols = data.loc[:, 'Responsible'].tolist()
    # Derive the list of untouchable units to post it into db later
    L_units = [x for x in L_items] 
      
    # Slice items starting from the keyword's index to the end of the sentence
    def slicer (x, L_items):
        L_items_temp = []
        if x == 'VIN' or x == 'vin' or x == 'VIV':
            for i in L_items:
                if x in i:
                    ind = i.index(x)
                    L_items_temp.append(i[:ind])
                else:
                    L_items_temp.append(i)
            
            L_items = [str(x).strip() for x in L_items_temp]
            L_items_temp.clear() 
            
            return L_items
        else:
            for i in L_items:
                if x in i:
                    ind = i.index(x)
                    L_items_temp.append(i[ind:])
                else:
                    L_items_temp.append(i)
            
            L_items = [str(x).strip() for x in L_items_temp]
            L_items_temp.clear() 
            
            return L_items

    L_keywords = ['г/н', '№', ' гн ', 'г/р', 'г.н.', 'Г/н', 'Truck', 'VIN', 'vin', 'VIV']
    for i in L_keywords:
        L_items = slicer(i, L_items)

        # Remove crap like 'г/н' etc
    def crapRemover(x, L_items):
        L_items = [i.replace(x, '') for i in L_items]
        return L_items

    L_keywords = ['г/н', '№', 'гн ', 'г.н.', 'Truck', 'г/р', 'Г/н', ')', ' ', 'RUS', ';', ',', ':']
    for i in L_keywords:
        L_items = crapRemover(i, L_items)

    # Fishing out plates by regex from long sentences
    def regexBomber(x, L_items):
        
        L_plates_temp = []
        for i in L_items:
            if re.findall(x, str(i)):
                L_plates_temp.append(''.join(re.findall(x, str(i))))
            else:
                L_plates_temp.append(i)
                # print(i)
    
        L_items = [str(x).strip() for x in L_plates_temp]
        L_plates_temp.clear() 
            
        return L_items

    L_items = regexBomber(re.compile('\s\D\s*\d+\D{2}\s*\d+'), L_items)
    L_items = regexBomber(re.compile('\D{1}\s*\d+\s*\D{2}\s*\d+'), L_items)

    # Remove paranthesis and content
    L_patterns = ['\(.*\)', '\(.*']
    for i in L_patterns:
        L_items = [''.join(re.sub(i, '', x)).strip() for x in L_items]

    # Remove pointless sentences
    L_items = [x if len(x) < 20 else '' for x in L_items]
    
    # Remove regions
    L_items = [x.removesuffix('186').strip() if x != None and len(x) == 9 else x for x in L_items]
    L_items = [x.removesuffix('797').strip() if x != None else x for x in L_items]

    
    L_items = [x.removesuffix('86').strip() if x != None and len(x) == 8 else x for x in L_items]
    L_items = [x.removeprefix('86').strip() if x != None and len(x) == 8 else x for x in L_items]
    L_items = [x.removesuffix('77').strip() if x != None and len(x) == 8 else x for x in L_items]
    L_items = [x.removeprefix('77').strip() if x != None and len(x) == 8 else x for x in L_items]
    L_items = [x.removesuffix('82').strip() if x != None and len(x) == 8 else x for x in L_items]
    L_items = [x.removesuffix('89').strip() if x != None and len(x) == 8 else x for x in L_items]
    L_items = [x.removesuffix('78').strip() if x != None and len(x) == 8 else x for x in L_items]
    L_items = [x.removesuffix('76').strip() if x != None and len(x) == 8 else x for x in L_items]
    L_items = [x.removesuffix('94').strip() if x != None and len(x) == 8 else x for x in L_items]

    
    
    # Make some crutches
    L_items = [x.removesuffix('7').strip() if x != None and len(x) == 7 else x for x in L_items]
    L_items = [x.removesuffix('86').strip() if x != None and len(x) == 6 and x.isnumeric() else x for x in L_items]
    L_items = [x.removeprefix('-30').strip() if x != None and len(x) == 9 else x for x in L_items]
    
    # L_items = [x if len(x) < 9 and len(x) > 5 else '' for x in L_items]
    # pprint(L_items)
    # Bring plates to 111abc format 
    L_total_literal = [''.join(re.findall(r'\D', x)).lower() for x in L_items]
    L_total_literal = [''.join(re.sub(r'\s+', '', x)).lower() for x in L_total_literal]
    L_total_index = [''.join(re.findall(r'\d+', x)) for x in  L_items]
    
    L_temp = [x + y for x, y in zip(L_total_index, L_total_literal)]
    L_acc_plate_index = [x for x in L_temp]
    
    # Crutch replacements
   
    D_crutches = {
                        '13621-наприцепетра': '',
                        '5668автоцистернаск': '5668ск',
                        '65221тягачкамаз': '652внт',
                        '2502/': '',
                        '2501/': '',
                        '300полуприцепм.рс-': '',
                        '25001-к-': '',
                        '461bhp': '461внр',
                        
            }        
        
    # Replacing crappy unit names into omnicomm smth
    for k, v in D_crutches.items():
        L_acc_plate_index = [x.replace(k, v) for x in L_acc_plate_index]

    pprint(L_acc_plate_index)
    # Push lists to DB
    cursor.execute("DROP TABLE IF EXISTS accountance_2;")
    cursor.execute("""
            CREATE TABLE IF NOT EXISTS accountance_2(
            Responsible text,
            Unit text,
            Plate_index text)
                    """)
    cursor.executemany("INSERT INTO accountance_2 VALUES (?, ?, ?)", zip(L_mols, L_units, L_acc_plate_index))
        
    db.commit()
    db.close()
    
if __name__ == '__main__':
    main()
