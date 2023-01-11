import time
import collections
import xlsxwriter
from win32com.client.gencache import EnsureDispatch
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
db = sqlite3.connect('data.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
cnx = sqlite3.connect('data.db')

pd.set_option('display.max_rows', None)

def main():
    # ****************************** IMPORTS FROM DB ***************************************
    L_mols = cursor.execute("SELECT Mols FROM accountance_3").fetchall()
    L_units = cursor.execute("SELECT Units FROM accountance_3").fetchall()
    L_plates = cursor.execute("SELECT Plates FROM accountance_3").fetchall()

    # ****************************** REUSABLE FUNCTIONS ************************************
    def revert_86(L_plates):
        L_plates_temp = []
        for i in L_plates:
            if i.startswith('86 '):
                L_plates_temp.append(i.replace('86 ', '') + ' 86')
            else:
                L_plates_temp.append(i)
        return L_plates_temp
    

    def transform_plates(plates):
        plates = [re.sub('\s+', '', x) for x in plates]
        L_regions_long = [102, 126, 156, 158, 174, 186, 188, 196, 797]
        L_regions_short = ['01', '02', '03', '04', '05', '06', '07', '09']
        # Trucks
        for i in L_regions_long:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 9 else x for x in plates]
        # Trailers
        for i in L_regions_short:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 else x for x in plates]
        for i in range(1, 10):
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 9 else x for x in plates] #6657 УС 7
        for i in range(10, 100):
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 else x for x in plates]
        
        # kzh Trailers
        plates = [x.removesuffix(str('07')).strip() if 'kzн' in str(x) and len(x) == 9 else x for x in plates]
        
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_numeric = [x[:4] if len(x) == 5 else x for x in plates_numeric]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        # plates = [str(x) for x  in plates_numeric]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates

    # ******************************* FUNCTION CALL PARAMS ****************************************
    # Reverting 86th region from the first position to the last i.e. 86 М 212 ХО  -> М 212 ХО 86
    L_plates = revert_86(L_plates)
    
    # Derivating plate index list
    L_PI = transform_plates(L_plates)
    L_PI = [x if len(x) == 6 else '' for x in L_PI]
    L_plates = [x.replace(x, '') if y == '' else x for x, y in zip(L_plates, L_PI) ]
    D = dict(zip(L_plates, L_PI))
    pprint(D)
    pprint(len(D))
    


if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))