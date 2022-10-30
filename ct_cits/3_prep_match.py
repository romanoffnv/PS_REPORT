from dataclasses import dataclass
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
# connection to cits.db
db = sqlite3.connect('cits.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

# connection to omnicomm.db
db_om = sqlite3.connect('omnicomm.db')
db_om.row_factory = lambda cursor, row: row[0]
cursor_om = db_om.cursor()


def main():
    L_crews = cursor.execute("SELECT Crews FROM Units_Locs_Parsed").fetchall()
    L_units = cursor.execute("SELECT Units FROM Units_Locs_Parsed").fetchall()
    L_plates = cursor.execute("SELECT Plates FROM Units_Locs_Parsed").fetchall()
    L_locs = cursor.execute("SELECT Locations FROM Units_Locs_Parsed").fetchall()
    
    
    # L_unit_om  = cursor_om.execute("SELECT Vehicle FROM final_DB").fetchall()
    # L_plates_om = cursor_om.execute("SELECT Plate FROM final_DB").fetchall()

    # df = pd.DataFrame(zip(L_unit_om, L_plates_om))
    # print(df)

    # Pre-cleaning fields
    L_locs = ['-' if v == 'None' else v for v in L_locs]

    # Slicing field name (Ю/Приобское м/р\nООО "ГАЗПРОМНЕФТЬ-ХАНТОС" - Ю/Приобское м/р)
    L_locs_temp = []
    for i in L_locs:
        if 'м/р' in i:
            ind = i.index('м/р')
            L_locs_temp.append(i[:ind + 3])
        else:
            L_locs_temp.append(i)
    
    L_locs = [x for x in L_locs_temp]
    L_locs_temp.clear()
    
    # Clean plates
    L_cleanit = ['\-', '/']
    for i in L_cleanit:
        L_plates = [re.sub(i, '', x) for x in L_plates]
   
    # Converting unconditioned plates into conditioned ones thru the manually supported dict
    D_ct_plates = json.load(open('D_ct_plates.json'))    
    for k, v in D_ct_plates.items():
        for j in L_plates:
            if k == j:
                ind = L_plates.index(j)
                L_plates[ind] = v

                
    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions = [186, 86, 797, '02', '07', 82, 78, 54, 77, 126, 188, 89, 88, 174, 74, 158, 196, 156, 76]
        
        for i in L_regions:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) > 7 else x for x in plates]
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates


    
    
    

    # zip(L_crews, L_units, L_plates, L_plates_ind, L_locs)
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))