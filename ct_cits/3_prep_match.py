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
    L_locs = cursor.execute("SELECT Locations FROM Units_Locs_Parsed").fetchall()
    L_plates = cursor.execute("SELECT Plates FROM Units_Locs_Parsed").fetchall()

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
    pprint(L_plates)
    # Fix plates from dict that have been registered @ omnicomm
    # D_plates = {
    #     '694ПС': 'а694мв 186',
    #     'ПС600': 'т600ак 186',
    #     '86 УМ 8475': 'УМ 8475 86',
    # }
    # Open paranthesis like '(ПС197)'
    # L_replacers = ['\(', '\ПС', '\)']
    # for i in L_replacers:
    #     L_plates = [''.join(re.sub(i, '', x)).strip() for x in L_plates]

    
    

if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))