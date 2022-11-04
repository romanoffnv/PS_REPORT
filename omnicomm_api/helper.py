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
import win32com
print(win32com.__gen_path__)


# Making connections to DBs
db = sqlite3.connect('omnicomm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()


def main():
    L1 = ['1', '2', '3', '4', '5']
    L2 = ['1', '2', '3', '4', '5']
    
    L3 = ['5', '4', '3', '2', '1']
    L4 = ['5', '4', '3', '2', '1']
    
    df1 = pd.DataFrame(zip(L1, L2), columns = ['A', 'B'])
    df2 = pd.DataFrame(zip(L3, L4), columns = ['C', 'D'])
    df = df1.join(df2, how = 'left')
    pprint(df)
    

if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))