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
db_om = sqlite3.connect('omnicomm.db')
db_om.row_factory = lambda cursor, row: row[0]
cursor_om = db_om.cursor()

db_cits = sqlite3.connect('cits.db')
db_cits.row_factory = lambda cursor, row: row[0]
cursor_cits = db_cits.cursor()

def main():
    L_index_om = cursor_om.execute("SELECT Plate_index FROM Final_DB").fetchall()
    L_num_index_om = [re.sub('\D', '', x) for x in L_index_om]
    
    L_dept_cits = cursor_cits.execute("SELECT Dept FROM Trucks").fetchall()
    L_unit_cits = cursor_cits.execute("SELECT Unit FROM Trucks").fetchall()
    L_plate_cits = cursor_cits.execute("SELECT Plate FROM Trucks").fetchall()
    L_plate_index_cits = cursor_cits.execute("SELECT Plate_index FROM Trucks").fetchall()

    # converting cits plate indexes in to omnicomm standard ( '081' to '081век')
    # creating dictionary of {'999нва': '999'} format
    D = dict(zip(L_index_om, L_num_index_om))

    # matching cits against dict keys and vals and substitute list item with dict keys if any match (match last)
    for i in L_plate_index_cits:
        ind = L_plate_index_cits.index(i)
        for k, v in D.items():
            if i == k or i == v:
                L_plate_index_cits[ind] = k

    

    def dict_matcher(x):
        L = []
        D = dict(zip(L_plate_index_cits, x))
        for i in L_index_om:
            L.append(D.get(i))
        return L
    
    L_dept_matched = dict_matcher(L_dept_cits)
    L_unit_matched = dict_matcher(L_unit_cits)    
    L_plate_matched = dict_matcher(L_plate_cits)    
    L_plate_index_matched = dict_matcher(L_plate_index_cits)    
    
    df = pd.DataFrame(zip(L_dept_matched, L_unit_matched, L_plate_matched, L_plate_index_matched), columns=['Dept', 'Unit', 'Plate', 'Plate_index'])
    
    
    def dict_unmatcher(x):
        L = []
        D = dict(zip(L_plate_index_cits, x))
        for i in L_plate_index_cits:
            if i not in L_index_om:
                L.append(D.get(i))
        return L
    
    L_dept_unmatched = dict_unmatcher(L_dept_cits)
    L_unit_unmatched = dict_unmatcher(L_unit_cits)    
    L_plates_unmatched = dict_unmatcher(L_plate_cits)    
    L_plate_index_unmatched = dict_unmatcher(L_plate_index_cits) 
    
    
    
    # Post into db
    cursor_cits.execute("DROP TABLE IF EXISTS Trucks_matched")
    df.to_sql(name='Trucks_matched', con=db_cits, if_exists='replace', index=False)
    db_cits.commit()
    
    df2 = pd.DataFrame(zip(L_dept_unmatched, L_unit_unmatched, L_plates_unmatched, L_plate_index_unmatched), columns=['Dept', 'Unit', 'Plate', 'Plate_index'])
    cursor_cits.execute("DROP TABLE IF EXISTS Trucks_unmatched")
    df2.to_sql(name='Trucks_unmatched', con=db_cits, if_exists='replace', index=False)
    db_cits.commit()
    db_cits.close()
   


            
    
    # print(df)
    # print(df.describe())
    
    # Collecting lists of umnached items
    # L_dept_unmatched, L_unit_unmatched, L_plate_unmatched, L_plate_index_unmatched = [], [], [], []
    # for i, j, k, l in zip(L_dept_cits, L_unit_cits, L_plate_cits, L_plate_index_cits):
    #     if l not in L_index_om:
    #         L_dept_unmatched.append(i)
    #         L_unit_unmatched.append(j)
    #         L_plate_unmatched.append(k)
    #         L_plate_index_unmatched.append(l)
    
    # df2 = pd.DataFrame(zip(L_dept_unmatched, L_unit_unmatched, L_plate_unmatched, L_plate_index_unmatched))
    # print(len(L_plate_index_unmatched))
    # pprint(L_unmatched)
    # pprint(L_plate_index_cits)
    # pprint(len(L_plate_index_cits))
    
    # pprint(len(L_unmatched))
if __name__ == '__main__':
    main()