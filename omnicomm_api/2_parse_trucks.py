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
    units = json.load(open('JSON_om_units.json')) 
    
    # Parser functions
    # Parser root function
    def root_parser(groupnum):
        L1, L2, L3 = [], [], []
        L_objects = units["children"][groupnum]['objects']
        for i in L_objects:
            L1.append(units["children"][groupnum]['name'])
            L2.append(i['name'])
            L3.append(i['uuid'])
        return L1, L2, L3
   
     # Parser nested level 1
    def level1_parser(groupnum):
        L1, L2, L3 = [], [], []
        L_objects = units["children"][groupnum]['children']
        
        for i in L_objects:
            for j in i['objects']:
                L1.append(i['name'])
                L2.append(j['name'])
                L3.append(j['uuid'])
        return L1, L2, L3

     # Parser nested level 3
    def level3_parser(groupnum, chldnum):
        L1, L2, L3 = [], [], []    
        L_objects = units["children"][groupnum]['children'][chldnum]['children']
        for i in L_objects:
            for j in i['objects']:
                L1.append(i['name'])
                L2.append(j['name'])
                L3.append(j['uuid'])
        return L1, L2, L3
    
    
    # appending ungroupped units
    L_units, L_names, L_id = [], [], []
    L_objects = units["objects"]
    for i in L_objects:
        L_names.append(units['name'])
        L_units.append(i['name'])
        L_id.append(i['uuid'])
 
    df_ungroupped = pd.DataFrame(zip(L_names, L_units, L_id), columns = ['Groups', 'Units', 'id'])

   
    
    # Parsing all groups without nested lists
    L_all_obj, L_all_names, L_all_units, L_all_uuid = [], [], [], []
    for i in range(0, 64):
        L_all_obj.append(root_parser(i))
    
    for i in L_all_obj:
        L_all_names.append(i[0])
        L_all_units.append(i[1])
        L_all_uuid.append(i[2])
    L_all_names = list(itertools.chain.from_iterable(L_all_names))
    L_all_units = list(itertools.chain.from_iterable(L_all_units))
    L_all_uuid = list(itertools.chain.from_iterable(L_all_uuid))

    df_all_root = pd.DataFrame(zip(L_all_names, L_all_units, L_all_uuid), columns = ['Groups', 'Units', 'id'])
    df = pd.merge(df_ungroupped, df_all_root, how="outer")
    
    # Parsing nested neftemash
    df_nested_neftemash = level1_parser(60)
    df_nested_neftemash = pd.DataFrame(zip(df_nested_neftemash[0], df_nested_neftemash[1], df_nested_neftemash[2]), columns = ['Groups', 'Units', 'id'])
    df = pd.merge(df, df_nested_neftemash, how="outer")
    
    # Parsing nested uts
    # Level 1
    uts_level1 = level1_parser(61)  
    df_uts_level1 = pd.DataFrame(zip(uts_level1[0], uts_level1[1], uts_level1[2]), columns = ['Groups', 'Units', 'id'])
    df = pd.merge(df, df_uts_level1, how="outer")

    # Level 3
    uts_atz = level3_parser(61, 3) 
    uts_drvs_ident = level3_parser(61, 6) 
    uts_spec_units = level3_parser(61, 11) 
    
    L_alllev3_names = [uts_atz[0] + uts_drvs_ident[0] + uts_spec_units[0]]
    L_alllev3_units = [uts_atz[1] + uts_drvs_ident[1] + uts_spec_units[1]]
    L_alllev3_uuid = [uts_atz[2] + uts_drvs_ident[2] + uts_spec_units[2]]
    L_alllev3_names = list(itertools.chain.from_iterable(L_alllev3_names))
    L_alllev3_units = list(itertools.chain.from_iterable(L_alllev3_units))
    L_alllev3_uuid = list(itertools.chain.from_iterable(L_alllev3_uuid))
    
    df_lev3 = pd.DataFrame(zip(L_alllev3_names, L_alllev3_units, L_alllev3_uuid), columns = ['Groups', 'Units', 'id'])
    df = pd.merge(df, df_lev3, how="outer")
    
    # Parsing nested trasportation dept
    trans_lev1 = level1_parser(62)
    
    df_trans = pd.DataFrame(zip(trans_lev1[0], trans_lev1[1], trans_lev1[2]), columns = ['Groups', 'Units', 'id'])
    df = pd.merge(df, df_trans, how="outer")
    df = df.drop_duplicates(subset='id', keep="last")
   
    pprint(df)
    
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS Groups_units")
    df.to_sql(name='Groups_units', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()

    
    

if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))