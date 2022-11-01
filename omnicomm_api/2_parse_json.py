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
db = sqlite3.connect('om.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()


def main():
    units = json.load(open('JSON_om_units.json')) 
    online = json.load(open('JSON_om_online.json')) 
    
    
    
    
    L_units, L_names = [], []
    def get_units(search):
        L_units = []
        for i in search:
            L_units.append(i["name"])
        return L_units
    def get_names(name, length):
        L_names = []
        for i in range(0, len(length)):
            L_names.append(name)
        L_names = [x.replace('\t', ' ') for x in L_names]
        return L_names
    
    # appending ungroupped units
    L_units.append(get_units(units["objects"]))
    L_names.append(get_names(units["name"], L_units)) 
    
    # appending groups that don't have nested groups (i.e Toyota, Frac, ct)
    for i in range(0, 61):
        L_units.append(get_units(units["children"][i]['objects']))
        
    for i in range(0, 61):
        L_names.append(get_names(units["children"][i]['name'], L_units[i]))
    
    L_units = list(itertools.chain.from_iterable(L_units))
    L_names = list(itertools.chain.from_iterable(L_names))
    
    # appending groups with level 1 nested list (ООО Нефтемаш [60])
    L_nested_units, L_nested_names = [], []
    for i in range(0, 6):
        L_nested_units.append(get_units(units["children"][60]['children'][i]['objects']))
        
    for i in range(0, 6):
        L_nested_names.append(get_names(units["children"][60]['name'] + '. ' + units["children"][60]['children'][i]['name'], L_nested_units[i]))
    
    L_nested_units = list(itertools.chain.from_iterable(L_nested_units))
    L_nested_names = list(itertools.chain.from_iterable(L_nested_names)) 
    
    # Merging dfs
    df_nested = pd.DataFrame(zip(L_nested_names, L_nested_units), columns = ['Groups', 'Units'])
    df = pd.DataFrame(zip(L_names, L_units), columns = ['Groups', 'Units'])
    df = pd.merge(df, df_nested, how="outer")
    
    # Collecting root folder for UTS [61]
    L_units.clear()
    L_names.clear()
    L_units.append(get_units(units["children"][61]['objects']))
    L_units = list(itertools.chain.from_iterable(L_units))
    L_names.append(get_names(units["children"][61]['name'], L_units))
    L_names = list(itertools.chain.from_iterable(L_names))
    
    # Mergin dfs
    df61 = pd.DataFrame(zip(L_names, L_units), columns = ['Groups', 'Units'])
    df = pd.merge(df, df61, how="outer")
   
    # Horrible UTS
    # appending groups with level 1 nested list (UTS [61])
    L_nested_units.clear()
    L_nested_names.clear()
    for i in range(0, 13):
        L_nested_units.append(get_units(units["children"][61]['children'][i]['objects']))
        
    for i in range(0, 13):
        L_nested_names.append(get_names(units["children"][61]['name'] + '. ' + units["children"][61]['children'][i]['name'], L_nested_units[i]))
    
    L_nested_units = list(itertools.chain.from_iterable(L_nested_units))
    L_nested_names = list(itertools.chain.from_iterable(L_nested_names)) 
    
    df_nested = pd.DataFrame(zip(L_nested_names, L_nested_units), columns = ['Groups', 'Units'])
    df = pd.merge(df, df_nested, how="outer")
    
    # appending groups with level 3 nested list (UTS ATZ[61])
    L_nested_units.clear()
    L_nested_names.clear()
    for i in range(0, 5):
        L_nested_units.append(get_units(units["children"][61]['children'][3]['children'][i]['objects']))
        
    for i in range(0, 5):
        L_nested_names.append(get_names(units["children"][61]['name'] + '. ' + units["children"][61]['children'][3]['children'][i]['name'], L_nested_units[i]))
    
    L_nested_units = list(itertools.chain.from_iterable(L_nested_units))
    L_nested_names = list(itertools.chain.from_iterable(L_nested_names)) 
    
    df_nested = pd.DataFrame(zip(L_nested_names, L_nested_units), columns = ['Groups', 'Units'])
    df = pd.merge(df, df_nested, how="outer")
    
    # appending groups with level 3 nested list (UTS special unit[61])
    L_nested_units.clear()
    L_nested_names.clear()
    for i in range(0, 5):
        L_nested_units.append(get_units(units["children"][61]['children'][11]['children'][i]['objects']))
        
    for i in range(0, 5):
        L_nested_names.append(get_names(units["children"][61]['name'] + '. ' + units["children"][61]['children'][11]['children'][i]['name'], L_nested_units[i]))
    
    L_nested_units = list(itertools.chain.from_iterable(L_nested_units))
    L_nested_names = list(itertools.chain.from_iterable(L_nested_names)) 
    
    df_nested = pd.DataFrame(zip(L_nested_names, L_nested_units), columns = ['Groups', 'Units'])
    df = pd.merge(df, df_nested, how="outer")
    
    # Collecting root folder for Transportation service [62]
    L_units.clear()
    L_names.clear()
    L_units.append(get_units(units["children"][62]['objects']))
    L_units = list(itertools.chain.from_iterable(L_units))
    L_names.append(get_names(units["children"][62]['name'], L_units))
    L_names = list(itertools.chain.from_iterable(L_names))
    
    # Mergin dfs
    df62 = pd.DataFrame(zip(L_names, L_units), columns = ['Groups', 'Units'])
    df = pd.merge(df, df62, how="outer")
    
    # appending groups with level 1 nested list
    L_nested_units.clear()
    L_nested_names.clear()
    for i in range(0, 10):
        L_nested_units.append(get_units(units["children"][62]['children'][i]['objects']))
        
    for i in range(0, 10):
        L_nested_names.append(get_names(units["children"][62]['name'] + '. ' + units["children"][62]['children'][i]['name'], L_nested_units[i]))
    
    L_nested_units = list(itertools.chain.from_iterable(L_nested_units))
    L_nested_names = list(itertools.chain.from_iterable(L_nested_names)) 
    
    df_nested = pd.DataFrame(zip(L_nested_names, L_nested_units), columns = ['Groups', 'Units'])
    df = pd.merge(df, df_nested, how="outer")
    
    # Collecting root Filatov [63]
    L_units.clear()
    L_names.clear()
    L_units.append(get_units(units["children"][63]['objects']))
    L_units = list(itertools.chain.from_iterable(L_units))
    L_names.append(get_names(units["children"][63]['name'], L_units))
    L_names = list(itertools.chain.from_iterable(L_names))
    
    # Mergin dfs
    df63 = pd.DataFrame(zip(L_names, L_units), columns = ['Groups', 'Units'])
    df = pd.merge(df, df63, how="outer")
    df = df.drop_duplicates(subset='Units', keep="last")
    pprint(df)
    
    # checking paths
    df = pd.DataFrame(list(units.items()))
    df = pd.DataFrame(list(units["children"][62]['children'][0]))
    # df = pd.DataFrame(list(units["children"][61]['children'][3]['children'][0]['objects']))
    
    
    print(df)
    print(df.describe())
    
    
    # # Posting df to DB
    # print('Posting df to DB')
    # cursor.execute("DROP TABLE IF EXISTS Groups_units")
    # df.to_sql(name='Groups_units', con=db, if_exists='replace', index=False)
    # db.commit()
    # db.close()

    
    

if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))