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
db = sqlite3.connect('cits.db')
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
    for i in range(0, 64):
        L_units.append(get_units(units["children"][i]['objects']))
        
    for i in range(0, 64):
        L_names.append(get_names(units["children"][i]['name'], L_units[i]))
    
    L_units = list(itertools.chain.from_iterable(L_units))
    L_names = list(itertools.chain.from_iterable(L_names))
    
    df = pd.DataFrame(zip(L_names, L_units))
    print(df)
    
    # appending level 1 nested groups (ООО Нефтемаш)
    # df = pd.DataFrame(list(units.items()))
    # df = pd.DataFrame(list(units["children"][61]['objects']))
    
    # print(df)
    # print(df.describe())
    
    check_units = get_units(units["children"][61]['children'][0]['objects'])
    check_names = get_names(units["children"][61]['name'], L_units[:1])
    df = pd.DataFrame(zip(check_names, check_units))
    print(df)
    
    # L_units_nested, L_names_nested = [], []
    # L_units_nested = get_units(units["children"][60]['objects'])
    # L_names_nested = get_names(units["children"][60]['name'], L_uts_units)
    # def nested_crasher_units(rng, path1, path2):
    #     pprint(path1)
    #     L_units_nested = []
    #     for i in range(0, rng):
    #         L_units_nested.append(get_units(path1 + [i] + path2))
    #         # L_units_nested = list(itertools.chain.from_iterable(L_units_nested))
    #     return L_units_nested
    
    # def nested_crasher_names(rng, ):
    #         L_names_nested = []
    #         for i in range(0, rng):
    #             L_names_nested.append(get_names(units["children"][60]['name'], L_units_nested[i]))
    #             L_names_nested = list(itertools.chain.from_iterable(L_names_nested))
    #         return L_names_nested  
    
    # path1 = units["children"][60]['children']
    # path2 = ['objects']
    # L_units_nested =  nested_crasher_units(6, path1, path2) 
    # pprint(L_units_nested)
    
    # L_units += L_units_nested
    # L_names += L_names_nested
    
    # dealing with unrepeated unnested group (ЮТС root)
    # L_uts_units = get_units(units["children"][61]['objects'])
    # L_uts_names = get_names(units["children"][61]['name'], L_uts_units)
    
    # L_units += L_uts_units
    # L_names += L_uts_names
    
    # dealing with uts level 1 nested groups
    
    
    
   
    
    
    
    uts_cranes_units = get_units(units["children"][61]['children'][0]['objects'])
    uts_cranes_names = get_names(units["children"][61]['name'], uts_cranes_units)

    uts_tank10_units = get_units(units["children"][61]['children'][1]['objects'])
    uts_tank10_names = get_names(units["children"][61]['name'], uts_tank10_units)
    
    uts_rentals_units = get_units(units["children"][61]['children'][2]['objects'])
    uts_rentals_names = get_names(units["children"][61]['name'], uts_rentals_units)
    
    uts_atz_tanks_units = get_units(units["children"][61]['children'][3]['children'][0]['objects'])
    uts_atz_tanks_names = get_names(units["children"][61]['name'], uts_atz_tanks_units)

    uts_atz_kamaz_units = get_units(units["children"][61]['children'][3]['children'][1]['objects'])
    uts_atz_kamaz_names = get_names(units["children"][61]['name'], uts_atz_kamaz_units)

    uts_atz_trailer_units = get_units(units["children"][61]['children'][3]['children'][2]['objects'])
    uts_atz_trailer_names = get_names(units["children"][61]['name'], uts_atz_trailer_units)

    uts_atz_truck_units = get_units(units["children"][61]['children'][3]['children'][3]['objects'])
    uts_atz_truck_names = get_names(units["children"][61]['name'], uts_atz_truck_units)
    
    uts_atz_ural_units = get_units(units["children"][61]['children'][3]['children'][4]['objects'])
    uts_atz_ural_names = get_names(units["children"][61]['name'], uts_atz_ural_units)
    
    # pprint(uts_atz_ural_units)
    # pprint(uts_atz_ural_names)
    # pprint(len(uts_atz_ural_units))
    # pprint(len(uts_atz_ural_names))
    # 
    # pprint(check_units)
    # pprint(check_names)

    # nobeloil_units = get_units(units["children"][50]['objects'])
    # nobeloil_names = get_names(units["children"][50]['name'], nobeloil_units)
    # !Для просмотра ООО "Няганьнефть"
    # "РН-Пурнефтегаз"

    
    

if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))