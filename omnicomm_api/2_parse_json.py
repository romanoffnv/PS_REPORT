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
    
    
    # df = pd.DataFrame(list(units.items()))
    # df = pd.DataFrame(list(units['objects']))
    
    # print(df)
    # print(df.describe())
    
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
    for i in range(0, 60):
        L_units.append(get_units(units["children"][i]['objects']))
        
    for i in range(0, 60):
        L_names.append(get_names(units["children"][i]['name'], L_units[i]))
    
    L_units = list(itertools.chain.from_iterable(L_units))
    L_names = list(itertools.chain.from_iterable(L_names))
    
    L_units_nested, L_names_nested = [], []
    for i in range(0, 6):
        L_units_nested.append(get_units(units["children"][60]['children'][i]['objects']))
        
    for i in range(0, 6):
        L_names_nested.append(get_names(units["children"][60]['name'], L_units_nested[i]))
        
    L_units_nested = list(itertools.chain.from_iterable(L_units_nested))
    L_names_nested = list(itertools.chain.from_iterable(L_names_nested))
    
    L_units += L_units_nested
    L_names += L_names_nested
    
    L_uts_units = get_units(units["children"][61]['objects'])
    L_uts_names = get_names(units["children"][61]['name'], L_uts_units)
    
    L_units += L_uts_units
    L_names += L_uts_names
    
    df = pd.DataFrame(zip(L_names, L_units))
    print(df)
    
    
    
   
    
    
    
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
    
    toyota_units = get_units(units["children"][0]['objects'])
    # toyota_names = get_names(units["children"][0]['name'], toyota_units)
    
    # uts_rental_units = get_units(units["children"][1]['objects'])
    # uts_rental_names = get_names(units["children"][1]['name'], uts_rental_units)

    # ct_1_units = get_units(units["children"][2]['objects'])
    # ct_1_names = get_names(units["children"][2]['name'], ct_1_units)
    
    # ct_10_units = get_units(units["children"][3]['objects'])
    # ct_10_names = get_names(units["children"][3]['name'], ct_10_units)
    
    # ct_11_units = get_units(units["children"][4]['objects'])
    # ct_11_names = get_names(units["children"][4]['name'], ct_11_units)
    
    # ct_14_units = get_units(units["children"][5]['objects'])
    # ct_14_names = get_names(units["children"][5]['name'], ct_14_units)
    
    # ct_14_units = get_units(units["children"][5]['objects'])
    # ct_14_names = get_names(units["children"][5]['name'], ct_14_units)
    
    # ct_16_units = get_units(units["children"][6]['objects'])
    # ct_16_names = get_names(units["children"][6]['name'], ct_16_units)
    
    # ct_17_units = get_units(units["children"][7]['objects'])
    # ct_17_names = get_names(units["children"][7]['name'], ct_17_units)
    
    # ct_18_units = get_units(units["children"][8]['objects'])
    # ct_18_names = get_names(units["children"][8]['name'], ct_18_units)
    
    # ct_2_units = get_units(units["children"][10]['objects'])
    # ct_2_names = get_names(units["children"][10]['name'], ct_2_units)
    
    # ct_22_units = get_units(units["children"][11]['objects'])
    # ct_22_names = get_names(units["children"][11]['name'], ct_22_units)
    
    # ct_3_units = get_units(units["children"][12]['objects'])
    # ct_3_names = get_names(units["children"][12]['name'], ct_3_units)
    
    # ct_31_units = get_units(units["children"][13]['objects'])
    # ct_31_names = get_names(units["children"][13]['name'], ct_31_units)
    
    # ct_4_units = get_units(units["children"][14]['objects'])
    # ct_4_names = get_names(units["children"][14]['name'], ct_4_units)
    
    # ct_6_units = get_units(units["children"][16]['objects'])
    # ct_6_names = get_names(units["children"][16]['name'], ct_6_units)
    
    # ct_7_units = get_units(units["children"][17]['objects'])
    # ct_7_names = get_names(units["children"][17]['name'], ct_7_units)
    
    # ct_8_units = get_units(units["children"][18]['objects'])
    # ct_8_names = get_names(units["children"][18]['name'], ct_8_units)

    # ct_9_units = get_units(units["children"][19]['objects'])
    # ct_9_names = get_names(units["children"][19]['name'], ct_9_units)
    
    # ct_reserve_units = get_units(units["children"][20]['objects'])
    # ct_reserve_names = get_names(units["children"][20]['name'], ct_reserve_units)
    
    # frac_1_units = get_units(units["children"][21]['objects'])
    # frac_1_names = get_names(units["children"][21]['name'], frac_1_units)
    
    # frac_14_units = get_units(units["children"][22]['objects'])
    # frac_14_names = get_names(units["children"][22]['name'], frac_14_units)
    
    # frac_15_units = get_units(units["children"][23]['objects'])
    # frac_15_names = get_names(units["children"][23]['name'], frac_15_units)
    
    # frac_16_units = get_units(units["children"][24]['objects'])
    # frac_16_names = get_names(units["children"][24]['name'], frac_16_units)
    
    # frac_17_units = get_units(units["children"][25]['objects'])
    # frac_17_names = get_names(units["children"][25]['name'], frac_17_units)
    
    # frac_2_units = get_units(units["children"][26]['objects'])
    # frac_2_names = get_names(units["children"][26]['name'], frac_2_units)
    
    # frac_3_units = get_units(units["children"][27]['objects'])
    # frac_3_names = get_names(units["children"][27]['name'], frac_3_units)
    
    # frac_4_units = get_units(units["children"][28]['objects'])
    # frac_4_names = get_names(units["children"][28]['name'], frac_4_units)
    
    # frac_5_units = get_units(units["children"][29]['objects'])
    # frac_5_names = get_names(units["children"][29]['name'], frac_5_units)
    
    # frac_6_units = get_units(units["children"][30]['objects'])
    # frac_6_names = get_names(units["children"][30]['name'], frac_6_units)
    
    # frac_7_units = get_units(units["children"][31]['objects'])
    # frac_7_names = get_names(units["children"][31]['name'], frac_7_units)
    
    # frac_8_units = get_units(units["children"][32]['objects'])
    # frac_8_names = get_names(units["children"][32]['name'], frac_8_units)
    
    # frac_9_units = get_units(units["children"][34]['objects'])
    # frac_9_names = get_names(units["children"][34]['name'], frac_9_units)
    
    # frac_reserve_units = get_units(units["children"][35]['objects'])
    # frac_reserve_names = get_names(units["children"][35]['name'], frac_reserve_units)
    
    # diesel_units = get_units(units["children"][36]['objects'])
    # diesel_names = get_names(units["children"][36]['name'], diesel_units)

    # drvidentification_units = get_units(units["children"][57]['objects'])
    # drvidentification_names = get_names(units["children"][57]['name'], drvidentification_units)
    
    # maz_units = get_units(units["children"][58]['objects'])
    # maz_names = get_names(units["children"][58]['name'], maz_units)
    
    # NEFTEMASH_units = get_units(units["children"][59]['objects'])
    # NEFTEMASH_names = get_names(units["children"][59]['name'], NEFTEMASH_units)
    
    # Neftemash_units = get_units(units["children"][60]['objects'])
    # Neftemash_names = get_names(units["children"][60]['name'], Neftemash_units)
    # pprint(uts_atz_ural_units)
    # pprint(uts_atz_ural_names)
    # pprint(len(uts_atz_ural_units))
    # pprint(len(uts_atz_ural_names))
    # check_units = get_units(units["children"][61]['children'][0]['objects'])
    # check_names = get_names(units["children"][61]['name'], ct_14_units)
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