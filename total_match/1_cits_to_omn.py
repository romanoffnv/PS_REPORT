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
db_cits = sqlite3.connect('cits.db')
db_cits.row_factory = lambda cursor, row: row[0]
cursor_cits = db_cits.cursor()

# connection to omnicomm.db
db_om = sqlite3.connect('omnicomm.db')
db_om.row_factory = lambda cursor, row: row[0]
cursor_om = db_om.cursor()

cnx_cits = sqlite3.connect('cits.db')
cnx_om = sqlite3.connect('omnicomm.db')

def main():
    # Get from cits.db
    df_cits = pd.read_sql_query("SELECT * FROM Final_cits", cnx_cits)
    df_om = pd.read_sql_query("SELECT * FROM Final_DB", cnx_om)
    # pprint(df_cits)

    # Destructuring cits df
    def cits_destructurer():
        L1 = df_cits['Crews'].tolist()
        L2 = df_cits['Units'].tolist()
        L3 = df_cits['Plates'].tolist()
        L4 = df_cits['Plate_index'].tolist()
        L5 = df_cits['Locations'].tolist()
        return L1, L2, L3, L4, L5
    
    L_all_cits = cits_destructurer()
    
    # Destructuring omnicomm df 
    L_Plate_dept_om = df_om['Department'].tolist()
    L_units_om = df_om['Vehicle'].tolist()
    L_Plates = df_om['Plate'].tolist()
    L_Plate_index_om = df_om['Plate_index'].tolist()
    L_Locs_om = df_om['Location_Omnicomm'].tolist()
    L_nodata_om = df_om['No_data'].tolist()
    
    
    # Get matched plates
    def matcher(L_values):
        D = dict(zip(L_all_cits[3], L_values))
        L = []
        for i in L_Plate_index_om:
            if i in D.keys():
                L.append(D.get(i))
            else:
                L.append('-')
        return L
    L_crews_cits = matcher(L_all_cits[0])
    L_units_cits = matcher(L_all_cits[1])
    L_Plates_cits = matcher(L_all_cits[2])
    L_Plate_index_cits = matcher(L_all_cits[3])
    L_Locations_cits = matcher(L_all_cits[4])
    
    df_matched = pd.DataFrame(zip(
                        # Omnicomm
                        L_Plate_dept_om,
                        L_units_om,
                        L_Plates,
                        L_Plate_index_om,
                        L_Locs_om,
                        L_nodata_om,
                        # Cits
                        L_crews_cits,
                        L_units_cits,
                        L_Plates_cits,
                        L_Plate_index_cits,
                        L_Locations_cits), columns= [
                        # Omnicomm
                        'Group', 
                        'Units_om',
                        'Plates_om',
                        'PI_om',
                        'Locs_om',
                        'No_data',
                        # Cits
                        'Crews_ct',
                        'Units_ct',
                        'Plates_ct',
                        'PI_ct',
                        'Locs_ct',
                        
                        ])
    
    
    # Get unmatched plates
    def dismatcher(L_values):
        D = dict(zip(L_all_cits[3], L_values))
        L = []
        for k, v in D.items():
            if k not in L_Plate_index_om:
                L.append(v)
                
        return L
    
    
    L_crews_cits = dismatcher(L_all_cits[0])
    L_units_cits = dismatcher(L_all_cits[1])
    L_Plates_cits = dismatcher(L_all_cits[2])
    L_Plate_index_cits = dismatcher(L_all_cits[3])
    L_Locations_cits = dismatcher(L_all_cits[4])
    
    # Blanking out omnicomm cols for unmatched items by the length of Crew col
    for i in L_crews_cits:
        L_Plate_dept_om.append('-')
    pprint(len(L_Plate_dept_om))
    
    # Omnicomm
                        # L_Plate_dept_om,
                        # L_units_om,
                        # L_Plates,
                        # L_Plate_index_om,
                        # L_Locs_om,
                        # L_nodata_om,
    
    # df_unmatched = pd.DataFrame(zip(L_crews_cits, L_units_cits, L_Plates_cits, L_Plate_index_cits, L_Locations_cits))
    
    # L_unmatched_PInd, L_unmatched_Crews = [], []
    # # Get umatched plates
    # for k, v in D.items():
    #     if k in L_Plate_index_om:
    #         pass
    #     else:
    #         L_unmatched_PInd.append(k)
    #         L_unmatched_Crews.append(v)
    
    
    # pprint(len(L_unmatched_PInd))
    # pprint(len(L_unmatched_Crews))
    # pprint(L_matched)
    # pprint(L_unmatched)
    # pprint(len(L_matched))
    # pprint(len(L_unmatched))
    # for k, v in D.items():
    #     if k in L_Plate_index_om:
    #         L_matched.append(v)
    #     else:
    #         L_matched.append('-')
    #         L_unmatched.append(k)
    
    # pprint(L_matched)
    # pprint(L_unmatched)
    # pprint(len(L_matched))
    # pprint(len(L_unmatched))
    # pprint(L_crews)
    # Matching multiple columns via dict
    # def dict_matcher(x):
    #         L = []
    #         D = dict(zip(L_total_frac_ind, x))
    #         for i in L_total_plates_ind:
    #             L.append(D.get(i))
    #         return L
        
    # L_group_matched = dict_matcher(L_frac_group)
    # L_unit_matched = dict_matcher(L_frac_unit)    
    # L_plates_matched = dict_matcher(L_frac_plates)    
    # L_mols_matched = dict_matcher(L_frac_mols)    
    # L_drivers_matched = dict_matcher(L_frac_drivers)    
    # L_discrepancies_matched = dict_matcher(L_frac_discrepancies)    
    # L_notes_matched = dict_matcher(L_frac_notes) 



if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))