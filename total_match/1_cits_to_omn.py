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

    # Listing cits df
    
    L_crews_cits = df_cits['Crews'].tolist()
    L_units_cits = df_cits['Units'].tolist()
    L_Plates_cits = df_cits['Plates'].tolist()
    L_Plate_index_cits = df_cits['Plate_index'].tolist()
    L_Locations_cits = df_cits['Locations'].tolist()

    # Listing om 
    L_Plate_dept_om = df_om['Department'].tolist()
    L_Plate_index_om = df_om['Plate_index'].tolist()
    L_Plate_unit_om = df_om['Vehicle'].tolist()
    
    L_matched_PInd = []
    # Get matched plates
    def matcher(L_values):
        L_Plate_index_om = df_om['Plate_index'].tolist()
        L_Plate_index_cits = df_cits['Plate_index'].tolist()
        
        D = dict(zip(L_Plate_index_cits, L_values))
        L = []
        for i in L_Plate_index_om:
            if i in D.keys():
                L.append(D.get(i))
            else:
                L.append('-')
        return L
    L_matched_crews = matcher(L_crews_cits)
    L_matched_units = matcher(L_units_cits)
    L_matched_plates = matcher(L_Plates_cits)
    L_matched_Locations = matcher(L_Locations_cits)
    
    df = pd.DataFrame(zip(L_Plate_dept_om, L_Plate_unit_om, L_Plate_index_om, L_matched_crews, L_matched_units, L_matched_plates, L_matched_Locations))
    print(df)
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