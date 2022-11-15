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
# connection to brm.db
db_brm = sqlite3.connect('brm.db')
db_brm.row_factory = lambda cursor, row: row[0]
cursor_brm = db_brm.cursor()

# connection to match.db
db_match = sqlite3.connect('match.db')
db_match.row_factory = lambda cursor, row: row[0]
cursor = db_match.cursor()

cnx_brm = sqlite3.connect('brm.db')
cnx_match = sqlite3.connect('match.db')

def main():
    df_brm = pd.read_sql_query("SELECT * FROM Units_locs", cnx_brm)
    df_match = pd.read_sql_query("SELECT * FROM cits_to_om", cnx_match)
    
    # Destructuring cits df
    def cits_destructurer():
        L0 = df_brm['Crews'].tolist()
        L1 = df_brm['Units'].tolist()
        L2 = df_brm['Plates'].tolist()
        L3 = df_brm['Plate_index'].tolist()
        L4 = df_brm['Locs'].tolist()
        return L0, L1, L2, L3, L4
    
    L_all_brm = cits_destructurer()
    
    # Destructuring match df 
    def match_destructurer():
        L0 = df_match['Groups_om'].tolist()
        L1 = df_match['Units_om'].tolist()
        L2 = df_match['id_om'].tolist()
        L3 = df_match['Status'].tolist()
        L4 = df_match['Plates'].tolist()
        L5 = df_match['PI_om'].tolist()
        L6 = df_match['Locations_om'].tolist()
        L7 = df_match['Crews_ct'].tolist()
        L8 = df_match['Units_ct'].tolist()
        L9 = df_match['Plates_ct'].tolist()
        L10 = df_match['Locs_ct'].tolist()
        return L0, L1, L2, L3, L4, L5, L6, L7, L8, L9, L10
    
    L_all_match = match_destructurer()

    # Get matched plates
    def matcher(L_values):
        D = dict(zip(L_all_brm[3], L_values))
        L = []
        L_mm = []
        for i in L_all_match[5]:
            if i in D.keys():
                L.append(D.get(i))
                L_mm.append(D.get(i))
            else:
                L.append('-')
            
        return L
        # return L_mm
    
    L_crews_brm = matcher(L_all_brm[0])
    L_units_brm = matcher(L_all_brm[1])
    L_Plates_brm = matcher(L_all_brm[2])
    L_Locations_brm = matcher(L_all_brm[4])

    df_matched = pd.DataFrame(zip(
                        # Omnicomm
                        L_all_match[0],
                        L_all_match[1],
                        L_all_match[2],
                        L_all_match[3],
                        L_all_match[4],
                        L_all_match[5],
                        L_all_match[6],
                        L_all_match[7],
                        L_all_match[8],
                        L_all_match[9],
                        L_all_match[10],
                        
                        # Brm
                        L_crews_brm,
                        L_units_brm,
                        L_Plates_brm,
                        L_Locations_brm), 
                        columns= [
                        # Match om
                        'Groups_om', 
                        'Units_om',
                        'id_om',
                        'Status_om',
                        'Plates_om',
                        'PI_om',
                        'Locs_om',
                        # Match ct
                        'Crews_ct',
                        'Units_ct',
                        'Plates_ct',
                        'Locs_ct',
                        # Brm
                        'Crews_brm',
                        'Units_brm',
                        'Plates_brm',
                        'Locs_brm',]
    )
    
    # Get unmatched plates
    def dismatcher(L_values):
        D = dict(zip(L_all_brm[3], L_values))
        
        L = []
        for k, v in D.items():
            if k not in L_all_match[5]:
                L.append(v)
                
                
        return L
    
    
    L_crews_brm = dismatcher(L_all_brm[0])
    L_units_brm = dismatcher(L_all_brm[1])
    L_Plates_brm = dismatcher(L_all_brm[2])
    L_Locations_brm = dismatcher(L_all_brm[4])
   
    
    # Blanking out omnicomm cols for unmatched items by the length of Crew col
    for i in L_all_match:
        i.clear() 
    
    for i in range(0, len(L_crews_brm)):
        L_all_match[0].append('-')
        L_all_match[1].append('-')
        L_all_match[2].append('-')
        L_all_match[3].append('-')
        L_all_match[4].append('-')
        L_all_match[5].append('-')
        L_all_match[6].append('-')
        L_all_match[7].append('-')
        L_all_match[8].append('-')
        L_all_match[9].append('-')
        L_all_match[10].append('-')

    df_unmatched = pd.DataFrame(zip(
                        # Omnicomm
                        L_all_match[0],
                        L_all_match[1],
                        L_all_match[2],
                        L_all_match[3],
                        L_all_match[4],
                        L_all_match[5],
                        L_all_match[6],
                        L_all_match[7],
                        L_all_match[8],
                        L_all_match[9],
                        L_all_match[10],
                        
                        # Brm
                        L_crews_brm,
                        L_units_brm,
                        L_Plates_brm,
                        L_Locations_brm), 
                        columns= [
                        # Match om
                        'Groups_om', 
                        'Units_om',
                        'id_om',
                        'Status_om',
                        'Plates_om',
                        'PI_om',
                        'Loctions_om',
                        # Match ct
                        'Crews_ct',
                        'Units_ct',
                        'Plates_ct',
                        'Locs_ct',
                        # Brm
                        'Crews_brm',
                        'Units_brm',
                        'Plates_brm',
                        'Locs_brm',]
    )
    
    df_total = pd.merge(df_matched, df_unmatched, how="outer")  
    pprint(df_total)
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS brm_to_om")
    df_total.to_sql(name='brm_to_om', con=db_match, if_exists='replace', index=False)
    db_match.commit()
    db_match.close()
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))