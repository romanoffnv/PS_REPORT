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

 # connection to match.db
db_match = sqlite3.connect('match.db')
db_match.row_factory = lambda cursor, row: row[0]
cursor = db_match.cursor()

cnx_cits = sqlite3.connect('cits.db')
cnx_om = sqlite3.connect('omnicomm.db')
cnx_match = sqlite3.connect('match.db')

def main():
    # Get from cits.db
    df_cits = pd.read_sql_query("SELECT * FROM Final_cits", cnx_cits)
    df_om = pd.read_sql_query("SELECT * FROM Final_DB", cnx_om)
    # pprint(df_cits)

    # Destructuring cits df
    def cits_destructurer():
        L0 = df_cits['Crews'].tolist()
        L1 = df_cits['Units'].tolist()
        L2 = df_cits['Plates'].tolist()
        L3 = df_cits['Plate_index'].tolist()
        L4 = df_cits['Locations'].tolist()
        return L0, L1, L2, L3, L4
    
    L_all_cits = cits_destructurer()
    
    # Destructuring omnicomm df 
    def om_destructurer():
        L0 = df_om['Department'].tolist()
        L1 = df_om['Vehicle'].tolist()
        L2 = df_om['Plate'].tolist()
        L3 = df_om['Plate_index'].tolist()
        L4 = df_om['Location_Omnicomm'].tolist()
        L5 = df_om['No_data'].tolist()
        return L0, L1, L2, L3, L4, L5
    
    L_all_om = om_destructurer()
    
    # Get matched plates
    def matcher(L_values):
        D = dict(zip(L_all_cits[3], L_values))
        L = []
        for i in L_all_om[3]:
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
                        L_all_om[0],
                        L_all_om[1],
                        L_all_om[2],
                        L_all_om[3],
                        L_all_om[4],
                        L_all_om[5],
                        
                        # Cits
                        L_crews_cits,
                        L_units_cits,
                        L_Plates_cits,
                        L_Plate_index_cits,
                        L_Locations_cits), 
                        columns= [
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
                        'Locs_ct',]
    )
    
    # pprint(df_matched)
    # Get unmatched plates
    def dismatcher(L_values):
        D = dict(zip(L_all_cits[3], L_values))
        
        L = []
        for k, v in D.items():
            if k not in L_all_om[3]:
                L.append(v)
                
                
        return L
    
    
    L_crews_cits = dismatcher(L_all_cits[0])
    L_units_cits = dismatcher(L_all_cits[1])
    L_Plates_cits = dismatcher(L_all_cits[2])
    L_Plate_index_cits = dismatcher(L_all_cits[3])
    L_Locations_cits = dismatcher(L_all_cits[4])
   
    # Blanking out omnicomm cols for unmatched items by the length of Crew col
    for i in L_all_om:
        i.clear() 
    
    for i in range(0, len(L_crews_cits)):
        L_all_om[0].append('-')
        L_all_om[1].append('-')
        L_all_om[2].append('-')
        L_all_om[3].append('-')
        L_all_om[4].append('-')
        L_all_om[5].append('-')
        
    
   
    df_unmatched = pd.DataFrame(zip(
                        # Omnicomm
                        L_all_om[0],
                        L_all_om[1],
                        L_all_om[2],
                        L_all_om[3],
                        L_all_om[4],
                        L_all_om[5],
                        
                        # Cits
                        L_crews_cits,
                        L_units_cits,
                        L_Plates_cits,
                        L_Plate_index_cits,
                        L_Locations_cits), 
                        columns= [
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
                        'Locs_ct',]
    )
    
    
    df_total = pd.concat([df_matched, df_unmatched])
        
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS om_cits")
    df_total.to_sql(name='om_cits', con=db_match, if_exists='replace', index=False)
    db_match.commit()
    db_match.close()

    df_match = pd.read_sql_query("SELECT * FROM om_cits", cnx_match)
    pprint(df_match)
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))