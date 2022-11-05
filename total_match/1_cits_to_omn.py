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
    df_om = pd.read_sql_query("SELECT * FROM Final_om", cnx_om)
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
        L0 = df_om['Groups'].tolist()
        L1 = df_om['Units'].tolist()
        L2 = df_om['id'].tolist()
        L3 = df_om['Status'].tolist()
        L4 = df_om['Plates'].tolist()
        L5 = df_om['PI'].tolist()
        L6 = df_om['Locations'].tolist()
        return L0, L1, L2, L3, L4, L5, L6
    
    L_all_om = om_destructurer()
    
    
    # Get matched plates
    def matcher(L_values):
        D = dict(zip(L_all_cits[3], L_values))
        L = []
        L_mm = []
        for i in L_all_om[5]:
            if i in D.keys():
                L.append(D.get(i))
                L_mm.append(D.get(i))
            else:
                L.append('-')
            
        return L
        # return L_mm
    
    L_crews_cits = matcher(L_all_cits[0])
    L_units_cits = matcher(L_all_cits[1])
    L_Plates_cits = matcher(L_all_cits[2])
    L_Locations_cits = matcher(L_all_cits[4])
    
    
    L_matched =  matcher(L_all_cits[3])
   
    df_matched = pd.DataFrame(zip(
                        # Omnicomm
                        L_all_om[0],
                        L_all_om[1],
                        L_all_om[2],
                        L_all_om[3],
                        L_all_om[4],
                        L_all_om[5],
                        L_all_om[6],
                        
                        # Cits
                        L_crews_cits,
                        L_units_cits,
                        L_Plates_cits,
                        L_Locations_cits), 
                        columns= [
                        # Omnicomm
                        'Groups_om', 
                        'Units_om',
                        'id_om',
                        'Status',
                        'Plates',
                        'PI_om',
                        'Locations_om',
                        # Cits
                        'Crews_ct',
                        'Units_ct',
                        'Plates_ct',
                        'Locs_ct',]
    )
    
    # pprint(df_matched)
    # Get unmatched plates
    def dismatcher(L_values):
        D = dict(zip(L_all_cits[3], L_values))
        
        L = []
        for k, v in D.items():
            if k not in L_all_om[5]:
                L.append(v)
                
                
        return L
    
    
    L_crews_cits = dismatcher(L_all_cits[0])
    L_units_cits = dismatcher(L_all_cits[1])
    L_Plates_cits = dismatcher(L_all_cits[2])
    L_Locations_cits = dismatcher(L_all_cits[4])
    pprint(f"this is the length of all ct units {len(L_all_cits[0])}")
    pprint(f"this is the length of matched items {len(L_matched)}")
    pprint(f"this is the length of calculated dismatched units {len(L_all_cits[0]) - len(L_matched)}")
    pprint(f"this is the length of unmatched units from the func {len(L_crews_cits)}")
    
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
        L_all_om[6].append('-')
        
    
   
    df_unmatched = pd.DataFrame(zip(
                        # Omnicomm
                        L_all_om[0],
                        L_all_om[1],
                        L_all_om[2],
                        L_all_om[3],
                        L_all_om[4],
                        L_all_om[5],
                        L_all_om[6],
                        
                        # Cits
                        L_crews_cits,
                        L_units_cits,
                        L_Plates_cits,
                        L_Locations_cits), 
                        columns= [
                        # Omnicomm
                        # Omnicomm
                        'Groups_om', 
                        'Units_om',
                        'id_om',
                        'Status',
                        'Plates',
                        'PI_om',
                        'Locations_om',
                        # Cits
                        'Crews_ct',
                        'Units_ct',
                        'Plates_ct',
                        'Locs_ct',]
    )
    
    
    
    pprint(df_matched)
    pprint(df_unmatched)
    df_total = pd.merge(df_matched, df_unmatched, how="outer")  
    # pprint(df_total)
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS om_cits")
    df_total.to_sql(name='om_cits', con=db_match, if_exists='replace', index=False)
    db_match.commit()
    db_match.close()

    writer = pd.ExcelWriter('DB.xlsx', engine='xlsxwriter')
    # Write each dataframe to a different worksheet.
    # data.index += 1
    df_total.to_excel(writer, index = True, header=True)
    writer.save()

if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))