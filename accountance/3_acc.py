# MATCH CLEAN DATA WITH OMNICOMM.DB
# Setup, imports, connections
from multiprocessing.sharedctypes import Value
from optparse import Values
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

# Excel connection  
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\acc.xls")
ws1 = wb.Worksheets(1)

# Pandas
pd.set_option('display.max_rows', None)

# db connections
db_om = sqlite3.connect('omnicomm.db')
db_om.row_factory = lambda cursor, row: row[0]
cursor_om = db_om.cursor()

db_acc = sqlite3.connect('accountance.db')
db_acc.row_factory = lambda cursor, row: row[0]
cursor_acc = db_acc.cursor()


def main():
    # Connect to omnicomm.db get Final_DB as L_om_index_plates
    L_om_plates = cursor_om.execute("SELECT Plate_index FROM Final_DB").fetchall()
    
    # Connect to accountance.db get accountance_2 (mols, items) as L_mols and L_acc_index_plates
    L_acc_mols = cursor_acc.execute("SELECT Responsible FROM accountance_2").fetchall()
    L_acc_plates = cursor_acc.execute("SELECT Plate_index FROM accountance_2").fetchall()
    
    # Get lists L_mols and L_acc_index_plates into dict
    D_mols_plates = dict(zip(L_acc_plates, L_acc_mols))
   
    # Match if Omnicomm items are in accountance
    L_mols_matched = []
    for i in L_om_plates:
        L_mols_matched.append(D_mols_plates.get(i))
   
    # Push L_mols_matched into final_DB as Responsible
    # Destructure final_DB into lists
    L_dept = cursor_om.execute("SELECT Department FROM final_DB").fetchall()
    L_vehicle = cursor_om.execute("SELECT Vehicle FROM final_DB").fetchall()
    L_plate = cursor_om.execute("SELECT Plate FROM final_DB").fetchall()
    L_vehicle_name = cursor_om.execute("SELECT Vehicle_name FROM final_DB").fetchall()
    L_plate_index = cursor_om.execute("SELECT Plate_index FROM final_DB").fetchall()
    L_loc_om = cursor_om.execute("SELECT Location_omnicomm FROM final_DB").fetchall()
    L_loc_cits = cursor_om.execute("SELECT Location_cits FROM final_DB").fetchall()
    L_nodata = cursor_om.execute("SELECT No_data FROM final_DB").fetchall()
    
    # Push all lists with L_mols_matched as Responsible
    
    cursor_om.execute("DROP TABLE IF EXISTS final_DB;")
    cursor_om.execute("""
        CREATE TABLE IF NOT EXISTS final_DB(
        Department text,
        Vehicle text, 
        Plate text, 
        Vehicle_name text, 
        Plate_index text, 
        Location_omnicomm text, 
        Location_cits text, 
        No_data text, 
        Responsible text)
                """)
    cursor_om.executemany("INSERT INTO final_DB VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)", zip(L_dept, L_vehicle, L_plate, 
                                                                                               L_vehicle_name, L_plate_index, L_loc_om,
                                                                                               L_loc_cits, L_nodata, L_mols_matched))
                                                                                            
    
    db_om.commit()
    db_om.close()

    # Match if accountance items are not in Omnicomm
    L_plates_unmatched = []
    L_mols_unmatched = []
    for k in D_mols_plates.keys():
        if k not in L_om_plates:
            L_plates_unmatched.append(k)
            L_mols_unmatched.append(D_mols_plates.get(i))
    
    # Get df of unmatched items
    data = pd.DataFrame(zip(cursor_acc.execute("SELECT Responsible FROM accountance_2").fetchall(), 
                            cursor_acc.execute("SELECT Unit FROM accountance_2").fetchall(), 
                            cursor_acc.execute("SELECT Plate_index FROM accountance_2").fetchall()), 
                            columns=['Responsible', 'Unit', 'Plate_index'])
    
    # Filter df for unmatched items
    data = data.loc[data.Plate_index.isin(L_plates_unmatched)]
    print(data)
    
    # Destructuring data into lists
    L_mols = data.loc[:, 'Responsible'].tolist()
    L_units = data.loc[:, 'Unit'].tolist()
    L_plates = data.loc[:, 'Plate_index'].tolist()

    # Push the lists to db
    cursor_acc.execute("DROP TABLE IF EXISTS acc_unmatched;")
    cursor_acc.execute("""
        CREATE TABLE IF NOT EXISTS acc_unmatched(
        Responsible text,
        Unit text, 
        Plate_index text)
                """)
    cursor_acc.executemany("INSERT INTO acc_unmatched VALUES (?, ?, ?)", zip(L_mols, L_units, L_plates))

    db_acc.commit()
    db_acc.close()
   
    
    
if __name__ == "__main__":
    main()