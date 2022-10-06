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

# Pandas
pd.set_option('display.max_rows', None)

# db connections
db_om = sqlite3.connect('omnicomm.db')
db_om.row_factory = lambda cursor, row: row[0]
cursor_om = db_om.cursor()

db_cits = sqlite3.connect('cits.db')
db_cits.row_factory = lambda cursor, row: row[0]
cursor_cits = db_cits.cursor()

def main():
    # Destrucuring omnincomm db
    L_dept_om = cursor_om.execute("SELECT Department FROM Final_DB").fetchall()
    L_vehicle_om = cursor_om.execute("SELECT Vehicle FROM Final_DB").fetchall()
    L_plate_om = cursor_om.execute("SELECT Plate FROM Final_DB").fetchall()
    L_veh_name_om = cursor_om.execute("SELECT Vehicle_name FROM Final_DB").fetchall()
    L_plate_index_om = cursor_om.execute("SELECT Plate_index FROM Final_DB").fetchall()
    L_loc_om_om = cursor_om.execute("SELECT Location_omnicomm FROM Final_DB").fetchall()
    L_loc_cits_om = cursor_om.execute("SELECT Location_cits FROM Final_DB").fetchall()
    L_nodata_om = cursor_om.execute("SELECT No_data FROM Final_DB").fetchall()
    L_mols_om = cursor_om.execute("SELECT Responsible FROM Final_DB").fetchall()

    # Destructuring cits db
    L_dept_cits = cursor_cits.execute("SELECT Dept FROM Trucks_matched").fetchall()
    L_unit_cits = cursor_cits.execute("SELECT Unit FROM Trucks_matched").fetchall()
    L_plate_cits = cursor_cits.execute("SELECT Plate FROM Trucks_matched").fetchall()

    # building df
    cursor_om.execute("DROP TABLE IF EXISTS final2_DB")
    cursor_om.execute("""
	CREATE TABLE IF NOT EXISTS final2_DB(
	Department text,
	Vehicle text,
    Plate text, 
    Vehicle_name text,
    Plate_index text,
    Location_omnicomm text,
    Location_cits text,
    No_data text,
    Responsible text,
    CT_crew text,
    CT_unit text,
    CT_plate text)
              """)
    cursor_om.executemany("INSERT INTO final2_DB VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", zip(L_dept_om, L_vehicle_om, 
    L_plate_om, L_veh_name_om, L_plate_index_om, L_loc_om_om, L_loc_cits_om, L_nodata_om, L_mols_om, L_dept_cits, L_unit_cits, L_plate_cits))

    db_om.commit()
    db_om.close()

    df = pd.DataFrame(zip(L_dept_om, L_vehicle_om, L_plate_om, L_loc_om_om, L_loc_cits_om, L_nodata_om, 
    L_mols_om, L_dept_cits, L_unit_cits, L_plate_cits), columns=['Службы по Омникомм', 'СПТ по Омникомм', '№ по Омникомм', 
    'Локация по Омникомм', 'Локация по сводкам', 'Данные по Омникомм', 'МОЛ по бухгалтерии', 'Службы по сводкам', 'СПТ по сводкам', '№ по сводкам'])
    print(df)
    
     # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('DB.xlsx', engine='xlsxwriter')
    # Write each dataframe to a different worksheet.
    df.index += 1
    df.to_excel(writer, index = True, header=True)
    writer.save()

    
    
    
    
if __name__ == '__main__':
    main()