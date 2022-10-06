import xlsxwriter
import pandas as pd
import sqlite3
import re
import os
from pprint import pprint

# Connect to the database
db = sqlite3.connect('omnicomm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

def main():
    # pprint(len(cursor.execute("SELECT * FROM total_vehicles_nobubs_nodata").fetchall()))
    # Extracting table columns from total_vehicles_nobubs
    L_total_dep = cursor.execute("SELECT Department FROM total_vehicles_nobubs_nodata").fetchall()
    L_total_veh = cursor.execute("SELECT Vehicle FROM total_vehicles_nobubs_nodata").fetchall()
    L_total_nodata = cursor.execute("SELECT Nodata FROM total_vehicles_nobubs_nodata").fetchall()
    
    # Slicing vehicles column(list) into plate, index, literal cols
    plates1 = re.compile("[А-Яа-я]*\d+[А-Яа-я]{2}\s*\d+")
    plates2 = re.compile("[А-Яа-я]{2}\d+\s\d+")
    plates3 = re.compile("\w{2}\s\D\d+\s\d{2}")
    plates4 = re.compile("\W\d+\s\d+\.\d+")
    pattern = re.compile("(\гос.\s*№)|(\s\W{2}\d+\s\d+)|(\гос. №)|(kz\W{2}\d+\s\d+)")
    

    # Derivating plates from vehicles
    L_total_plate = [''.join(re.findall(plates1, x)) or 
                     ''.join(re.findall(plates2, x)) or 
                     ''.join(re.findall(plates3, x)) or 
                     ''.join(re.findall(plates4, x)) for x in L_total_veh]

    # Derivating indeces from plates
    # Stripping regions
    L_total_index = [x.removesuffix('186').strip() if x != None else x for x in L_total_plate]
    L_total_index = [x.removesuffix('86').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('116').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('82').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('89').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('156').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('56').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('797').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('54').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('77').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('07').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('126').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('78').strip() if x != None else x for x in L_total_index]
    L_total_index = [x.removesuffix('23').strip() if x != None else x for x in L_total_index]
    
    # Stripping literals
    L_total_literal = [''.join(re.findall(r'\D', x)).lower() for x in L_total_index]
    L_total_literal = [''.join(re.sub(r'\s+', '', x)).lower() for x in L_total_literal]
    L_total_index = [''.join(re.findall(r'\d+', x)) for x in  L_total_index]
    
    L_temp = [x + y for x, y in zip(L_total_index, L_total_literal)]
    L_total_index = [x for x in L_temp]

    # Cleaning vehicles names
    L_veh_name = [''.join(re.sub(plates1, '', x)).strip() or ''.join(re.sub(pattern, '', x)).strip() for x in L_total_veh]
    L_veh_name = [str(x).replace('  ', ' ') for x in L_veh_name]
    rem = re.compile(r'\D{2}\d+\s\d+')
    L_veh_name = [''.join(re.sub(rem, '', x)).strip() for x in L_veh_name]
    
    
    data = pd.DataFrame(zip(L_total_dep, L_total_veh, L_total_plate, L_veh_name, 
                            L_total_index, L_total_nodata), 
                        columns =['Department', 'Vehicle', 'Plate', 'Vehicle_name', 
                                  'Plate_Index', 'No_data'])
    
    # print(data)
    
    
    # Posting dataframe back into the sql database
    cursor.execute("DROP TABLE IF EXISTS final_DB")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS final_DB(
        Department text,
        Vehicle text,
        Plate text,
        Vehicle_name text,
        Plate_index text,
        No_data text
        )
        """)

    data.to_sql('final_DB', db, if_exists='replace', index = False)


    db.commit()
    db.close()
    print("5_pre_final is complete")
    
if __name__ == '__main__':
    main()
    