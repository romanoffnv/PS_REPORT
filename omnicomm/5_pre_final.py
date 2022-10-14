import json
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
    # pprint(L_total_veh)
    
    # Fishing out plates by regex from long sentences
    # Slicing vehicles column(list) into plate, index, literal cols
    plates1 = re.compile("[А-Яа-я]*\d+[А-Яа-я]{2}\s*\d+")
    plates2 = re.compile("[А-Яа-я]{2}\d+\s\d+")
    plates3 = re.compile("\w{2}\s\D\d+\s\d{2}")
    plates4 = re.compile("\W\d+\s\d+\.\d+")
    plates5 = re.compile("\ДЭС.*")
    plates6 = re.compile("\дэс.*")
    plates7 = re.compile("\D{2}\s\d+\s\d+")
    
    
    # Derivating plates from vehicles
    L_total_plate = [''.join(re.findall(plates1, x)) or 
                     ''.join(re.findall(plates2, x)) or 
                     ''.join(re.findall(plates3, x)) or 
                     ''.join(re.findall(plates4, x)) or
                     ''.join(re.findall(plates5, x)) or
                     ''.join(re.findall(plates6, x)) or
                     ''.join(re.findall(plates7, x)) for x in L_total_veh]


    pprint(L_total_plate)

    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions = [186, 86, 797, 116, '02', '07',89, 82, 78, 54, 77, 126, 188, 88, 174, 74, 158, 196, 156, 56, 76, 23]
        
        for i in L_regions:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) > 7 else x for x in plates]
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates
    

    L_total_plates_ind = transform_plates(L_total_plate) 
    
    D_om_diesels = json.load(open('D_om_diesels.json'))
    for k, v in D_om_diesels.items():
        L_total_plates_ind = [''.join(x.replace(k, v)).strip() for x in L_total_plates_ind]
    
    # pprint(L_total_plates_ind)
    # pprint(len(L_total_plates_ind))

    
    
    # Posting dataframe back into the sql database
    cursor.execute("DROP TABLE IF EXISTS final_DB")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS final_DB(
        Department text,
        Vehicle text,
        Plate text,
        Plate_index text,
        No_data text
        )
        """)

    cursor.executemany("INSERT INTO final_DB VALUES (?, ?, ?, ?, ?)", zip(L_total_dep, L_total_veh, L_total_plate, L_total_plates_ind, L_total_nodata))



    db.commit()
    db.close()
    print("5_pre_final is complete")
    
if __name__ == '__main__':
    main()
    