import sqlite3
from pprint import pprint

db = sqlite3.connect('omnicomm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

def main():
    print(len(cursor.execute("SELECT * FROM total_vehicles_nobubs").fetchall()))
    # Extracting table columns from total_vehicles_nobubs
    L_total_dep = cursor.execute("SELECT Department FROM total_vehicles_nobubs").fetchall()
    L_total_veh = cursor.execute("SELECT Vehicle FROM total_vehicles_nobubs").fetchall()
    # L_total_veh = [x.strip() for x in L_total_veh]
    print(len(L_total_dep))
    
    # Excracting red column from red_vehicles to match with total
    L_red_veh = cursor.execute("SELECT Red_Vehicle FROM red_vehicles").fetchall()
    L_grey_veh = cursor.execute("SELECT Grey_Vehicle FROM grey_vehicles").fetchall()
    L_orange_veh = cursor.execute("SELECT Orange_Vehicle FROM orange_vehicles").fetchall()
    L_red_veh = [x.strip() for x in L_red_veh]
    L_grey_veh = [x.strip() for x in L_grey_veh]
    L_orange_veh = [x.strip() for x in L_orange_veh]
    L_nodata = []
    for i in L_total_veh:
        if i in L_red_veh:
            L_nodata.append('Нет данных более 36ч')
        elif i in L_orange_veh:
            L_nodata.append('Нет данных более 5ч')
        elif i in L_grey_veh:
            L_nodata.append('Нет данных в Омникомм')
        else:
            L_nodata.append('')
    
    # # Updating table of total vehicles with red column
    
    cursor.execute("DROP TABLE IF EXISTS total_vehicles_nobubs_nodata")
    cursor.execute("""
                    CREATE TABLE IF NOT EXISTS total_vehicles_nobubs_nodata(
                    Department text,
                    Vehicle text,
                    Nodata text
                                )
                    """)
    cursor.executemany("INSERT INTO total_vehicles_nobubs_nodata VALUES (?, ?, ?)", zip(L_total_dep, L_total_veh, L_nodata))
    pprint(cursor.execute("SELECT * FROM total_vehicles_nobubs_nodata").fetchall())
    pprint(len(cursor.execute("SELECT * FROM total_vehicles_nobubs_nodata").fetchall()))
    
    # refreshing database
    db.commit()
    # closing database
    db.close()
    print("4_omn_match_nodata is complete")
if __name__ == '__main__':
    main()