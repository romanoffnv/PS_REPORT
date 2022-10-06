import xlsxwriter
from multiprocessing.sharedctypes import Value
import pandas as pd
import os
import sqlite3
import re
from pprint import pprint
from win32com.client.gencache import EnsureDispatch
import win32com
print(win32com.__gen_path__)
# Get the Excel Application COM object
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\geozonesreport.xlsx")
ws = wb.Worksheets(1)

def main():
    # Getting lists from xls file

    L_vehicles, L_locations = [], []
    row = 10
    while ws.Cells(row, 1).Value != None:
        L_vehicles.append(ws.Cells(row, 1).Value)
        L_locations.append(ws.Cells(row, 2).Value)
        row += 1
    # removing duplicated spaces (xls file has some items with duplicated spaces)
    L_vehicles = [re.sub('\s+', ' ', x) for x in L_vehicles]

    # converting list into data frame
    data = pd.DataFrame(zip(L_vehicles, L_locations), columns = ['Locations', 'Vehicles'])
   
    # sorting dataframe by the list of locations
    data = data.loc[data.Locations.isin(['Барсуковское м/р', 'Бахиловское м/р', 'БПО Вынгапуровский', 'БПО г.Богучаны',
                                        'БПО г.Бузулук', 'БПО г.Губкинский', 'БПО г.Нижневартовск', 'БПО г.Новый Уренгой', 
                                        'БПО г.Ноябрьск', 'БПО г.Радужный', 'БПО г.Сургут', 'БПО г.Сургут (Инженерная, 20)',
                                        'В-Мессояхское м/р', 'Ван-Еганское м/р', 'Верхнепурпейское м/р', 'Восточно-Сургутское м/р',
                                        'Вынгапуровское м/р', 'Вынгаяхинское м/р', 'Губкинское м/р', 'Дулисьминское м/р',
                                        'З-Варьеганское м/р', 'З-Мессояхское', 'З/Иркинское м/р', 'Карамовское м/р', 'Колик-Еганское м/р',
                                        'Комсомольское м/р', 'Кондинское м/р.', 'Кошильское м/р', 'КПП Карамовского м- я', 'КПП Торкасинское',
                                        'Крайнее м/р', 'Кудринское м/р', 'Кузоваткинское м/р', 'Куюмбинское м/р', 'Луцеяхское м/р', 
                                        'Малобалыкское м/р', 'Мамонтовское м/р', 'Метельное м/р', 'Московцева м/р', 'Новопурпейское м/р',
                                        'Омбинское м/р', 'Песчаное м/р', 'Петеленское м/р', 'пост на зимник С- Самбурского', 
                                        'пост СБ на зимник С- Самбурского', 'пост СБ С- Самбурского', 'Приобское м/р', 'Присклоновое м/р',
                                        'РН-Пурнефтегаз', 'Родниковское м/р', 'Романовское м-е', 'Русское м/р', 'С- Комсомольское м/р', 
                                        'С-Талинское м/р', 'Салымское м/р', 'Самбургское м/р', 'Северо-Варьеганское м/р', 'Северо-Самбурское м/р',
                                        'Северо-Харампурское м/р', 'Северо-Хохряковское м/р', 'Северо-Южное м/р', 'Северо-Ютымское м/р', 
                                        'Соровское м/р', 'Спорышевское м/р', 'Среднебалыкское м/р', 'Тагринское м/р', 'Тазовское, куст 92 ПО',
                                        'Тарасовское м/р', 'Тортасинское м/р', 'Фестивальное м/р', 'ЦДНГ Тортасинское', 'Чапровское м/р', 'Эргинское м/р',
                                        'Ю-Мессояхское м/р', 'Ю-Талинское м/р', 'Южно-Приобское м/р', 'Южно-Приобское м/р (левый берег)', 
                                        'Южно-Сургутское м/р', 'Южно-Харампурское м/р', 'Южнобалыкское м/р', 'Юрубчено-Тохомское м/р'])]

   
    # deleting duplicated records
    data = data.drop_duplicates(subset = ["Vehicles"])
    
    # Extracting table columns as lists from final_DB
    db = sqlite3.connect('omnicomm.db')
    db.row_factory = lambda cursor, row: row[0]
    cursor = db.cursor()

    
    L_department = cursor.execute("SELECT Department FROM final_DB").fetchall()
    L_vehicle = cursor.execute("SELECT Vehicle FROM final_DB").fetchall()
    L_plate = cursor.execute("SELECT Plate FROM final_DB").fetchall()
    L_veh_name = cursor.execute("SELECT Vehicle_Name FROM final_DB").fetchall()
    L_plate_index = cursor.execute("SELECT Plate_index FROM final_DB").fetchall()
    L_nodata = cursor.execute("SELECT No_data FROM final_DB").fetchall()

    
    # Converting dataframes for vehicles and locations into into lists
    loc_veh = [x for x in data.loc[:, 'Vehicles']]
    loc_loc = [x for x in data.loc[:, 'Locations']]

    # Forming temp list
    L_locco = []

    # Checking if items in total vehilce list are present in the list of vehicles obtained from location report
    # Getting and item's location by index into the temp list
    for i in L_vehicle:
        if i in loc_veh:
            ind = loc_veh.index(i)
            L_locco.append(loc_loc[ind])
        else:
            L_locco.append('Вне геозоны')

    data = pd.DataFrame(zip(L_vehicle, L_locco))
    

    # Updating L_locations from temp L_locco list
    L_locations = [x for x in L_locco]

    # Building dataframe
    data = pd.DataFrame(zip(L_department, L_vehicle, L_plate, L_veh_name, L_plate_index, L_locations, L_nodata), 
                columns = ['Department', 'Vehicle', 'Plate', 'Vehicle_name', 'Plate_index', 'Location_Omnicomm', 'No_data'])
    
    # Posting dataframe back into the sql database
    cursor.execute("DROP TABLE IF EXISTS final_DB")
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS final_DB(
        Department text,
        Vehicle text,
        Plate text,
        Vehicle_name text,
        Plate_index text,
        Location_Omnicomm text,
        No_data text
        )
        """)

    data.to_sql('final_DB', db, if_exists='replace', index = False)


    db.commit()
    db.close()

     # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('DB.xlsx', engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    data.index += 1
    data.to_excel(writer, index = True, header=True)
    writer.save()

    wb.Close(True)
    
    print("6_omn_match_locations_final.py is complete")
    
if __name__ == '__main__':
    main()