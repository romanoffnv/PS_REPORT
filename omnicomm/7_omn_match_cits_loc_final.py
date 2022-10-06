from datetime import datetime
import xlsxwriter
import pandas as pd
import os
import sqlite3
import re
from pprint import pprint
from win32com.client.gencache import EnsureDispatch
import win32com
print(win32com.__gen_path__)
import itertools

# Global excel connections
xl = EnsureDispatch('Excel.Application')
# wb = xl.Workbooks.Open(f"{os.getcwd()}\\Sverka.xlsx")
wb2 = xl.Workbooks.Open(f"{os.getcwd()}\\1.xlsx")
wb3 = xl.Workbooks.Open(f"{os.getcwd()}\\2.xlsx")

# Global variables
prevMonth = datetime.now().month - 1
currMonth = datetime.now().month

def main():
    def get_ct():
        pass
    # data_ct = sverka_ct()
    # data_fr = sverka_fr()
    # data_ct = ct(data_ct)
    # data_fr = fr(data_fr)
    # merge_ctfr = mergeDF(data_ct, data_fr)
    # match(merge_ctfr)
    

# def sverka_ct():
#     # Get CT sheet
#     ws1 = wb.Worksheets(1)
    
#     # Get ct crew names col(B) and plate col(D) from excel file Sheet(1) of Sverka.xlsx into respective lists
#     L_crews_ct_cits, L_plates_ct_cits = [],[]
#     row = 4
#     while True:
#         L_crews_ct_cits.append(ws1.Cells(row, 2).Value)
#         L_plates_ct_cits.append(ws1.Cells(row, 4).Value)
#         row += 1
#         if ws1.Cells(row, 2).Value == None:
#             break
#     # print(len(L_crews_ct_cits))
#     # print(len(L_plates_ct_cits))

#     # Update plates up to the standard (index, literal, stip region)
#     # Derivate indeces from plates
#     L_index_ct_cits = [x.removesuffix('186').strip() if x != None else x for x in L_plates_ct_cits]
#     L_index_ct_cits = [x.removesuffix('86').strip() if x != None else x for x in L_index_ct_cits]
#     L_index_ct_cits = [x.removesuffix('116').strip() if x != None else x for x in L_index_ct_cits]
#     L_index_ct_cits = [x.removesuffix('82').strip() if x != None else x for x in L_index_ct_cits]
#     L_index_ct_cits = [x.removesuffix('89').strip() if x != None else x for x in L_index_ct_cits]
#     L_index_ct_cits = [x.removesuffix('56').strip() if x != None else x for x in L_index_ct_cits]
#     L_index_ct_cits = [x.removesuffix('797').strip() if x != None else x for x in L_index_ct_cits]
#     L_index_ct_cits = [x.removesuffix('54').strip() if x != None else x for x in L_index_ct_cits]
#     L_index_ct_cits = [x.removesuffix('77').strip() if x != None else x for x in L_index_ct_cits]
#     L_index_ct_cits = [re.sub('\D', '', str(x)) for x in L_index_ct_cits]
#     # pprint(L_index_ct_cits)
#     # pprint(len(L_index_ct_cits))

#     # Derivate literals from plates
#     L_literal_ct_cits = [''.join(re.findall('\D', str(x))).lower().strip() for x in L_plates_ct_cits]
#     # pprint(L_literal_ct_cits)
#     # pprint(len(L_literal_ct_cits))

#     # Pack data frame in the shape of
#     data_ct = pd.DataFrame(zip(L_crews_ct_cits, L_index_ct_cits, L_literal_ct_cits))
#     # print(data_ct)
#         # 0     ГНКТ 1  0775    уу
#         # 1     ГНКТ 1   315   вам
#     return data_ct

# def sverka_fr():
#     # Get Frac sheet
#     ws2 = wb.Worksheets(2)
#     # Get frac crew names col(B) and plate col(D) from excel file of Sheet(2) Сверка по персоналу и технике 30.07.2022 v2.xlsx into respective lists
#     L_crews_fr_cits, L_plates_fr_cits = [],[]
#     row = 4
#     while True:
#         L_crews_fr_cits.append(ws2.Cells(row, 2).Value)
#         L_plates_fr_cits.append(ws2.Cells(row, 4).Value)
#         row += 1
#         if ws2.Cells(row, 2).Value == None:
#             break
#     # print(len(L_crews_fr_cits))
#     # print(len(L_plates_fr_cits))
#     wb.Close(True)
#     # Update plates up to the standard (index, literal, stip region)
#     # Derivating indeces from plates. Stripping regions
#     L_index_fr_cits = [x.removesuffix('186').strip() if x != None else x for x in L_plates_fr_cits]
#     L_index_fr_cits = [x.removesuffix('86').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('116').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('82').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('89').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('156').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('56').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('797').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('54').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('77').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('07').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('126').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('78').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [x.removesuffix('23').strip() if x != None else x for x in L_index_fr_cits]
#     L_index_fr_cits = [re.sub('\D', '', str(x)) for x in L_index_fr_cits]
#     # pprint(L_index_fr_cits)
#     # pprint(len(L_index_fr_cits))

#     # Derivate literals from plates
#     L_literal_fr_cits = [''.join(re.findall('\D', str(x))).lower().strip() for x in L_plates_fr_cits]
#     L_literal_fr_cits = [''.join(re.sub('\s+', '', str(x))).lower().strip() for x in L_literal_fr_cits]
#     # pprint(L_literal_fr_cits)
#     # pprint(len(L_literal_fr_cits))
#         # ГРП 1 686 вмт

#     # Pack data frame in the shape of
#     data_fr = pd.DataFrame(zip(L_crews_fr_cits, L_index_fr_cits, L_literal_fr_cits))
#     # print(data_fr)
#     return data_fr

# def ct(data_ct):
#     #Populate list of fields by Ct crew, taken from cits for the 30th date i.e. 
#     # Open up cits ct report to th 30th date
    
#     for i in wb2.Sheets:
#         if f"30.0{prevMonth}" in str(i.Name):
#             ws30_ct = wb2.Worksheets(i.Name)
#         elif f"30.0{currMonth}" in str(i.Name):
#             ws30_ct = wb2.Worksheets(i.Name)
#     # print(ws30_ct.Name)
#     # pprint(data_ct)

#     # Populating and cleaning list of fields provided by cits for ct
#     L_ct_fields = []
#     row = 4
#     while True:
#         L_ct_fields.append(ws30_ct.Cells(row, 5).Value)
#         row += 1
#         if ws30_ct.Cells(row, 1).Value == 'ГНКТ №32':
#             break
#     L_ct_fields = [x for x in L_ct_fields if x != None]
#     L_ct_fields = [x.split('\n') for x in L_ct_fields if x != None]
#     L_ct_fields = list(itertools.chain.from_iterable(L_ct_fields))
#     L_ct_fields = [re.sub(r'\мест\D+\s*', 'м/р', x) for x in L_ct_fields]
#     L_ct_fields = [x.strip() for x in L_ct_fields if 'м/р' in x or 'БПО' in x]
#     # pprint(L_ct_fields)
#     # pprint(len(L_ct_fields))

#     # Extracting list of crews from data_ct df
#     L_ct_crews = data_ct.iloc[:, 0].tolist()
#     L_ct_indeces = data_ct.iloc[:, 1].tolist()
#     L_ct_literals = data_ct.iloc[:, 2].tolist()
    
#     # Populating list of factors, cleaning off blanks
#     L_field_factors = []
#     for i in range(1, 32):
#         L_field_factors.append(L_ct_crews.count(f"ГНКТ {i}"))
#     L_field_factors = [x for x in L_field_factors if x != 0]
#     # print(L_field_factors)
#     # print(len(L_field_factors))

#     # Leveling fields list (no element for 31 crew) with field factors list (+ factor 31 crew)
#     if len(L_ct_fields) != len(L_field_factors):
#         L_ct_fields.append('')
#         # print(len(L_ct_fields) == len(L_field_factors))
    
#     # Fixing space problem with м/р
#     # pprint(L_ct_fields)
#     L_ct_fields2 = []
#     for i in L_ct_fields:
#         if 'м/р' in i:
#             ind = i.index('м/р')
#             if i[ind - 1] != ' ':
#                 L_ct_fields2.append(i.replace('м/р', ' м/р'))
#             else:
#                 L_ct_fields2.append(i)
#         else:
#             L_ct_fields2.append(i)
#     L_ct_fields = [x for x in L_ct_fields2]
#     # pprint(L_ct_fields2)
     
#     # Multiply number of crew items by respective field
#     L_ct_fields = [i.split(',') * j for i, j in (zip(L_ct_fields, L_field_factors))]
#     L_ct_fields = list(itertools.chain.from_iterable(L_ct_fields))
#     data_ct = pd.DataFrame(zip(L_ct_crews, L_ct_indeces, L_ct_literals, L_ct_fields))
#     # pprint(data_ct)
#     # Get the table of the following type:
#         # L1      L2   L3  L4  
#         # ГНКТ 1  0775 уу Барсуковское
#         # ГНКТ 1  315 вам Барсуковское
#         # ГНКТ 2  5848 уа Ю-Приобское
#         # ГНКТ 2  397 нкс Ю-Приобское
#     wb2.Close(True)
#     return data_ct

# def fr(data_fr):
#     # Extracting list of crews from data_ct df
#     L_fr_crews = data_fr.iloc[:, 0].tolist()
#     L_fr_indeces = data_fr.iloc[:, 1].tolist()
#     L_fr_literals = data_fr.iloc[:, 2].tolist()
#     # print(len(L_fr_crews))
#     # print(len(L_fr_indeces))
#     # print(len(L_fr_literals))
    
#     #Populate list of fields by Frac crew, taken from cits for the 30th date i.e. 
#     # Open up cits ct report to th 30th date
#     for i in wb3.Sheets:
#         if f"30.0{prevMonth}" in str(i.Name):
#             ws30_fr = wb3.Worksheets(i.Name)
#         elif f"30.0{currMonth}" in str(i.Name):
#             ws30_fr = wb3.Worksheets(i.Name)
#     print(ws30_fr.Name)
#     # pprint(data_ct)
    
#    # Populating and cleaning list of fields provided by cits for frac
#     L_fr_fields = []
#     row = 2
#     while True:
#         L_fr_fields.append(ws30_fr.Cells(row, 4).Value)
#         row += 1
#         if 'Начальник смены' in str(ws30_fr.Cells(row, 1).Value):
#             break
#     L_fr_fields = [x for x in L_fr_fields if x != None and len(x) > 3]
#     # pprint(L_fr_fields)
#     # pprint(len(L_fr_fields))

#     # Populating list of factors, cleaning off blanks
#     L_field_factors = []
#     for i in range(1, 16):
#         L_field_factors.append(L_fr_crews.count(f"ГРП {i}"))
#     L_field_factors = [x for x in L_field_factors if x != 0]
#     print(L_field_factors)
#     print(len(L_field_factors))

#     # Multiply number of crew items by respective field
#     L_fr_fields = [i.split(',') * j for i, j in (zip(L_fr_fields, L_field_factors))]
#     L_fr_fields = list(itertools.chain.from_iterable(L_fr_fields))
#     data_fr = pd.DataFrame(zip(L_fr_crews, L_fr_indeces, L_fr_literals, L_fr_fields))
#     # pprint(data_fr)
#     # Get the table of the following type:
#         # L1     L2  L3  L4  
#         # ГРП 1  686 вмт Тарасовское
#         # ГРП 1  889 вмт Тарасовское
#         # ГРП 2  788 вар С-Талинское
#         # ГРП 2  077 ввм С-Талинское
#     wb3.Close(True)
#     return data_fr

# def mergeDF(data_ct, data_fr):
#     data_mrg = pd.merge(data_ct, data_fr, how="outer")
#     return data_mrg
#     # print(data_mrg)
#     # Merge CT and Frac dataframes
#         # L1      L2   L3  L4  
#         # ГНКТ 1  0775 уу Барсуковское
#         # ГНКТ 1  315 вам Барсуковское
#         # ГНКТ 2  5848 уа Ю-Приобское
#         # ГНКТ 2  397 нкс Ю-Приобское
#         # ГРП 1  686 вмт Тарасовское
#         # ГРП 1  889 вмт Тарасовское
#         # ГРП 2  788 вар С-Талинское
#         # ГРП 2  077 ввм С-Талинское

# def match(merge_ctfr):
#     # Connect to the database
#     db = sqlite3.connect('omnicomm.db')
#     db.row_factory = lambda cursor, row: row[0]
#     cursor = db.cursor()

#     # final_DB destructuring
#     L_om_dep = cursor.execute("SELECT Department FROM final_DB").fetchall()
#     L_om_vehicle = cursor.execute("SELECT Vehicle FROM final_DB").fetchall()
#     L_om_plate = cursor.execute("SELECT Plate FROM final_DB").fetchall()
#     L_om_vehilce_name = cursor.execute("SELECT Vehicle_name FROM final_DB").fetchall()
#     L_om_plate_index = cursor.execute("SELECT Plate_index FROM final_DB").fetchall()
#     L_om_location = cursor.execute("SELECT Location_Omnicomm FROM final_DB").fetchall()
#     L_om_nodata = cursor.execute("SELECT No_data FROM final_DB").fetchall()
 
#     L_cits_index = merge_ctfr.iloc[:, 1]
#     L_cits_literal = merge_ctfr.iloc[:, 2]
#     L_cits_indlit = [x + y for x, y in zip(L_cits_index, L_cits_literal)]
#     L_cits_field = merge_ctfr.iloc[:, 3]

#     D_cits = dict(zip(L_cits_indlit, L_cits_field))
    
#     L_cits_fields = []
#     for i in L_om_plate_index:
#         L_cits_fields.append(D_cits.get(i))
       

#     # Building dataframe
#     data = pd.DataFrame(zip(L_om_dep, L_om_vehicle, L_om_plate, L_om_vehilce_name, L_om_plate_index, L_om_location, L_cits_fields, L_om_nodata), 
#                 columns = ['Department', 'Vehicle', 'Plate', 'Vehicle_name', 'Plate_index', 'Location_omnicomm', 'Location_cits', 'No_data'])
    
#     print(data)
#     # Posting dataframe back into the sql database
#     cursor.execute("DROP TABLE IF EXISTS final_DB")
#     cursor.execute("""
#         CREATE TABLE IF NOT EXISTS final_DB(
#         Department text,
#         Vehicle text,
#         Plate text,
#         Vehicle_name text,
#         Plate_index text,
#         Location_omnicomm text,
#         Location_cits text,
#         No_data text
#         )
#         """)

#     data.to_sql('final_DB', db, if_exists='replace', index = False)


#     db.commit()
#     db.close()

#     # Create a Pandas Excel writer using XlsxWriter as the engine.
#     writer = pd.ExcelWriter('DB.xlsx', engine='xlsxwriter')

#     # Write each dataframe to a different worksheet.
#     data.index += 1
#     data.to_excel(writer, index = True, header=True)
#     writer.save()
    print("7_omn_match_cits_loc.py is complete")
    
if __name__ == '__main__':
    main()
