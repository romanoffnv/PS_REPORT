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
import win32com
print(win32com.__gen_path__)


# Making connections to DBs
db = sqlite3.connect('drivers.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

# Pandas
pd.set_option('display.max_rows', None)
df = pd.read_excel('cunt.xlsx')


def main():
    # FUNCTIONS
    # 1. Data query
    def getter(df, fleet, crew, drvs):
        # Get data from xls
        L0 = df['Целевая техника'].tolist()
        L1 = df[fleet].tolist()
        L2 = df[drvs].tolist()
        
        # Do filtering and merging
        df = pd.DataFrame(zip(L0, L1, L2), columns=['Name', 'Units', 'Drivers'])
        df = df.dropna(how='any', subset=['Units'], thresh=1)
        df = df.loc[df['Units'].str.len() > 6]
        df["Units"] = df["Name"].str.cat(df["Units"], sep=". ")
        df = df.drop('Name', axis=1)

        # Listing filtered df cols
        L_units = df['Units'].tolist()
        L_drvs = df['Drivers'].tolist()

        # Cleaning units
        L_clean_words = ['\n', r'\[Н]']
        for i in L_clean_words:
            L_units = [re.sub(i, ' ', x) for x in L_units]

        # Fixing Nan in drivers
        L_drvs = [str(x) for x in L_drvs]
        L_drvs = [str(x).replace('nan', 'Нет закрепленных водителей') for x in L_drvs]

        # Attaching crew
        L_crew = []
        for i in L_units:
            L_crew.append(crew)
        df = pd.DataFrame(zip(L_crew, L_units, L_drvs), columns = ['Crew', 'Units', 'Drivers'])
        df = df.loc[df['Drivers'].str.len() != 4]
        return df

    # Fishing out plates by regex from long sentences
    def plate_ripper(L_units):
        def plate_fisher(regex, L_units):
            L_plates_temp = []
            for i in L_units:
                if 'ДЭС' in i:
                    L_plates_temp.append(i)
                else:
                    if re.findall(regex, str(i)):
                        L_plates_temp.append(''.join(re.findall(regex, str(i))))
                    else:
                        L_plates_temp.append(i)
                    # print(i)
        
            L_units = [str(x).strip() for x in L_plates_temp]
            L_plates_temp.clear() 
                
            return L_units

        L_regex = [
            '\s\D{2}\s*\d{2}\s*\d{2}\s*\d+', #ВВ  4553 86, # АН 78 96 82, ВВ  4553 86
            '\s\D\s*\d{3}\s*\D{2}\s*\d+', #Е 898 СВ 186, У 039 ВК186
            '\s\D\s\d{4}\s+\d+', #H 0762  07
            '\s\d{4}\s\D{2}\s+\d+', #7713 НХ 77
            '\s\D{2}\-\D+\-\d+', #CT-DV-141, CT-CTU-1000
            '\s\D{3}\-\d+', #HFU-2000
            '\№\s\d+', #№ 0079
            '\s\D\s*\d{3}\s*\D{2}\s*\d+', #runs again to choose bw paranthesis and outside par Е 898 СВ 186
        ]

        L_plates = plate_fisher(re.compile(L_regex[0]), L_units) 
        
        for regex in L_regex:
            L_plates = plate_fisher(re.compile(regex), L_plates)

        return L_plates

        # Turn plates into 123abc type
    def transform_plates(plates):
        plates = [re.sub('\s+', '', x) for x in plates]
        L_regions_long = [126, 156, 158, 174, 186, 188, 196, 797]
        L_regions_short = ['01', '02', '03', '04', '05', '06', '07', '09']
        for i in L_regions_long:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 9 else x for x in plates]
        for i in L_regions_short:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        for i in range(10, 100):
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates
    # ================================================================================================================================================
    # PROCEDURES
    # Getting and merging dfs by fleets
    df_1 = getter(df, 'Флот №1', 'ГРП 1', 'Водители_1')
    df_2 = getter(df, 'Флот №2', 'ГРП 2', 'Водители_2')
    df_m = pd.merge(df_1, df_2, how="outer")
    df_3 = getter(df, 'Флот №3', 'ГРП 3', 'Водители_3')
    df_m = pd.merge(df_m, df_3, how="outer")
    df_4 = getter(df, 'Флот №4', 'ГРП 4', 'Водители_4')
    df_m = pd.merge(df_m, df_4, how="outer")
    df_5 = getter(df, 'Флот №5', 'ГРП 5', 'Водители_5')
    df_m = pd.merge(df_m, df_5, how="outer")
    df_6 = getter(df, 'Флот №6', 'ГРП 6', 'Водители_6')
    df_m = pd.merge(df_m, df_6, how="outer")
    df_7 = getter(df, 'Флот №7', 'ГРП 7', 'Водители_7')
    df_m = pd.merge(df_m, df_7, how="outer")
    df_8 = getter(df, 'Флот №8', 'ГРП 8', 'Водители_8')
    df_m = pd.merge(df_m, df_8, how="outer")
    df_9 = getter(df, 'Флот №9', 'ГРП 9', 'Водители_9')
    df_m = pd.merge(df_m, df_9, how="outer")
    df_14 = getter(df, 'Флот №14', 'ГРП 11', 'Водители_14')
    df_m = pd.merge(df_m, df_14, how="outer")
    df_15 = getter(df, 'Флот №15', 'ГРП 15', 'Водители_15')
    df_m = pd.merge(df_m, df_15, how="outer")
    df_16 = getter(df, 'Флот №16', 'ГРП 16', 'Водители_16')
    df_m = pd.merge(df_m, df_16, how="outer")
    
    # Derivating plates from units
    L_units = df_m['Units'].tolist()
    L_plates = plate_ripper(L_units)
    
    # Conerting plates into 123abc format
    L_PI_cunt = transform_plates(L_plates) 
    
    
    df_final = pd.DataFrame(zip(L_plates, L_PI_cunt), columns= ['Plates', 'PI'])
    df_final = df_m.join(df_final, how = 'left')
    pprint(df_final)
    
    # Post df to DB
    cursor.execute("DROP TABLE IF EXISTS frac_drivers")
    df_final.to_sql(name='frac_drivers', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()

  
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))