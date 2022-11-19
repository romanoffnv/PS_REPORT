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

# Get the Excel Application COM object
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\cunt.xlsx")
Sheets = wb.Sheets.Count
ws = wb.Worksheets(Sheets)

# Making connections to DBs
db_match = sqlite3.connect('match.db')
db_match.row_factory = lambda cursor, row: row[0]
cursor_match = db_match.cursor()

db_acc = sqlite3.connect('accountance.db')
db_acc.row_factory = lambda cursor, row: row[0]
cursor_acc = db_acc.cursor()

cnx_match = sqlite3.connect('match.db')
cnx_acc = sqlite3.connect('accountance.db')
df_match = pd.read_sql_query("SELECT * FROM acc_to_match", cnx_match)
df_acc = pd.read_sql_query("SELECT * FROM accountance_3", cnx_acc)



# Pandas
pd.set_option('display.max_rows', None)
pd.options.display.width = 1200
pd.options.display.max_colwidth = 30
pd.options.display.max_columns = 30


def main():
    df = pd.read_excel('cunt.xlsx')
    
    wb.Close(True)
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
    print(df_m)
    print(df_m.describe())
    
    # Listing crew names and row numbers of the crew blocks
    # def getter(col):
    #     row = 2
    #     L_units = []
    #     while True:
    #         L_units.append(ws.Cells(row, col).Value)
    #         row += 1
    #         if row == 53:
    #             break
    #     return L_units
    
    # L_units = getter(2)

    # def cleaner(L):
    #     L = [str(x).split('___________________') for x in L]
    #     L = [', '.join(map(str, x)) for x in L]
    #     L = [re.sub('\n', ' ', x) for x in L]
    #     L = [str(x).split(',') for x in L]
            
    #     L = list(itertools.chain.from_iterable(L))
    #     L = [str(x).split(')') for x in L]
    #     L = list(itertools.chain.from_iterable(L))
    #     L = [str(x).strip() for x in L if x != 'None']
    #     L = [str(x).split('. ') for x in L]
    #     L = list(itertools.chain.from_iterable(L))

    #     L = [re.sub('\s+', ' ', x) for x in L]
    #     L = [''.join(re.sub(r'\[Н]', '', x)).strip() for x in L]
    #     L = [''.join(re.sub('\(', '', x)).strip() for x in L]
    #     return L
    
    # L_units = cleaner(L_units)

    #  # Fishing out plates by regex from long sentences
    # L_plates_temp = []
    # def plate_fisher(regex, L_units):
    #     for i in L_units:
    #         if 'ДЭС' in i:
    #             L_plates_temp.append(i)
    #         else:
    #             if re.findall(regex, str(i)):
    #                 L_plates_temp.append(''.join(re.findall(regex, str(i))))
    #             else:
    #                 L_plates_temp.append(i)
    #             # print(i)
    
    #     L_units = [str(x).strip() for x in L_plates_temp]
    #     L_plates_temp.clear() 
            
    #     return L_units

    
    # L_plates = plate_fisher(re.compile('\s\D{2}\s\d{4}\s\d+'), L_units)
    # L_plates = plate_fisher(re.compile('\s\D\s\d+\s\D{2}\s\d+'), L_plates)
    
    
    # D = dict(zip(L_units, L_plates))
    # # pprint(L_plates)
    # L_drivers = getter(3)
    # L_drivers = cleaner(L_drivers)
    # pprint(L_drivers)
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))