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
db = sqlite3.connect('match.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

db_acc = sqlite3.connect('accountance.db')
db_acc.row_factory = lambda cursor, row: row[0]
cursor_acc = db_acc.cursor()

cnx_match = sqlite3.connect('match.db')
df_match = pd.read_sql_query("SELECT * FROM gen_PI", cnx_match)



# Pandas
pd.set_option('display.max_rows', None)
pd.options.display.width = 1200
pd.options.display.max_colwidth = 30
pd.options.display.max_columns = 30


def main():
    xl = pd.ExcelFile('gen_report.xls')
    # pprint(xl.sheet_names)  # see all sheet names
    
    def sheets_parser(service):
        df = xl.parse(service)
        df = df.drop(range(0, 11))
        L_units = df['Unnamed: 3']
        L_plates = df['Unnamed: 9']
        L_comments = df['Unnamed: 15']
        df = pd.DataFrame(zip(L_units, L_plates, L_comments), columns=['Units', 'Plates', 'Comments'])
        return df
    
    df_ct = sheets_parser('ГНКТ')
    df_fr = sheets_parser('ГРП')
    df = pd.merge(df_ct, df_fr, how="outer")

    df_trans = sheets_parser('ТР.Служба')
    df = pd.merge(df, df_trans, how="outer")
    df = df.dropna(how='any', subset=['Comments'], thresh=1)

    
    
    # List 'Units', 'Plates', 'Comments'
    L_units = df['Units'].tolist()
    L_plates = df['Plates'].tolist()
    L_comments = df['Comments'].tolist()
    L_plates = [str(x).strip() for x in L_plates]
    # L_plates = [re.sub('\s+', '', x) for x in L_plates]
    L_comments = [str(x).strip() for x in L_comments]
    
    # Pull items from units if 'Нет данных'
    L_plates_temp = []
    for i, j in zip(L_units, L_plates):
        if j == 'Нет данных' or j == None:
            L_plates_temp.append(i)
        else:
            L_plates_temp.append(j)
   
    L_units = [x for x in L_plates_temp]
    L_plates_temp.clear()

    
   # Fishing out plates by regex from long sentences
    
    def plate_fisher(regex, L_units):
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

    
    L_plates = plate_fisher(re.compile('\s\D\d+\D{2}\s\d+'), L_units)
    L_plates = plate_fisher(re.compile('\(\d{4}\D{2}\d{2}\)'), L_plates) #(7250ах86)
    L_plates = plate_fisher(re.compile('\D{2}\d{4}\s\d{2}'), L_plates) #s/n № 1000004402
    L_plates = plate_fisher(re.compile('s/n\s\№\s\d+'), L_plates) #НВД №1 ВВ8684 86
    L_plates = plate_fisher(re.compile('\s\D{1}\s*\d+\s*\D{2}\s*\d+'), L_plates)
    L_plates = plate_fisher(re.compile('\s\d+\s*\D{2}\s*\d+'), L_plates)
    L_plates = plate_fisher(re.compile('\s\инв.№\s*\d+'), L_plates)
    L_plates = plate_fisher(re.compile('\skz\s\D\d+\s\d+'), L_plates)
    L_plates = plate_fisher(re.compile('\s\D{2}\s*\d{4}\s*\d+'), L_plates)
    L_plates = plate_fisher(re.compile('\d{4}\D{2}\s\d+'), L_plates) #(8804ах 86)
    
    L_cleaners = ['№', '(', ')', '.']
    for i in L_cleaners:
        L_plates = [x.replace(i, '') for x in L_plates]

    

    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions_long = [126, 156, 158, 174, 186, 188, 196, 797]
        L_regions_short = ['01', '02', '03', '04', '05', '06', '07', '09']
        for i in L_regions_long:
            plates = [x.removesuffix(str(i)).strip() if x != None and (len(x) == 9 or len(x) == 8) else x for x in plates]
        for i in L_regions_short:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        for i in range(10, 100):
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates
    
   
    L_plates = [re.sub('\s+', '', x) for x in L_plates]
    L_PI = transform_plates(L_plates) 

    
    
    # Cruch fixing the items that don't match with gen_PI
    D_replacers = {
        '0013': '304000013дэсад--т/инвю/нб-',
        '113': '113пгу-озрд',
        '141206023': '445141206023000003798дэсcumminscdsинв',
        '2219': '000002219',
        '3346': '5000003346дэсинв',
        '4236': '5064000004236дэсchvинв',
        '4312': '442440000004312дэсhgpcchkинв',
        '4326': '441945000004326дэсhgpccnnинв',
        '7061': '7061умншситерра',
        '7062': '7062умншситерра',
        '7640': '7640ан',
        'F06007': '39606007дэсинвf',
        'F06010': '39906010дэсинвf',
        'H 0762  07': '076207н',
        'H0783  07': '078307н',
        'WT10023035': '10023035000 дэс1wtинвэл-',
        '10023035wt': '10023035000 дэс1wtинвэл-',
        '39786кс': '397нкс',
        '06007f': '39606007дэсинвf',
        '06010f': '39906010дэсинвf',
        '155954т': '1559тт',
        '672986а': '6729ва',
        '182386н': '1823ан',
        '181986в': '1819ва',
        '525786в': '5257вв',
        '078307h': '078307н',
        '665дэсcumminscd': 'ДЭС Cummins C66D5',
        '660386а': '6603ва',
        '286786а': '2867ва',
        '81386ос': '813рос',
        '197486в': '1974вв',
        '195686в': '1956вв',
        '240186а': '2401ва',
        '1дэс': 'дэс1',
        
    }
    
    for k, v in D_replacers.items():
        L_PI = [x.replace(k, v) for x in L_PI]
    
    df = pd.DataFrame(zip(L_units, L_plates, L_PI, L_comments))
    # pprint(df)

    # 1. See whatever of L_PI doesn't match to the match.db PI_gen
    L_matched, L_unmatched = [], []
    L_matched_comm = []
    for i, j in zip(L_PI, L_comments):
        if cursor.execute(f"SELECT PI_gen FROM gen_PI WHERE PI_gen like '%{i}%'").fetchall():
            L_matched.append(i)
            L_matched_comm.append(j)
        else:
            L_unmatched.append(i)


    df = pd.DataFrame(zip(L_matched, L_matched_comm), columns=['PI', 'Comments'])
    df = df.drop_duplicates(subset='PI', keep="last")

    # Posting df to DB
    print('Posting df to DB')
    cursor_acc.execute("DROP TABLE IF EXISTS accountance_3")
    df.to_sql(name='accountance_3', con=db_acc, if_exists='replace', index=False)
    db_acc.commit()
    db_acc.close()

if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
