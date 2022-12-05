from dataclasses import dataclass
from sys import prefix
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


# Making connections to DBs
# connection to cits.db
db = sqlite3.connect('data.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()


def main():
    L_crews = cursor.execute("SELECT Crews FROM cits_parse").fetchall()
    L_units = cursor.execute("SELECT Units FROM cits_parse").fetchall()
    L_plates = cursor.execute("SELECT Plates FROM cits_parse").fetchall()
    L_locs = cursor.execute("SELECT Locations FROM cits_parse").fetchall()
    
    
    # Pre-cleaning 
    L_units = [re.sub('\s+', ' ', x) for x in L_units]
    L_locs = ['-' if v == 'None' else v for v in L_locs]

    # Slicing field name (Ю/Приобское м/р\nООО "ГАЗПРОМНЕФТЬ-ХАНТОС" - Ю/Приобское м/р)
    L_locs_temp = []
    for i in L_locs:
        if 'м/р' in i:
            ind = i.index('м/р')
            L_locs_temp.append(i[:ind + 3])
        else:
            L_locs_temp.append(i)
    
    L_locs = [x for x in L_locs_temp]
    L_locs_temp.clear()
    
    # Clean plates
    L_cleanit = ['\-', '/']
    for i in L_cleanit:
        L_plates = [re.sub(i, '', x) for x in L_plates]
   
    # Converting unconditioned plates into conditioned ones thru the manually supported dict
    D_ct_plates = json.load(open('D_ct_plates.json'))    
    for k, v in D_ct_plates.items():
        for j in L_plates:
            if k == j:
                ind = L_plates.index(j)
                L_plates[ind] = v

                
    # Turn plates into 123abc type
    def transform_plates(plates):
        L_regions = [186, 116, 86, 797, '02', '07', 82, 78, 54, 77, 126, 188, 89, 88, 174, 74, 158, 196, 156, 56, 76]
        
        for i in L_regions:
            plates = [x.removesuffix(str(i)).strip() if x != None and len(x) > 7 else x for x in plates]
        plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
        plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
        plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
        plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
        return plates


    L_plates_ind = transform_plates(L_plates)
    
    # Fixing diesels
    D_diesels = {
                
                '000004235': 'ДЭС CHH1961 (инв. №000004235)',
                '000002219': 'ДЭС АД-30С-Т400 727816 (инв. №000002219)',
                '400': 'ДЭС АД-30-Т/400 (инв. № Ю/НБ-0013)',
                
                # '445141206023000003798дэсcumminscd(s)№(инв.№)': 'ДЭС Cummins C44D5(S) №141206023 (инв. №000003798)',
                # '442440000004312дэсhgpcchk(инв.№)': 'ДЭС HG44PC CHK2440 (инв. №000004312)',
                # '39606007дэсинв.(f)': ' ДЭС инв.396 (F06007)',
                # '39906010дэсинв.(f)': 'ДЭС инв.399 (F06010)',
                # '441945000004326дэсhgpccnn(инв.№)': 'ДЭС HG44PC CNN1945 (инв. №000004326)',
                # '5064000004236дэсchv(инв.№)': 'ДЭС CHV5064 (инв. №000004236)',
                # '5000003346дэс№(инв.№)': 'ДЭС № 5 (инв. №000003346)',
                # '100230350001дэсwt(инв.№эл-)': 'ДЭС WT10023035 (инв. №ЭЛ-0001)',
                # '448805000005567дэсalfa(инв.№)': 'ДЭС ALFA 448805 (инв. №000005567)',
                # '696602000005568дэсalfa(инв.№)': ' ДЭС ALFA 696602 (инв. №000005568)',
                # '2439000004408дэсchк(инв.№)': 'ДЭС CHК2439 (инв. №000004408)',
                # '665дэсcumminscd': 'ДЭС Cummins C66D5',
                # '902788000003647дэсkubota(инв.№)': 'ДЭС Kubota 902788 (инв. №000003647)',
                # '5767130002дэс(инв.№эл-)': 'ДЭС 576713 (инв. №ЭЛ-0002)',
                # '441907010дэсhggl№': ' ДЭС HG44GL № 1907010',
                # '397дэсинв.': ' ДЭС инв.397',
                # '398дэсинв.': ' ДЭС инв.398',
                # '1дэс': ' дэс1',
            }
    
    for k, v in D_diesels.items():
        L_plates_ind = [''.join(x.replace(k, v)).strip() for x in L_plates_ind]
        
    for i in L_plates_ind:
        if len(i) != 6:
            print(i)

    df = pd.DataFrame(zip(L_crews, L_units, L_plates, L_plates_ind, L_locs), columns=['Crews', 'Units', 'Plates', 'Plate_index', 'Locations'])
    df = df.drop_duplicates(subset='Plate_index', keep="first")
    # print(df)
    # Posting df to DB
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS cits_final")
    df.to_sql(name='cits_final', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))