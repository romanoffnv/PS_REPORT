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
db = sqlite3.connect('data.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()


def main():
    online = json.load(open('JSON_om_online.json'))
    L_uuid_loc, L_status_loc = [], [] 
    L_total_groups = cursor.execute(f"SELECT Groups FROM om_parse_trucks").fetchall()
    L_total_units = cursor.execute(f"SELECT Units FROM om_parse_trucks").fetchall()
    L_total_uuid = cursor.execute(f"SELECT id FROM om_parse_trucks").fetchall()
    
    for i in online:
        if cursor.execute(f"SELECT id FROM om_parse_trucks WHERE id like '%{i['uuid']}%'").fetchall():
            L_status_loc.append(i['status'])
            L_uuid_loc.append(i['uuid'])
        else:
            L_status_loc.append('')
            L_uuid_loc.append('')
            
    
    D = dict(zip(L_uuid_loc, L_status_loc))
    L_stat_total = []
    for i in L_total_uuid:
            L_stat_total.append(D.get(i))
        
    L_stat_total = [str(x) for x in L_stat_total]
    L_stat_total = [x.replace('3', 'Не в сети более 36 часов') for x in L_stat_total]
    L_stat_total = [x.replace('2', 'Не в сети более 5 часов') for x in L_stat_total]
    L_stat_total = [x.replace('1', 'В сети') for x in L_stat_total]
    L_status = [x.replace('None', 'Нет данных') for x in L_stat_total]
    
    L_total_groups = [x.replace('\t', ' ') for x in L_total_groups]
    df = pd.DataFrame(zip(L_total_groups, L_total_units, L_total_uuid, L_status), columns=['Groups', 'Units', 'id', 'Status'])
    pprint(df)
    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS om_parse_status")
    df.to_sql(name='om_parse_status', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
if __name__== '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))