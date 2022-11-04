import time
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

# Making connections to db
db = sqlite3.connect('omnicomm.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
cnx = sqlite3.connect('omnicomm.db')

def main():
    # Getting lists from xls file

    L_units, L_locs = [], []
    row = 10
    while ws.Cells(row, 1).Value != None:
        L_locs.append(ws.Cells(row, 1).Value)
        L_units.append(ws.Cells(row, 2).Value)
        row += 1
    
    wb.Close(True)
    xl.Quit()

    # removing duplicated spaces (xls file has some items with duplicated spaces)
    L_units = [re.sub('\s+', ' ', x) for x in L_units]
    L_locs = [x if 'Итого' not in x else None for x in L_locs]
    
    # converting list into df frame
    df = pd.DataFrame(zip(L_units, L_locs), columns = ['Units', 'Locations'])
    df = df.dropna(how='any', subset=['Locations'], thresh=1)
    df = df.drop_duplicates(subset='Units', keep="last")

    pprint(df)

    print('Posting df to DB')
    cursor.execute("DROP TABLE IF EXISTS parse_locs")
    df.to_sql(name='parse_locs', con=db, if_exists='replace', index=False)
    db.commit()
    db.close()
    
    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))