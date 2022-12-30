import time
import xlsxwriter
from win32com.client.gencache import EnsureDispatch
import os
import re
from pprint import pprint
import pandas as pd
from functools import reduce
import itertools
import sqlite3
import win32com
print(win32com.__gen_path__)

# Excel connection  
# xl = EnsureDispatch('Excel.Application')
# wb = xl.Workbooks.Open(f"{os.getcwd()}\\Книга1.xlsx")
# ws1 = wb.Worksheets(1)


def main():
    df = pd.read_excel('Книга1.xlsx')
    df = df.dropna(how='any', subset=['Флот', 'СПТ'], thresh=1)
    df = df.drop_duplicates(subset='СПТ', keep="last")


    pprint(df)

    import xlsxwriter
    writer = pd.ExcelWriter('Книга2.xlsx', engine='xlsxwriter')
    # Write each dataframe to a different worksheet.
    df.index += 1
    df.to_excel(writer, index = True, header=True)
    writer.save()

    # string = ws1.Cells(1, 1).Value
    # string = string.split(', ')
    # for i in string:
    #     print(re.sub('\(.*\)', '', i))

    # row = 1
    # L = []
    # while True:
    #     if ws1.Cells(row, 1).Value != None:
    #         L.append(ws1.Cells(row, 1).Value)
    #     row += 1
    #     if row == 82:
    #         break
    # pprint(len(L))

    # L = set(L)
    # pprint(len(L))
    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))