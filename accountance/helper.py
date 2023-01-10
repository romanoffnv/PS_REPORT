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
    fullstring = "StackAbuse"
    substring = "tack"

    if fullstring.find(substring) != -1:
        print("Found!")
    else:
        print("Not found!")
    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))