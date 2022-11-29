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

# connection to match.db
db_match = sqlite3.connect('match.db')
db_match.row_factory = lambda cursor, row: row[0]
cursor = db_match.cursor()
cnx_match = sqlite3.connect('match.db')

def main():
    df = pd.read_sql_query("SELECT * FROM cunt_to_match", cnx_match)
    writer = pd.ExcelWriter('DB.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index = True, header=True)
    writer.save()

if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))