import json
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

# Pandas
pd.set_option('display.max_rows', None)

# db connections
db = sqlite3.connect('cits.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()

def main():
   L_dept = cursor.execute("SELECT Dept FROM Trucks").fetchall()
   L_units = cursor.execute("SELECT Unit FROM Trucks").fetchall()
   

   # Removing rows with 'del' marker
   df = pd.DataFrame(zip(L_dept, L_units), columns=['Dept', 'Unit'])
   df = df[df["Unit"].str.contains("del") == False]
   # Crutch
   df = df[df["Unit"].str.contains('с ГНКТ') == False]
   # print(df)
   # print(df.describe())

   
   L_dept = df.loc[:, 'Dept'].tolist()
   L_units = df.loc[:, 'Unit'].tolist()
   
      
   # Splitting each plate into separate substrings
   L_plates = []
   for i in L_units:
      L_plates.append(i.split())

   for i in L_plates:
      for j in i:
         if j.isalpha() == True and len(j) > 2:
               i.remove(j)
               
   for i in L_plates:
      for j in i:
         if j.isalpha() == True and len(j) > 2:
               i.remove(j)

   
   L_plates_temp = []
   for i in L_plates:
      L_plates_temp.append(' '.join(i))
   L_plates = [x for x in L_plates_temp]
   L_plates_temp.clear()
   
   
   for i in L_plates:
      ind = re.search('\инв', i)
      if ind:
         ind = ind.start()
         L_plates_temp.append(i[ind:])
      else:
         L_plates_temp.append(i)

   L_plates = [x for x in L_plates_temp]
   L_plates_temp.clear()    
 
   
   # Turn plates into 123abc type
   def transform_plates(plates):
      L_regions = [186, 86, 797, 116, '02', '07', 82, 89, 78, 54, 77, 126, 188, 88, 56, 174, 74, 158, 196, 156, 76]
        
      for i in L_regions:
         plates = [x.removesuffix(str(i)).strip() if x != None and len(x) > 7 else x for x in plates]
      plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
      plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
      plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
      plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
      plates = ['' if len(x) > 6 else x for x in plates]
            
      return plates
    

   L_plates_ind = transform_plates(L_plates)

   # Crutch patching plates
   D_trucks = json.load(open('D_trucks.json'))
   for k, v in D_trucks.items():
      L_plates_ind = [re.sub(k, v, x) for x in L_plates_ind]
     
   # temp crutch for diesels
   for i in L_plates_ind:
      if i == '':
         ind = L_plates_ind.index(i)
         L_plates_ind[ind] = 111111
   pprint(L_plates_ind)

   df = pd.DataFrame(zip(L_dept, L_units, L_plates, L_plates_ind), columns=['Dept', 'Unit', 'Plate', 'Plate_index'])
   df = df.drop_duplicates(subset=['Plate'], keep='first')
   # print(df)
  
   # Post into db
   cursor.execute("DROP TABLE IF EXISTS Trucks")
   df.to_sql(name='Trucks', con=db, if_exists='replace', index=False)
   db.commit()
   db.close()

if __name__ == '__main__':
    main()