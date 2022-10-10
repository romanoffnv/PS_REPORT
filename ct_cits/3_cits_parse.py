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
   def list_generator(L):
      L_ct = []
        
      for i in L:
         if type(i) == str:
            L_ct.append(i.replace(';', ','))
        
      L_ct = [x.replace(':', ',') for x in L_ct]
        
        
      # Splits
        
      L_ct = [x.replace(' 86- ', ' 86 split ') for x in L_ct ]
      L_ct = [x.replace(',', ' split ') for x in L_ct]
      L_ct_compressed = [x.split('split') for x in L_ct]
      L_ct_compressed = list(itertools.chain.from_iterable(L_ct_compressed))
      L_ct_compressed = [x.split('+') for x in L_ct_compressed]
      L_ct_compressed = list(itertools.chain.from_iterable(L_ct_compressed))
      L_ct_compressed = [x.split('RUS') for x in L_ct_compressed]
      L_ct_compressed = list(itertools.chain.from_iterable(L_ct_compressed))
       
        
      return L_ct_compressed
    
   # Splitting data in strings into the lists(which get extracted) into the list by running list_generator func
   L_units = json.load(open('L_ct_units.json'))
   
   L_units_temp = []
   for i in L_units:
      L_units_temp.append((list_generator(i)))
   L_units = [x for x in L_units_temp]
   L_units_temp.clear()
    
   
   def list_cleaner(i):
      L_ct_clean = [x.strip() for x in i]
        
       
        
      # Replacing crappy unit names into omnicomm smth
      D_replacers = json.load(open('D_ct_replacers.json'))
      for k, v in D_replacers.items():
         L_ct_clean = [x.replace(k, v) for x in L_ct_clean]
        
        
      # removing trash 
      D_patterns = json.load(open('D_ct_patterns.json'))
      for k, v in D_patterns.items():
         L_ct_clean = [re.sub(k, v, x) for x in L_ct_clean]
        
      # extracting plates from brackets
      L_ct_clean = [x.split('(') for x in L_ct_clean]
      L_ct_clean = list(itertools.chain.from_iterable(L_ct_clean)) 
        
      # removing spaces
      L_ct_clean = [''.join(re.sub('\s+', ' ', x)).strip() for x in L_ct_clean]

      # Removing items that don't have numbers (i.e. plates)
      pattern_D = re.compile(r'\d')
      L_ct_clean = [x for x in L_ct_clean if re.findall(pattern_D, str(x))]
        

      # The ultimate list should contain items of 3 types: 100% - МЗКТ УУ 0775 86, 80% - Автокран 766, 50% - 232
      return L_ct_clean
    
   # Running list_cleaner func to clean up trash
   for i in L_units:
      L_units_temp.append((list_cleaner(i)))
   L_units = [x for x in L_units_temp]
   L_units_temp.clear()
    

   L_crews = json.load(open('L_ct_crews.json'))
   L_fields = json.load(open('L_ct_fields.json'))
   
   
   
   # Pre-cleaning fields
   L_fields = ['-' if v is None else v for v in L_fields]
   L_fields_temp = []
   
   # Slicing field name (Ю/Приобское м/р\nООО "ГАЗПРОМНЕФТЬ-ХАНТОС" - Ю/Приобское м/р)
   for i in L_fields:
      if 'м/р' in i:
         ind = i.index('м/р')
         L_fields_temp.append(i[:ind + 3])
      else:
         L_fields_temp.append(i)
   
   L_fields = [x for x in L_fields_temp]
   L_fields_temp.clear()
   
   # Mulitplying crews and fields by units
   def multiplier(x):
      L = [(k + '**').split('**') * len(v) for k, v in zip(x, L_units)]
      L = list(itertools.chain.from_iterable(L))
      L = list(filter(None, L))

      return L
   
   L_crews = multiplier(L_crews)
   L_fields = multiplier(L_fields)
   
   
   
   # Merging units lists into one list
   L_units = list(itertools.chain.from_iterable(L_units))
       
     
   # Removing rows with 'del' marker
   df = pd.DataFrame(zip(L_crews, L_units, L_fields), columns=['Crew', 'Unit', 'Loc'])
   df = df[df["Unit"].str.contains("del") == False]
   # Crutch
   df = df[df["Unit"].str.contains('с ГНКТ') == False]
   
   L_dept = df.loc[:, 'Crew'].tolist()
   L_units = df.loc[:, 'Unit'].tolist()
   L_fields = df.loc[:, 'Loc'].tolist()
   
   print(df)
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
   D_trucks = json.load(open('D_ct_trucks.json'))
   for k, v in D_trucks.items():
      L_plates_ind = [re.sub(k, v, x) for x in L_plates_ind]
     
 

   df = pd.DataFrame(zip(L_dept, L_units, L_plates, L_plates_ind, L_fields), columns=['Dept', 'Unit', 'Plate', 'Plate_index', 'huis'])
   df = df.drop_duplicates(subset=['Plate'], keep='first')
   
   

   # Post into db
   cursor.execute("DROP TABLE IF EXISTS Trucks")
   df.to_sql(name='Trucks', con=db, if_exists='replace', index=False)
   db.commit()
   db.close()

if __name__ == '__main__':
    main()