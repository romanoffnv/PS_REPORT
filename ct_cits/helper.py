L_crutches = [' в 8', ' в', ' эл', ' на ЮР 5']
L = ['567 в',  '7954УМ86 эл',  '097 в',  '7824 в 8',  '373 на ЮР 5',  '4786 УВ 86',  'у572еу116']

import re
# removing trash 1 iteration
# pattern = re.compile('(\s*\в\s*\d*)|(\s*\эл)|(\s*\на ЮР 5)')
L_crutches = ['\s*\в\s*\d*', '\s*\эл', '\s*\на ЮР 5']
for crutch in L_crutches:
    L_modified = [''.join(re.sub(crutch, '', x)).strip() for x in L]

print(L_modified)