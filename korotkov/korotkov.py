from functools import reduce
from pprint import pprint
from win32com.client.gencache import EnsureDispatch
import os
import win32com
print(win32com.__gen_path__)
# Get the Excel Application COM object
xl = EnsureDispatch('Excel.Application')
wb = xl.Workbooks.Open(f"{os.getcwd()}\\Сверка по персоналу и технике 30.07.2022 v2.xlsx")
wb2 = xl.Workbooks.Open(f"{os.getcwd()}\\Коротков.xlsx")
ws2 = wb.Worksheets(2)
ws13 = wb2.Worksheets(13)

def main():
    # Collecting plates from sverka frac1
    start = 5
    end = 36
    
    row = start
    L_plates = []
    while True:
        L_plates.append(ws2.Cells(row, 4).Value)
        row += 1
        if row == end:
            break

   
    # Adding spaces to match with Korotkov
    L_plates2 = []
    for i in L_plates:
        i = i.strip()
        if i[1].isdigit():
            L_plates2.append(i[0:1] + ' ' + i[1:])    
        else:
            L_plates2.append(i)
    L_plates3 = []
    for i in L_plates2:
        if i[5] != ' ' and i[5].isalpha():
            L_plates3.append(i[0:5] + ' ' + i[5:])
        else:
            L_plates3.append(i)
    L_plates4 = []
    for i in L_plates3:
        if i[-4] != ' ' and i[-4].isalpha():
            L_plates4.append(i[0:-3] + ' ' + i[-3:])
        else:
            L_plates4.append(i)
    # pprint(L_plates4)    
    
    # Running array of plates to match with Korotov col fleet 1 (col B)
    row = 2
    end = 48
    L_drivers = []
    while True:
        for i in L_plates4:
            if i in str(ws13.Cells(row, 2).Value):
                L_drivers.append(ws13.Cells(row, 3).Value)
            # else:
            #     L_drivers.append('Нет данных')   
                
        row += 1
       
        if row == end:
            break
        # if True add Col C value to the list
        # if False add Нет данных
    
    
    pprint(L_drivers)
    pprint(len(L_plates4))
    pprint(len(L_drivers))

if __name__ == '__main__':
    main()
