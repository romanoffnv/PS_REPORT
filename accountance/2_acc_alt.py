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


# Pandas
pd.set_option('display.max_rows', None)

# db connections
db = sqlite3.connect('data.db')
db.row_factory = lambda cursor, row: row[0]
cursor = db.cursor()
pd.set_option('display.max_rows', None)

def main():
    # Get accountance_1 db cols as lists L_mols and L_units
    L_mols = cursor.execute("SELECT Mol FROM accountance_1").fetchall()
    L_units = cursor.execute("SELECT Unit FROM accountance_1").fetchall()

    # Build df and filter out unneeded items
    
    
    # L_keywords = ['pH-метр', 'PH-метр', 'лабораторный', 'Shinko', 'Вискозиметр', 'Мешалка', 'Пресс', 'диагност', 'видео', 'сварочный', 'моечная', 'индукционный',
    #                'Станок', 'Стенд', 'стенд', 'Съемник', 'замок', 'фильтровентиляционный', 'труболовка', 'ясс', 'ловильн', 'компрессор', 'агнит', 
    #                'овитель труб', 'Металлошламоуловитель', 'Овершот', 'Труболовка', 'Труборез', 'Удлинитель', 'ловитель', 'ударно-вращатель', 'Фрез',
    #                'Штанголовка', 'Яс', 'Core', 'SuperМicro', 'Xerox', 'Kyocera', 'МФУ', 'принтер', 'Ноутбук', 'Радиотелефон', 'Сервер', 'вентиляц', 
    #                'Системный блок', 'Телевизор', 'Вертлюг', 'Анализатор', 'Блок манифольда', 'Катушка', 'орпус', 'Локатор', 'Лубрикатор', 'Мотопомпа', 
    #                'Репитер', 'Узел установки намотки', 'Нежилое здание', 'Мойка', 'Земельный участок', 'Сварочный аппарат', 'Стационарный пункт охраны', 'Весы',
    #                'Сварочное устройство', 'комплексный прибор', 'многозондовый влагомер', 'Автономный модуль', 'шумомер', 'цифровой манометр', 'механизм намотки трубы',
    #                'влагомер', 'Шкаф', 'Ангар', 'Баня водная', 'Waring', 'WARING', 'Warning', 'сепаратор', 'стола ротора', 'стол ротора', 'Волокуша', 'Газоанализатор',
    #                'Устьевой герметизатор', 'сброса шаров', 'прокатки труб', 'насыпной плотности', 'запасовки геофизического', 'Головка разводная', 'Датчик',
    #                'Дефектоскоп', 'Дозиметр', 'ИБП', 'Инвентор', 'микроскоп', 'манометр', 'Клапан редукционный', 'Коллектор геофизический', 'Комплект оборудования',
    #                'фланцевое соединение', 'Просеивающая машина', 'Расходомер', 'Измеритель плотности', 'Пакер', 'пакер', 'Перфоратор', 'Превентор', 'Резак', 
    #                'Установка для ремонта БДТ', 'Узел намотки', 'Турникет', 'Тренажер', 'Трамбователь', 'Точечный источник', 'Толщинометр', 'ележка', 'Твердомер',
    #                'Стол письменный', 'Стингер', 'Стеллаж', 'Система', 'Сборная вышка', 'еометр', 'Резьбо-нарезная', 'Радиостанция', 'Пробоотборник', 'Приёмное оборудование',
    #                'Преобразователь давл', 'Позиционный станок', 'Подвеска', 'Площадка приустьевая', 'Плазменный резак', 'Нежилое помещение', 'Модуль усилия', 'Модуль рахсодомера',
    #                'Модуль бытовой', 'Инвертор']
    # for i in L_keywords:
    #     df = df[df["Units"].str.contains(str(i)) == False]
    
    # # Listing filtered df
    # L_mols = df['Mols'].tolist()
    # L_units = df['Units'].tolist()
    
    # pprint(L_units)

    def splitter(split, L):
      L = [str(x) for x in L]  
      L = [x.split(split) for x in L]
      L = list(itertools.chain.from_iterable(L))
    #   L = [x.strip() for x in L if sum(map(str.isalpha, x)) > 1]
      return L

    L_keywords = ['г/н', 
                '43118', 'г.н.'
                'г.н.', 'гн', '(', 'г/р', ';', '43118', 'Гос.№', ',', 'зав.', 'мод.', 'зав', '№', ')', 'ст ', 'Г/н', 
                'АЦН'
                ]
    L_plates = splitter(L_keywords[0], L_units)
    for i in L_keywords:
        L_plates = splitter(str(i), L_plates)
    L_plates = [x.strip() for x in L_plates if sum(map(str.isalpha, x)) < 4]
    df = pd.DataFrame(zip(L_plates), columns=['Plates'])
    pprint(df)
    
    # for i in L_plates:
    #     pprint(i)
    #     pprint(sum(map(str.isalpha, i)) == 2)
    # pprint(len(L_plates))

    # L_units = [str(x) for x in L_units]
    # L_units = [x.split('(') if x != None else 'None' for x in L_units]
    # L_units = list(itertools.chain.from_iterable(L_units))

    # def slicer(side, x, L_units):
    #     L_units_temp = []
    #     if side == 'left':
    #         for i in L_units:
    #             if x in i:
    #                 ind = i.index(x)
    #                 L_units_temp.append(i[ind:])
    #             else:
    #                 L_units_temp.append(i)
    #         return L_units_temp
    #     elif side == 'right':
    #         for i in L_units:
    #             if x in i:
    #                 ind = i.index(x)
    #                 L_units_temp.append(i[:ind])
    #             else:
    #                 L_units_temp.append(i)
    #         return L_units_temp
        

    # # Slice from left
    # L_plates = slicer('left', 'г/н', L_units)
    # L_keywords = ['г.н.', 
    #                 'гос.№', 
    #                 'гос№', 
    #                 'гос. №', 
    #                 'Truck', 
    #                 'г/р',
    #                 'гн',
    #                 '43118',
    #                 'зав.',
    #                 'зав №'
    #                 ]
    # for i in L_keywords:
    #     L_plates = slicer('left', str(i), L_plates) 
    
    
    # # Slice from right
    # L_keywords = [',', ';']
    # for i in L_keywords:
    #     L_plates = slicer('right', str(i), L_plates) 
    # pprint(L_plates)
    
    # Build dataframe and filtering df
    # df = pd.DataFrame(zip(L_units, L_plates), columns=['Units', 'Plates'])
    # L_keywords = ['pH-метр', 'Овершот', 'Мешалка', 'Прибор', 'Станок', 
    #             'Безопасный замок', 'резьба', 'труболовка', 
    #             'стенд', 'Гидроясс', 'ловильн', 'видео', 'еханический',
    #             'механизм намотки трубы', 'Штанголовка', 'Шламоуловитель',
    #             'металлошламоуловитель', 'Фрез', 'герметизатор', 'ударно-вращательное',
    #             'Труборез', 'Труболовка', 'Тележка', 'Телевизор', 'Съемник', 'суперяс',
    #             'Стеллаж', 'правый', 'левый', 'диаметром', 'коммутатор', 'маршрутизатор',
    #             'внутр.d', 'правостор', 'левостор', 'агнит', 'овитель', 'шламоуловит',
    #             'гидропаук', 'мм\)', 'компьютер', 'МФУ', 'Ноутбук', 'принтер', 
    #             'Xerox', 'Сервер', 'Системный блок', 'Вертлюг', 'Катушка переводная',
    #             'Лубрикатор', 'Катушка переходная', 'Радиотелефон', 'спутниковый',
    #             'Xian Landrill', 'противовыбросовый', 'Преобразователь',
    #             'Резьбо-нарезная', 'Сварочный аппарат', 'OD\)', 'Intel', 'Резак скважинный',
    #             'азоанализатор', 'Пакер', 'Репитер', 'Клапан редукционный', 'Уголки высоко']
    # for i in L_keywords:
    #     df = df[df["Units"].str.contains(str(i)) == False]
    # pprint(df)
    
    # # Populate list of keywords for df filtration
    # L_units_filter = ['г/н', 'гос.№', 'гос№', 'гос. №', 'Truck', 'VIN', 'Насосная установка', 
    #                 'Mercedes', 'KENWORTH', 'Передвижная паровая установка', 'ППУ', 
    #                 'Полуприцеп', 'прицеп', 'тягач', 'Кран', 'Гидратационная установка', 'Автоцистерна', 'смеситель',
    #                 'блендер', 'КАМАЗ', 'Камаз', ' гн ']
    
    # # Filter df by keywords
    # data = data[data['Item'].str.contains('|'.join(L_units_filter))]
    
    # # Destructure df into lists
    # L_units = data.loc[:, 'Item'].tolist()
    # L_mols = data.loc[:, 'Responsible'].tolist()
    # # Derive the list of untouchable units to post it into db later
    # L_units_original = [x for x in L_units] 
      
    # # Slice items starting from the keyword's index to the end of the sentence
    # def slicer (x, L_units):
    #     L_units_temp = []
    #     if x == 'VIN' or x == 'vin' or x == 'VIV':
    #         for i in L_units:
    #             if x in i:
    #                 ind = i.index(x)
    #                 L_units_temp.append(i[:ind])
    #             else:
    #                 L_units_temp.append(i)
            
    #         L_units = [str(x).strip() for x in L_units_temp]
    #         L_units_temp.clear() 
            
    #         return L_units
    #     else:
    #         for i in L_units:
    #             if x in i:
    #                 ind = i.index(x)
    #                 L_units_temp.append(i[ind:])
    #             else:
    #                 L_units_temp.append(i)
            
    #         L_units = [str(x).strip() for x in L_units_temp]
    #         L_units_temp.clear() 
            
    #         return L_units

    # L_keywords = ['г/н', '№', ' гн ', 'г/р', 'г.н.', 'Г/н', 'Truck', 'VIN', 'vin', 'VIV']
    # for i in L_keywords:
    #     L_units = slicer(i, L_units)

    #     # Remove crap like 'г/н' etc
    # def crapRemover(x, L_units):
    #     L_units = [i.replace(x, '') for i in L_units]
    #     return L_units

    # L_keywords = ['г/н', '№', 'гн ', 'г.н.', 'Truck', 'г/р', 'Г/н', ')', ' ', 'RUS', ';', ',', ':']
    # for i in L_keywords:
    #     L_units = crapRemover(i, L_units)

    # # Fishing out plates by regex from long sentences
    # def regexBomber(x, L_units):
        
    #     L_plates_temp = []
    #     for i in L_units:
    #         if re.findall(x, str(i)):
    #             L_plates_temp.append(''.join(re.findall(x, str(i))))
    #         else:
    #             L_plates_temp.append(i)
    #             # print(i)
    
    #     L_units = [str(x).strip() for x in L_plates_temp]
    #     L_plates_temp.clear() 
            
    #     return L_units

    # L_units = regexBomber(re.compile('\s\D\s*\d+\D{2}\s*\d+'), L_units)
    # L_units = regexBomber(re.compile('\D{1}\s*\d+\s*\D{2}\s*\d+'), L_units)

    # # Remove paranthesis and content
    # L_patterns = ['\(.*\)', '\(.*', '\-']
    # for i in L_patterns:
    #     L_units = [''.join(re.sub(i, '', x)).strip() for x in L_units]

    # # Remove pointless sentences
    # L_units = [x if len(x) < 20 else '' for x in L_units]
       
    # # Converting leading region into the trailing region in plates (i.e. '86УК7801': 'УК780186',)
    # L_units_temp = []
    # L_units = [''.join(re.sub('\s', '', x)).strip() for x in L_units]
    # for i in L_units:
    #     try:
    #         if i[:1].isdigit() and str(i[2]).isalpha():
    #             L_units_temp.append(i[2:] + i[:2])
    #         else:
    #             L_units_temp.append(i)
    #     except IndexError:
    #         L_units_temp.append(i)

    # L_units = [x for x in L_units_temp]
    
    # # Turn plates into 123abc type
    # def transform_plates(plates):
    #     L_regions_long = [126, 156, 158, 174, 186, 188, 196, 797]
    #     L_regions_short = ['01', '02', '03', '04', '05', '06', '07', '09']
    #     for i in L_regions_long:
    #         plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 9 else x for x in plates]
    #     for i in L_regions_short:
    #         plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
    #     for i in range(10, 100):
    #         plates = [x.removesuffix(str(i)).strip() if x != None and len(x) == 8 or 'kzн' in str(x) else x for x in plates]
        
    #     plates_numeric = [''.join(re.findall(r'\d+', x)).lower() for x in plates if x != None]
    #     plates_literal = [''.join(re.findall(r'\D', x)).lower() for x in plates if x != None]
    #     plates = [str(x) + str(y) for x, y in zip(plates_numeric, plates_literal)]
    #     plates = [''.join(re.sub(r'\s+', '', x)).lower() for x in plates if x != None]
    #     return plates
    
    
    # L_PI_acc = transform_plates(L_units) 
     
    # # Crutch replacements for PI
    # D_crutches = {
    #                     '13621-наприцепетра': '',
    #                     '5668автоцистернаск': '5668ск',
    #                     '65221тягачкамаз': '652внт',
    #                     '2502/': '',
    #                     '2501/': '',
    #                     '300полуприцепм.рс-': '300',
    #                     '25001-к-': '',
    #                     '461bhp': '461внр',
    #                     '66577ус': '6657ус'
                        
    #         }        
        
    # # Replacing crappy unit names into omnicomm smth
    # for k, v in D_crutches.items():
    #     L_PI_acc = [x.replace(k, v) for x in L_PI_acc]

   
    # df = pd.DataFrame(zip(L_mols, L_units_original, L_units, L_PI_acc), columns=['Mols', 'Units', 'Plates', 'PI'])
    
    # # Posting df to DB
    # print('Posting df to DB')
    # cursor.execute("DROP TABLE IF EXISTS accountance_2")
    # df.to_sql(name='accountance_2', con=db, if_exists='replace', index=False)
    # db.commit()
    # db.close()
    
if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
