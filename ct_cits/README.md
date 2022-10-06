# 1_main_trucks.py

1. Get the Excel Application COM object
2. Making connection to DB
3. Adding crews to array

   ```
   'ГНКТ № 1',
   'ГНКТ № 2',
   'ГНКТ № 3',
   'ГНКТ № 4',
   ... ]
   ```
4. Adding main trucks to array

   ```
   [ 'Бурильщик',
   None,
   'Пом.бур',
   None,
   'Маш-т  НТ\nгос№\n8392 УМ 86',
   None,
   'Маш-т НКА\nгос№\nТТ 1559 54'

   .... ]
   ```
5. Cleaning main trucks

   ```
   ['',
   'УУ0775',
   'В315АМ',
   'АТ3824',
   '',
   '',
   'УА5848',
   'Н397КС',
   'АХ1895',
   '',
   '',

   ... ]
   ```
6. Collecting multiplitaction indeces to match up with crews

   ```
   [3, 3, 5, 3, 3, 3, 3, 5, 4, 3, 3, 3, 2, 3, 3, 3, 3, 2]
   ```
7. Multiplying crews to counts

   ```
   0    ГНКТ 1  уу0775
   1    ГНКТ 1  в315ам
   2    ГНКТ 1  ат3824
   3    ГНКТ 2  уа5848
   4    ГНКТ 2  н397кс
   5    ГНКТ 2  ах1895
   ```
8. Assigning duplicated groups to the plate numbers

   ```
   7            ГНКТ 3, ГНКТ 19  ах7262

   18  ГНКТ 31, ГНКТ 18, ГНКТ 6  р252ам
   ```
9. Making colums: plate indeces, plate literals

   ```
   0                     ГНКТ 1  уу0775  0775   уу
   1                     ГНКТ 1  в315ам   315  вам
   2                     ГНКТ 1  ат3824  3824   ат
   3                     ГНКТ 2  уа5848  5848   уа
   4                     ГНКТ 2  н397кс   397  нкс
   ```
10. Extracting table columns as lists from final_DB for matching against
11. Searching plates in omnicom by index and literal to build list of vehicle names and patch the plates from omnicomm

    ```
    0   ГНКТ 1     Волат  МЗКТ   0775уу 86
    1   ГНКТ 1     KENWORTH  в315ам 186
    ```
12. Patching plates and vehicle names not available in omnicomm
13. Building dataFrame

    ```
    0    ГНКТ 1    Волат  МЗКТ   0775уу 86  0775      уу
    1    ГНКТ 1    KENWORTH  в315ам 186   315     вам
    ```
14. Posting dataFrame to cits.db

# 2_aux_trucks.py
