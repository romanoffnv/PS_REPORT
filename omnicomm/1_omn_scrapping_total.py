from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from time import sleep
import time
from pprint import pprint
import pandas as pd
import itertools
import sqlite3
import socket
import os

class Scrapper:
    def __init__(self):
        self.time = time
        self.wait = wait
        
        # Checking pc by verification of existing path
        laptopPath = 'C://Users//roman//OneDrive//Рабочий стол//SANDBOX//REPORT//Chrome_driver//chromedriver.exe'
        isExist = os.path.exists(laptopPath)
        if isExist:
            currPath = laptopPath
        else:
            currPath = 'D://Users//Sauron//Desktop//SANDBOX//chromedriver'
            
        # options
        self.options = webdriver.ChromeOptions()
        self.driver = webdriver.Chrome(
            executable_path = currPath
        )
        self.login = 'pakerservice'
        self.password = 'w2sRkVWZ'
        self.url = 'https://online.omnicomm.ru/mainpage'


        self.driver.get(self.url)
        self.time.sleep(5)
        self.login_input = self.driver.find_element(By.NAME, 'login')
        self.pass_input = self.driver.find_element(By.NAME, 'password')
        self.button = self.driver.find_element(By.XPATH, ('/html/body/div/div/div[2]/div[1]/form/div[4]/button'))
        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        self.login_input.send_keys(self.login)
        self.pass_input.send_keys(self.password)
        self.button.click()
        self.time.sleep(3)
        self.driver.get('https://online.omnicomm.ru/acreports')
        self.time.sleep(3)
        # wait until "местоположение" report card is clicable and click it
        self.wait(self.driver, 20).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/div/div[3]/div/div/div[1]/div/div[2]/div/div"))).click()
        self.time.sleep(5)
        self.searchField = self.driver.find_element(By.XPATH, ('/html/body/div/div/div[1]/div[2]/div[2]/div[2]/div[1]/section[1]/input'))
        self.selectAll = self.driver.find_element(By.XPATH, ("/html/body/div/div/div[1]/div[2]/div[2]/div[2]/div[1]/section[3]/ul[1]/li/div[1]/a"))

    def getInfo(self, key):
        if key == 'Без группы':
            # Crutch
            self.group = 'АДПМ ВВ6169 86'
        else:
            self.group = ''
            self.time.sleep(1)
            self.searchField.clear()
            self.time.sleep(1)
            self.searchField.send_keys(key)
            self.time.sleep(1)
            self.searchField.send_keys(Keys.ENTER)
            self.time.sleep(5)
            self.selectAll.click()
            self.time.sleep(3)
            self.group = self.driver.find_element(By.XPATH, '/html/body/div/div/div[3]/div/div/div/div/div/div/div[3]/div[2]/h3/span[2]').text
            self.selectAll.click()
        return self.group

    def postDb (self, key, group):
        self.L2 = group.split(',')
        self.L1 = [str(key) for x in range(len(self.L2))]
        return self.L1, self.L2

    ##Building data frame
    def buildDF(self, L_col1, L_col2):
        data = pd.DataFrame(zip(L_col1, L_col2))
        data.columns = ['Department', 'Vehicle']
        return data

    # Adding list of data into sqlite3 database
    def addToDB(self, L_col1, L_col2):
        db = sqlite3.connect('omnicomm.db')
        cursor = db.cursor()

        cursor.execute("DROP TABLE IF EXISTS total_vehicles;")
        cursor.execute("""
                    CREATE TABLE IF NOT EXISTS total_vehicles(
                    Department text,
                    Vehicle text
                            )
                       """)
        cursor.executemany("INSERT INTO total_vehicles VALUES (?, ?)", zip(L_col1, L_col2))
        cursor.execute("""
                    SELECT *
                    FROM total_vehicles
                    """)
        pprint(cursor.fetchall())
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        print(cursor.fetchall())
        # refreshing database
        db.commit()
        # closing database
        db.close()

session = Scrapper()
L_groups = [
            'TOYOTA', 
            'Аренда ТС от ООО "Югратехсрой"', 
            'ГНКТ 1',
            'ГНКТ 2', 
            'ГНКТ 3', 
            'ГНКТ 4', 
            'ГНКТ 5', 
            'ГНКТ 6', 
            'ГНКТ 7',
            'ГНКТ 8', 
            'ГНКТ 9', 
            'ГНКТ-10', 
            'ГНКТ-11', 
            'ГНКТ-14', 
            'ГНКТ-16', 
            'ГНКТ-17', 
            'ГНКТ-18', 
            'ГНКТ-19', 
            'ГНКТ-22', 
            'ГНКТ-31', 
            'ГНКТ-Резерв', 
            'ГРП 1', 
            'ГРП 2', 
            'ГРП 3', 
            'ГРП 4', 
            'ГРП 5', 
            'ГРП 6', 
            'ГРП 7', 
            'ГРП 8', 
            'ГРП 9',
            'ГРП резерв', 
            'ГРП РЕМОНТ', 
            'ГРП-14', 
            'ГРП-15', 
            'ГРП-16', 
            'ГРП-17', 
            'Дизельная электростанция',
            'Идентификация водителей ТС', 
            'Маз с739кр186 с Насосной установкой', 
            'ООО "Нефтемаш"', 
            'ООО "Югратехстрой"', 
            'ТР. Служба'
            
            ]
L_col1, L_col2 = [], []

for i in L_groups:
    group = session.getInfo(i)
    db = session.postDb(i, group)

    L_col1.append(db[0])
    L_col2.append(db[1])
    print('Processing ' + i)

L_col1 = list(itertools.chain.from_iterable(L_col1))
L_col2 = list(itertools.chain.from_iterable(L_col2))

df = session.buildDF(L_col1, L_col2)
db = session.addToDB(L_col1, L_col2)
print(df)
print(df.describe())
print('1_omn_scrapping_total.py is complete')