from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from time import sleep
import time
import os
from pprint import pprint
import pandas as pd
import itertools
import sqlite3
import socket

def main():
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

            self.selectRed = self.driver.find_element(By.XPATH, ('/html/body/div/div/div[1]/div[2]/div[2]/div[2]/div[1]/section[2]/div[3]/div/span'))
            self.selectGrey = self.driver.find_element(By.XPATH, ('/html/body/div/div/div[1]/div[2]/div[2]/div[2]/div[1]/section[2]/div[4]/div/span'))
            self.selectOrange = self.driver.find_element(By.XPATH, ('/html/body/div/div/div[1]/div[2]/div[2]/div[2]/div[1]/section[2]/div[2]/div/span'))
            self.selectAll = self.driver.find_element(By.XPATH, ("/html/body/div/div/div[1]/div[2]/div[2]/div[2]/div[1]/section[3]/ul[1]/li/div[1]/a"))

        def getInfo(self, key):
            print(key)
            if key == 'red':
                self.button = self.selectRed
            elif key == 'grey':
                self.button = self.selectGrey
            elif key == 'orange':
                self.button = self.selectOrange

            self.group = ''
            self.time.sleep(1)
            self.button.click()
            self.time.sleep(1)
            self.selectAll.click()
            self.time.sleep(2)
            self.group = self.driver.find_element(By.XPATH, '/html/body/div/div/div[3]/div/div/div/div/div/div/div[3]/div[2]/h3/span[2]').text
            self.selectAll.click()
            return self.group


        def postDb (self, key, group):
            self.L2 = group.split(',')
            self.L1 = [str(key) for x in range(len(self.L2))]
            return self.L1, self.L2

        ##Building data frame for red
        def buildDFRed(self, L_col1_red, L_col2_red):
            data = pd.DataFrame(zip(L_col1_red, L_col2_red))
            data.columns = ['Red', 'Vehicle']
            return data

        ##Building data frame for grey

        def buildDFGrey(self, L_col1_grey, L_col2_grey):
            data = pd.DataFrame(zip(L_col1_grey, L_col2_grey))
            data.columns = ['Grey', 'Vehicle']
            return data
        
        ##Building data frame for orange
        def buildDFOrange(self, L_col1_orange, L_col2_orange):
            data = pd.DataFrame(zip(L_col1_orange, L_col2_orange))
            data.columns = ['Orange', 'Vehicle']
            return data

         # Adding list of data into sqlite3 database
        def addToDB(self, L_col1_red, L_col2_red, L_col1_grey, L_col2_grey, L_col1_orange, L_col2_orange):


            db = sqlite3.connect('omnicomm.db')
            db.row_factory = lambda cursor, row: row[0]
            cursor = db.cursor()

            cursor.execute("DROP TABLE IF EXISTS red_vehicles;")
            cursor.execute("""
                        CREATE TABLE IF NOT EXISTS red_vehicles(
                        Red text,
                        Red_Vehicle text
                                )
                           """)
            cursor.executemany("INSERT INTO red_vehicles VALUES (?, ?)", zip(L_col1_red, L_col2_red))
            pprint(cursor.execute("SELECT * FROM red_vehicles").fetchall())
            
            cursor.execute("DROP TABLE IF EXISTS grey_vehicles;")
            cursor.execute("""
                        CREATE TABLE IF NOT EXISTS grey_vehicles(
                        Grey text,
                        Grey_Vehicle text
                                )
                           """)
            cursor.executemany("INSERT INTO grey_vehicles VALUES (?, ?)", zip(L_col1_grey, L_col2_grey))
            pprint(cursor.execute("SELECT * FROM grey_vehicles").fetchall())
            
            cursor.execute("""
                        CREATE TABLE IF NOT EXISTS orange_vehicles(
                        Orange text,
                        Orange_Vehicle text
                                )
                           """)
            cursor.executemany("INSERT INTO orange_vehicles VALUES (?, ?)", zip(L_col1_orange, L_col2_orange))
            pprint(cursor.execute("SELECT * FROM orange_vehicles").fetchall())
            
            print(cursor.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall())
            
            # refreshing database
            db.commit()
            # closing database
            db.close()


    session = Scrapper()

    L_col1_red, L_col2_red = [], []
    L_col1_grey, L_col2_grey = [], []
    L_col1_orange, L_col2_orange = [], []

    red = session.getInfo('red')
    grey = session.getInfo('grey')
    orange = session.getInfo('orange')

    db = session.postDb('red', red)

    L_col1_red.append(db[0])
    L_col2_red.append(db[1])

    L_col1_red = list(itertools.chain.from_iterable(L_col1_red))
    L_col2_red = list(itertools.chain.from_iterable(L_col2_red))

    df_red = session.buildDFRed(L_col1_red, L_col2_red)

    print(df_red)


    db = session.postDb('grey', grey)

    L_col1_grey.append(db[0])
    L_col2_grey.append(db[1])

    L_col1_grey = list(itertools.chain.from_iterable(L_col1_grey))
    L_col2_grey = list(itertools.chain.from_iterable(L_col2_grey))

    df_grey = session.buildDFGrey(L_col1_grey, L_col2_grey)

    print(df_grey)
    
    db = session.postDb('orange', orange)

    L_col1_orange.append(db[0])
    L_col2_orange.append(db[1])

    L_col1_orange = list(itertools.chain.from_iterable(L_col1_orange))
    L_col2_orange = list(itertools.chain.from_iterable(L_col2_orange))

    df_orange = session.buildDFOrange(L_col1_orange, L_col2_orange)

    print(df_orange)
    
    db = session.addToDB(L_col1_red, L_col2_red, L_col1_grey, L_col2_grey, L_col1_orange, L_col2_orange)

    print("2_omn_scrapping_nodata is complete")
if __name__ == '__main__':
    main()
