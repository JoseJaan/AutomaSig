import os
from dotenv import load_dotenv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from time import sleep

browser = webdriver.Firefox()

#Read xlsx
excel_path = 'Database/schedules.xlsx'
df = pd.read_excel(excel_path)

#Open Browswer
browser.get('https://sig.ufla.br/modulos/login/index.php')

sleep(3)

#Login
load_dotenv()

userLogin = os.getenv('LOGIN')
userPass = os.getenv('PASSWORD')

loginPlace = browser.find_element(By.ID,'login')
loginPlace.send_keys(userLogin)

passPlace = browser.find_element(By.ID,'senha')
passPlace.send_keys(userPass)

submitPlace = browser.find_element(By.ID,'entrar')
submitPlace.submit()

sleep(3)

#Navegating
Bolsas = browser.find_element(By.LINK_TEXT,'Bolsas Institucionais')
Bolsas.click()

sleep(3)

Relatorio = browser.find_element(By.XPATH, '//a[@title="Relatório Mensal de Atividades"]')
Relatorio.click()
    
sleep(3)


browser.quit()