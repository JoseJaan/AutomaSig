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

#Browsing
Bolsas = browser.find_element(By.LINK_TEXT,'Bolsas Institucionais')
Bolsas.click()

sleep(3)

Relatorio = browser.find_element(By.XPATH, '//a[@title="Relatório Mensal de Atividades"]')
Relatorio.click()
    
sleep(3)

#Open report
month = 'Julho'

xpath = f"//tr[td[text()='{month}']]//a[@title='Definir Atividades do Bolsista Institucional']"
defineActivities = browser.find_element(By.XPATH, xpath)
defineActivities.click()

sleep(3)

#Read xlsx
excel_path = 'Database/schedules.xlsx'
data = pd.read_excel(excel_path)

for index, row in data.iterrows():
    
    row_number = index + 1
    
    xpath_day = f'//tr[{row_number}]/td[@class="cel_data"]/input[@name="datas[]"]'
    xpath_init_hour = f'//tr[{row_number}]/td[@class="cel_horas_inicio"]/input[@name="horas_inicio[]"]'
    xpath_end_hour = f'//tr[{row_number}]/td[@class="cel_horas_termino"]/input[@name="horas_termino[]"]'
    xpath_place = f'//tr[{row_number}]/td[@class="cel_local"]/input[@name="local_atividade[]"]'
    xpath_description = f'//tr[{row_number}]/td[@class="cel_descricao"]/input[@name="atividades[]"]'
    
    day = browser.find_element(By.XPATH, xpath_day)
    day.clear()
    init_hour = browser.find_element(By.XPATH, xpath_init_hour)
    init_hour.clear()
    end_hour = browser.find_element(By.XPATH, xpath_end_hour)
    end_hour.clear()
    place = browser.find_element(By.XPATH, xpath_place)
    description = browser.find_element(By.XPATH, xpath_description)
    
    day.send_keys(row['Day'])
    init_hour.send_keys(row['InitHour'].strftime('%H:%M'))
    end_hour.send_keys(row['EndHour'].strftime('%H:%M'))
    place.send_keys(row['Place'])
    description.send_keys(row['Description'])
    
    insertNewDate = browser.find_element(By.CLASS_NAME,'adicionar')
    insertNewDate.click()

    sleep(3)

browser.quit()