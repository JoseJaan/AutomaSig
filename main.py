import os
from dotenv import load_dotenv
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

browser = webdriver.Firefox()

#Open Browswer
browser.get('https://sig.ufla.br/modulos/login/index.php')

sleep(2)

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

sleep(2)

#Changing tabs
Bolsas = browser.find_element(By.LINK_TEXT,'Bolsas Institucionais')
Bolsas.click()

sleep(2)

Relatorio = browser.find_element(By.XPATH, '//a[@title="Relatório Mensal de Atividades"]')
Relatorio.click()
    
sleep(2)

#Open report
month = 'Julho'

xpath = f"//tr[td[text()='{month}']]//a[@title='Definir Atividades do Bolsista Institucional']"
defineActivities = browser.find_element(By.XPATH, xpath)
defineActivities.click()

sleep(2)

#Read xlsx
data = pd.read_excel('Database/schedules.xlsx')

#Insert data
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
    
    #zfill is used to guarantee that the date will have two digits
    day.send_keys(str(row['Day']).zfill(2))
    init_hour.send_keys(row['InitHour'].strftime('%H:%M'))
    end_hour.send_keys(row['EndHour'].strftime('%H:%M'))
    place.send_keys(row['Place'])
    description.send_keys(row['Description'])
    
    if(index < len(data)-1):
        insertNewDate = browser.find_element(By.CLASS_NAME,'adicionar')
        insertNewDate.click()

    sleep(2)

#Saving changes
save = browser.find_element(By.NAME,'alterar')
save.click()

sleep(2)

browser.quit()