# -*- coding: utf-8 -*-

import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import ElementNotVisibleException, ElementNotSelectableException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import xlsxwriter
from selenium.webdriver.common.keys import Keys

import time
import os
import sys

#Variables
Chrome_Dir = ".\chromedriver.exe"
Url = "https://res3.toteat.com/#/reportes/detallepagos"
#f.correa.cood@gmail.com
#Remotito1

#Relatives Paths
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

#Iniciar Chrome y grabar perfil
dir_path = os.getcwd()
profile = os.path.join(dir_path, "profile", "wpp")
s=Service(resource_path('./driver/chromedriver.exe'))
op = webdriver.ChromeOptions()
op.add_experimental_option('excludeSwitches', ['enable-logging'])
op.add_argument(
    r"user-data-dir={}".format(profile))
driver = webdriver.Chrome(service=s, options=op)
driver.get(Url)
time.sleep(5)

#Wait
wait = WebDriverWait(driver, 5, poll_frequency=1, ignored_exceptions=[ElementNotVisibleException, ElementNotSelectableException])

#Crear Excel
# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('dato.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 32)
worksheet.set_column('B:B', 12)
worksheet.set_column('C:C', 13)
worksheet.set_column('D:D', 28)
worksheet.set_column('E:E', 25)
worksheet.set_column('G:G', 20)






# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Write some simple text.
worksheet.write('A1', 'Dia', bold)
worksheet.write('B1', 'Comensales', bold)
worksheet.write('C1', 'Venta Bruto', bold)
worksheet.write('D1', 'Garzon', bold)
worksheet.write('E1', 'Comanda', bold)
worksheet.write('F1', 'Mesa', bold)
# worksheet.write(0, 6, 'Item1', bold)
# worksheet.write(0, 7, 'Item2', bold)



#Lista de comandas (para no repetir)
comandaList = []
comandaListCompleta = []

#Ir a la comanda anterior.
clickDias = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select'))).click()
clickDiaAnterior = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select/option[3]'))).click()
DiaAnteriorInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select/option[3]'))).text

time.sleep(2)



# Text with formatting.
# worksheet.write('A2', 'VARIABLE DIA', bold)

# Write some numbers, with row/column notation.
# worksheet.write(2, 0, 123)
# worksheet.write(3, 0, 123.456)

# Insert an image.
# worksheet.insert_image('B5', 'logo.png')


celda = 2
j = 2
# items = G

try:
    itemFilaContador = 1
    i = 2
    while i < 1000:
        print(f'i es {i}')
    
        diaInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select/option[2]'))).text
        diaInfo = diaInfo.replace(' - Turno Unico', '')
        diaInfoAnterior = DiaAnteriorInfo.replace(' - Turno Unico', '')

        worksheet.write(f'A{celda}', diaInfoAnterior)

      
        #Información general
        comandaInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="tablaDetallePagos"]/tbody/tr['+str(i)+']/td[2]'))).text
        mesaInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="tablaDetallePagos"]/tbody/tr['+str(i)+']/td[4]'))).text

        
        #Si la comanda no esta en la lista, agregarla.
        try:
            if comandaInfo not in comandaList:
          
        #Si la comanda no esta en la lista, clickearla
                comandaClick = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="tablaDetallePagos"]/tbody/tr['+str(i)+']/td[2]'))).click()
            
                comandaList.append(comandaInfo)
                print(f'Comanda: {comandaList}')
                



















        except:
            print('Ya existe la comanda')
        #Click comanda
        # if comandaInfo in comandaList:
        #     comandaClick = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="tablaDetallePagos"]/tbody/tr['+str(i)+']/td[2]'))).click()

            

        




 

        #Información comanda
        comensalesInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="pgizq"]/div[1]/span[6]'))).text
        garzonInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="pgizq"]/div[1]/span[4]'))).text
        brutoInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="pgizq"]/div[2]/div[2]/div[2]/span'))).text
      
        #Información Items Debe ir con un TRY y un for
        #//*[@id="reportes"]/div[3]/div/div[2]/div[2]/div[2]/div[1]/div/span[2]
        #//*[@id="reportes"]/div[3]/div/div[2]/div[2]/div[2]/div[2]/div/span[2]
        try:
            itemContador = 1
            j = 2
            columnaItem = 6
            while j < 1000:
            
              itemsInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[3]/div/div[2]/div[2]/div[2]/div['+str(itemContador)+']/div/span[2]'))).text 
              print(itemsInfo)
              worksheet.write(itemFilaContador, columnaItem, itemsInfo)
              worksheet.write(0, columnaItem, f'Item{itemContador}', bold)
              columnaItem = columnaItem+1
              itemContador = itemContador+1
              j = j+1
       
            

        except:
            print('no más items')

               
        worksheet.write(f'B{celda}', comensalesInfo)
        worksheet.write(f'C{celda}', brutoInfo)
        worksheet.write(f'D{celda}', garzonInfo)
        worksheet.write(f'E{celda}', comandaInfo)
        worksheet.write(f'F{celda}', mesaInfo)
        #Incrementar letra.
       
        print(comensalesInfo)
        
        time.sleep(0.5)
        regresar = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[3]/div/div[1]/button[1]/span'))).click()
                # driver.get(Url)
        time.sleep(0.5)

        celda += 1
        i += 1
        itemFilaContador += 1
     
except:
    print('listo')
    workbook.close()

    # driver.close()







  


