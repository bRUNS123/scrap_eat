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
time.sleep(4)

#Wait
wait = WebDriverWait(driver, 5, poll_frequency=1, ignored_exceptions=[ElementNotVisibleException, ElementNotSelectableException])


try:
    textoInvalido= wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mensajeSplashContainer"]/div')))
    
except:
    try: 
        textoInvalido= wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="mensajeSplashContainer"]/div')))
        clickMenu = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="toggleMenu"]'))).click()
        usuario = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="sideBar"]/div[2]/div/div[1]/span[2]'))).click() 
        cerrarSesion = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="sideBar"]/div[2]/div/div[2]/div[3]/span[2]'))).click()
        
        print('Elemento encontrado')
      
       
    except:
        print('No se encontro elemento')   

    

#Logueo automatico
try:
    revisarSiHayTabla = wait.until(EC.element_to_be_clickable((By.XPATH,' //*[@id="tablaDetallePagos"]/tbody/tr[1]/th[2]'))).is_displayed()

except:
    print('No hay tabla')
    clickMenu = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="toggleMenu"]'))).click()
    clickMenu = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="sideBar"]/div[2]/div/div[1]/span[3]'))).click()
    clickMenu = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="sideBar"]/div[2]/div/div[2]/div[4]/span[2]'))).click()
    textBoxClear = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="loginToteat"]/div[2]/div/div/input[1]'))).clear()
    passBoxClear = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="loginToteat"]/div[2]/div/div/input[2]'))).clear()
    textBox = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="loginToteat"]/div[2]/div/div/input[1]')))
    passBox = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="loginToteat"]/div[2]/div/div/input[2]')))
    textBox.send_keys('f.correa.cood@gmail.com')
    passBox.send_keys('Remotito1')
    botonClick =wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="loginToteat"]/div[2]/div/div/button'))).click()
    time.sleep(1.5)
    clickMenu = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="toggleMenu"]'))).click()
    menuIdioma = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="sideBar"]/div[11]/div/div[1]/span[3]'))).click() 
    espanol = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="sideBar"]/div[11]/div/div[2]/div[4]/span[1]'))).click()
    reportes = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="sideBar"]/div[9]/div[1]/span[3]'))).click()
    irReportes = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="sideBar"]/div[9]/div[2]/div[5]/span[2]'))).click()
    time.sleep(1.5)

#Crear carpetas
pathAnual = './anual'
pathDiario = './diario'
pathTodos = './todos'

if os.path.exists(pathAnual):
    print('Ya existe la carpeta')
else: 
   os.makedirs(pathAnual)
   
if os.path.exists(pathDiario):
    print('Ya existe la carpeta')
else: 
   os.makedirs(pathDiario)

if os.path.exists(pathTodos):
    print('Ya existe la carpeta')
else: 
   os.makedirs(pathTodos)




#Crear archivo de texto
try:
    f = open("./anual/listafechas.txt", "r", encoding='utf-8')
    f.close()
except:
    f = open("./anual/listafechas.txt", "w", encoding='utf-8')
    f.close()



#Recorrer Dias
k = 2
x = 0
while x < 1000:
    print(f'parte chekeo: {k}')  
    clickListaDias = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select'))).click()
    # clickDia = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select/option['+str(k)+']'))).click()
    DiaInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select/option['+str(k)+']'))).text

    if os.path.getsize('./anual/listafechas.txt') != 0:
        with open('./anual/listafechas.txt', encoding='utf-8') as file:
            last_line = file.readlines()[-1]
            def func(value):
              return ''.join(value.splitlines())
        
   
            if DiaInfo in func(last_line): 
            
                print('Son iguales')
                break
                


            else:
                print('No son iguales')
                k = k + 1
                x = x + 1  
                
       
       
                # x=1000
    else:
        print('No hay nada en el archivo')
        break

try:
    while k < 1000:
        clickListaDias = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select'))).click()
        clickDia = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select/option['+str(k)+']'))).click()
        DiaInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select/option['+str(k)+']'))).text
        time.sleep(1)

   

        if '2021' in DiaInfo:
            k = 1000

        #Formatear dia
        capitlizeDia = DiaInfo.title()
        formatearDia = capitlizeDia.replace(" ", "")
        sinComa = formatearDia.replace(",", "")
        diafinal = sinComa.replace('-TurnoUnico', '')
        print(diafinal)

        #Crear Excel
        # Create an new Excel file and add a worksheet.
        workbook = xlsxwriter.Workbook(f'./anual/{diafinal}.xlsx')
        worksheet = workbook.add_worksheet()   

        # Darle Ancho a las columnas
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

        #Lista de comandas (para no repetir)
        comandaList = []
        comandaListCompleta = []

        celda = 2
        j = 2


        try:
            itemFilaContador = 1
            i = 2
            while i < 1000:
                print(f'i es {i}')
    
                
                DiaInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[1]/div/div/div[3]/div/select/option['+str(k)+']'))).text
                diaInfo = DiaInfo.replace(' - Turno Unico', '')
                # diaInfoAnterior = DiaAnteriorInfo.replace(' - Turno Unico', '')
                worksheet.write(f'A{celda}', diaInfo)

            
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

                        comensalesInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="pgizq"]/div[1]/span[6]'))).text
                        garzonInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="pgizq"]/div[1]/span[4]'))).text
                        brutoInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="pgizq"]/div[2]/div[2]/div[2]/span'))).text

                        try:
                            itemContador = 1
                            j = 2
                            columnaItem = 6
                            while j < 1000:
            
                                itemsInfo = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[3]/div/div[2]/div[2]/div[2]/div['+str(itemContador)+']/div/span[2]'))).text 
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

                        # print(comensalesInfo)
                        time.sleep(0.5)
                        regresar = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="reportes"]/div[3]/div/div[1]/button[1]/span'))).click()
               
                        time.sleep(0.2)
                        celda += 1
                        i += 1
                        itemFilaContador += 1

                     else:
                        print('Comanda repetida')
                        i += 1

                except:
                    print('Ya existe la comanda')

        except: 
          print('Se termino el día') 


        k=k+1 
        f = open("./anual/listafechas.txt", "a", encoding='utf-8')
        f.write(DiaInfo+"\n")
        f.close()
        workbook.close()

           
except:
    print('listo')

    # driver.close()







  


