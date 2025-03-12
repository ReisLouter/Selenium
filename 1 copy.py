from selenium import webdriver
from selenium.webdriver.firefox.options import Options #firefox
from selenium.webdriver.firefox.service import Service #chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from firebase_admin import credentials, firestore, db
from openpyxl import load_workbook
import firebase_admin
import pyautogui
import time
import sys
import re


##Clases##
def encontrarId(id):
    global driver
    element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, f"{id}"))
    )
    # Interact with the element here
    return element
def noencontrarXpath(xpath):
    global driver
    contador = 1
    while True:
        try:
            element = WebDriverWait(driver, 20).until_not(
                
            EC.presence_of_element_located((By.XPATH, f"{xpath}"))
            
            
            )
            
            # Interact with the element here
            return element
        except TimeoutException:
            contador+=1
            if contador >3:
                driver.close()
                sys.exit()
            driver.refresh()
def encontrarXpath(xpath):
    global driver
    contador = 1
    while True:
        try:
            element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, f"{xpath}"))
            )
            # Interact with the element here
            return element
        except TimeoutException:
            contador+=1
            if contador >3:
                driver.close()
                sys.exit()
            driver.refresh()
        
##Fin - Clases##

##Firebase##
cred_obj = credentials.Certificate(
    'keys.json')
firebase_admin.initialize_app(cred_obj)
bd = firestore.client()
a = bd.collection("aplicacionesOma").document("clave").get().to_dict()
if a["clv"] != True:
    sys.exit()
##Fin - firebase##

##variables##
a = "/html/body/app-root/div/div/app-prov-ficha/div/div/div[1]/div[3]/div/div[1]/div[2]/div/div/div/span/img" 
b = "/html/body/app-root/div/div/app-prov-ficha/div/div/div[1]/div[3]/div[1]/app-loading-icon/div"
c = "/html/body/app-root/div/div/app-prov-buscador/div[3]/div[1]/div[2]/div[2]/div/app-tile/a/div/div[1]/div[2]"
opciones = Options()
opciones.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
service = Service(r'D:\seleniunDrivers\geckodriver.exe')
driver = webdriver.Firefox(service=service, options=opciones)
ruc = ["15612117672","", ""]
##Fin - variables##


##Excel##
df = load_workbook("BASE GENERAL DEL PERSONAL  -  CLAVE SOL 2024 MELO.xlsx")["BASE- PERSONAL"]
hj = df["A"]
##Fin - Excel##
for i in hj[1:]:
    # if len((i.value)) == 8:
        




    driver.get(f"https://apps.osce.gob.pe/perfilprov-ui/")
    time.sleep(2)
    busqueda = encontrarId("textBuscar")
    busqueda.send_keys(i.value)
    busqueda = encontrarId("btnBuscar")
    busqueda.click()

    busqueda = encontrarXpath(c)
    busqueda.click()
    var = encontrarXpath(a)

    if var:
        time.sleep(1)
        driver.save_screenshot(f'imagenes\\alimenticia\\{i.value}.png')
        
        
    
    
    
    



driver.quit()
sys.exit()

