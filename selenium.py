from selenium import webdriver
from selenium.webdriver.firefox.options import Options #firefox, chrome cambiar segun los gustos
from selenium.webdriver.firefox.service import Service 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

#añadir firestore para controlar el uso de la aplicacion
from firebase_admin import credentials, firestore
import firebase_admin

#Añadir openpyxl para grandes entradas de datos (controlar excel)
from openpyxl import load_workbook

#Utileria
import time
import sys



##Clases## 
""" Acciones repetitivas de la libeia selenium. pasadoas a funciones para una codificacion mlimpia mas adelante """
def encontrarId(id):
    global driver #<- señalamos el uso del navegador qen que se buscaran los objetos
    element = WebDriverWait(driver, 10).until( #<- agregamos una seccion para que la pagina busque el objeto en un rango determinado de tiempo.
    EC.presence_of_element_located((By.ID, f"{id}")) #<- busqueda del objeto
    )
    # regresamos el elemento encontrado. (todas las clases utilizan la misma estructura)
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
"""" Control Remoto par la aplicacion. """
cred_obj = credentials.Certificate(
    'keys.json')
firebase_admin.initialize_app(cred_obj)
bd = firestore.client()
a = bd.collection("aplicacionesOma").document("clave").get().to_dict()
if a["clv"] != True:
    sys.exit()
##Fin - firebase##

##variables##
""" redirecciones, variables y xpath para un uso limpio posteriormente """
a = "/html/body/app-root/div/div/app-prov-ficha/div/div/div[1]/div[3]/div/div[1]/div[2]/div/div/div/span/img" 
b = "/html/body/app-root/div/div/app-prov-ficha/div/div/div[1]/div[3]/div[1]/app-loading-icon/div"
c = "/html/body/app-root/div/div/app-prov-buscador/div[3]/div[1]/div[2]/div[2]/div/app-tile/a/div/div[1]/div[2]"
opciones = Options()
opciones.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
service = Service(r'D:\seleniunDrivers\geckodriver.exe')
driver = webdriver.Firefox(service=service, options=opciones) #El navegador señalado debe ser le mismo que al declarar las librerias

##Fin - variables##


##Excel##
""" creacion de lista para busqueda por ruc """
df = load_workbook("BASE GENERAL DEL PERSONAL.xlsx")["BASE- PERSONAL"]
hj = df["A"]
##Fin - Excel##




""" Nos saltamos la cabeceza """
for i in hj[1:]:
    # if len((i.value)) == 8:
        
    driver.get(f"https://apps.osce.gob.pe/perfilprov-ui/") #get = ir a X url
    time.sleep(2)
    busqueda = encontrarId("textBuscar")
    busqueda.send_keys(i.value) #send_key = Escribir o precionar tecla en el elemento encontrado
    busqueda = encontrarId("btnBuscar")
    busqueda.click() #click = enviar un press (o click) al elemento encontrado

    busqueda = encontrarXpath(c)
    busqueda.click()
    
    var = encontrarXpath(a) #-> esperar a que cargue la pagina correctamente

    if var:
        
        time.sleep(1)
        driver.save_screenshot(f'imagenes\\alimenticia\\{i.value}.png') #Tomar captura y guardarla
        
        
    
    
    
    



driver.quit() # Cerrrar naavegador
sys.exit() #terminar ejecucion

