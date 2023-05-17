import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
import sys
import os


# -----------------VARIABLES DE ESPERA---------------------------
# time_load = 30
TIME_LOAD = 10
PATH_EXE = os.getcwd()
TIME_SEARCH_NIT = 3
# -----------------VARIABLES DE ESPERA---------------------------

# CONFIGURACION DE EXCEL
nombre_archivo = input('Ingrese el nombre del archivo: \n')
hoja = input('Ingresa el nombre de la hoja: \n')

archivo_path = os.path.join(PATH_EXE,r"Archivos", nombre_archivo)

print(archivo_path)

# with open(archivo_path) as f:
#     archivo = pd.read_excel(f, sheet_name=hoja, usecols=['tdoc', 'nid', 'dv', 'apl1', 'apl2', 'nom1', 'nom2', 'raz'])

archivo = "Archivos/" + nombre_archivo
archivo = pd.read_excel(os.path.join(PATH_EXE,archivo),sheet_name=hoja)
lista_verificacion = []
Lista_Error = []
lista_validacion_DV = []
# CONFIGURACION DE EXCEL

# LIMPIAMOS LAS COLUMNAS BASES A TRABAJAR
try:
    archivo = archivo.loc[:,['tdoc','nid','dv','apl1','apl2','nom1','nom2','raz']]
except:
    print('EXISTE UN PROBLEMA CON EL NOMBRE DE LAS COLUMNAS')
    print('CONTACTA CON TU DESARROLLADOR!!!!')
    print('Presiona ENTER para finalizar!')
    input()
    sys.exit()
# LIMPIAMOS LAS COLUMNAS BASES A TRABAJAR

# ELIMINAMOS LA INFORMACION QUE NO ES NIT
archivo = archivo[archivo['tdoc'] == 31]
# ELIMINAMOS LA INFORMACION QUE NO ES NIT


# validacion espacio
archivo = archivo.applymap((lambda x: " ".join(x.split()) if type(x) is str else x))
# validacion espacio

# Configuraciones para descarga
chrome_options = webdriver.ChromeOptions()
preferences = {"download.default_directory" : PATH_EXE ,
"directory_upgrade": True,
"safebrowsing.enable": True}
chrome_options.add_experimental_option("prefs",preferences)

# Cargamos el driver
driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=os.path.join(PATH_EXE,"Driver\chromedriver.exe"))

# Mostramos la pagina
driver.maximize_window()
url = "https://muisca.dian.gov.co/WebRutMuisca/DefConsultaEstadoRUT.faces"
driver.get(url)


# Esperamos la carga
wait = WebDriverWait(driver,TIME_SEARCH_NIT * 2)
wait.until(ec.visibility_of_element_located((By.ID,'vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit')))

# Validamos el df para buscar vacios
for column in archivo:
    archivo[""+str(column)+""].fillna('0',inplace=True)


for index, row in archivo.iterrows():
    # Buscamos la información de NIT
    if (row['tdoc'] == 31):
        banderapaso = False
        txt_nit = driver.find_element("xpath", '//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT:numNit"]')
        txt_nit.clear()
        txt_nit.send_keys(int(row['nid']))               
        txt_nit.send_keys(Keys.ENTER)

        try:
            wait = wait = WebDriverWait(driver,2)
            wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="divMensaje"]/table/tbody/tr[2]/td[2]/div/img')))
            salir = driver.find_element("xpath", '//*[@id="divMensaje"]/table/tbody/tr[2]/td[2]/div/img')
            salir.click()
            banderapaso = False
            print('NIT {} Con problemas'.format(int(row['nid'])))
        except:         
            # Buscamos La razon social
            wait = WebDriverWait(driver,TIME_LOAD)
            try:
                wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT"]/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]')))
                wait.until(ec.text_to_be_present_in_element((By.XPATH, '//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT"]/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]'),"Razón Social"))
                banderapaso = True
            except:
                banderapaso = False

        if (banderapaso == True):
            try:
                razon_social = driver.find_element("xpath", '//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT"]/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr/td[2]')
                DV = driver.find_element("xpath", '//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT:dv"]')
            except Exception as e:
                print(e)
                razon_social = ''

            # Validamos la existencia
            if (razon_social.text.replace(".","").strip() == ''):
                razon_social = 'NO ENCONTRADO'
            else:
                razon_social = razon_social.text.replace(".","").strip()

            # Validamos que sea igual
            if (row['dv'] == int(DV.text)):
                lista_validacion_DV.append('IGUAL')
            else:
                lista_validacion_DV.append(DV.text)

            # Añado a la lista
            lista_verificacion.append(razon_social)
            if (row['raz'].replace(".","").strip() == razon_social):
                Lista_Error.append('IGUAL')
            else:
                Lista_Error.append('ERROR')


        elif (banderapaso == False):
            # Buscamos La razon social
            wait = WebDriverWait(driver,TIME_LOAD)
           # wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT"]/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[1]')))
            try:
                wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT"]/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[1]')))
                wait.until(ec.text_to_be_present_in_element((By.XPATH, '//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT"]/table[2]/tbody/tr[2]/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[1]'),"Primer Apellido"))
                Newbanderapaso = True
            except:
                Newbanderapaso = False

            if (Newbanderapaso == True):
                primer_apellido = driver.find_element("xpath", '//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerApellido"]')
                segundo_apellido = driver.find_element("xpath", '//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT:segundoApellido"]')
                primer_nombre = driver.find_element("xpath", '//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT:primerNombre"]')
                otros_nombres = driver.find_element("xpath", '//*[@id="vistaConsultaEstadoRUT:formConsultaEstadoRUT:otrosNombres"]')

                resultado = primer_apellido.text.replace(".","").strip() + " " + segundo_apellido.text.replace(".","").strip() + " " + primer_nombre.text.replace(".","").strip() + " " + otros_nombres.text.replace(".","").strip()

                lista_verificacion.append(resultado)
                Lista_Error.append('ERROR: NO ES UN NIT VALIDO.')
                lista_validacion_DV.append("DV SIN VALIDACION")
                continue
            lista_verificacion.append(str(row['nid']))
            Lista_Error.append('ERROR: NO ES UN NIT VALIDO.')
            lista_validacion_DV.append("DV SIN VALIDACION")
    else:
        lista_verificacion.append(str(row['nid']))
        Lista_Error.append('NO ES NIT')
        lista_validacion_DV.append("DV SIN VALIDACION")

        

archivo.insert(8,'Consulta',lista_verificacion)
archivo.insert(9,'VALIDACION',Lista_Error)
archivo.insert(10,'VALIDACION DV', lista_validacion_DV)

archivo.to_excel(os.path.join(PATH_EXE,"Archivos\Validacion.xlsx"),engine='xlsxwriter',index=False)

# Cerramos el Driver
driver.close()
print('Finaliza la busqueda!!!')