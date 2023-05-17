import pandas as pd
from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec, wait
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
import re
import os
import time

# -----------------VARIABLES DE ESPERA---------------------------
time_load = 10
path_exe = os.getcwd()
time_search_cc = 5
# -----------------VARIABLES DE ESPERA---------------------------


# Configuraciones para descarga
download_dir = path_exe
chrome_options = webdriver.ChromeOptions()
preferences = {"download.default_directory" : download_dir ,
"directory_upgrade": True,
"safebrowsing.enable": True}
chrome_options.add_experimental_option("prefs",preferences)

# Cargamos el driver
driver = webdriver.Chrome(chrome_options=chrome_options,executable_path=os.path.join(path_exe,"Driver\chromedriver.exe"))

# Mostramos la pagina
driver.maximize_window()
url = "https://antecedentes.policia.gov.co:7005/WebJudicial/index.xhtml"
driver.get(url)
time.sleep(time_load)

# TyC = driver.find_element_by_xpath('//*[@id="aceptaOption:0"]')
# TyC.click()
BtnTyC = driver.find_element_by_xpath('//*[@id="continuarBtn"]')
BtnTyC.send_keys(Keys.ENTER)

# Configuramos la espera del DIV
wait = WebDriverWait(driver,time_load * 3)
wait.until(ec.visibility_of_element_located((By.XPATH,'//*[@id="j_idt7_content"]')))

txt_cc = driver.find_element_by_xpath('//*[@id="cedulaInput"]')
txt_cc.clear()
txt_cc.send_keys('1020467166')

captcha = driver.find_element_by_xpath('//*[@id="recaptcha-anchor"]/div[1]')
captcha.send_keys(Keys.ESCAPE)
time.sleep(3)

btnBusquedad = driver.find_element_by_xpath('//*[@id="j_idt17"]')
btnBusquedad.send_keys(keys.ENTER)

print('hola')