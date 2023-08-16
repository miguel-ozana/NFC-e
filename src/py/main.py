# Imports the following functions from the selenium library
import threading
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By

# Imports the following functions from the selenium
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", {
  "download.default_directory": r"C:\Users\migue\downloads",
  "download.prompt_for_download": False,
  "download.directory_upgrade": True,
  "safebrowsing.enabled": True
})

# Driver quiting
def close_on_enter(driver):
  input("Pressione Enter para fechar o navegador...")
  driver.quit()

# Code
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
driver.get(r'C:\Users\migue\OneDrive\Documents\NFC-e\src\html\login.html')
  
# Login page
driver.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys('admin')
driver.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys('admin123')
time.sleep(0.5)
driver.find_element(By.XPATH, '/html/body/div/form/button').click()
  
# xlsx 
workbook = openpyxl.load_workbook(r'C:\Users\migue\OneDrive\Documents\NFC-e\xml\NotasEmitir.xlsx')
sheet = workbook.active
  
# Loops
for row in sheet.iter_rows(min_row=2, values_only=True):
  Cliente, cpf_cnpj, cep, Endereco, Bairro, Municipio, uf, Inscricao, Descricao, Quantidade, ValorUnit, ValorTotal  = row
    
  # Full forms
  driver.find_element(By.NAME, 'nome').send_keys(Cliente)
  driver.find_element(By.NAME, 'endereco').send_keys(Endereco)
  driver.find_element(By.XPATH, '//*[@id="formulario"]/input[3]').send_keys(Bairro)
  driver.find_element(By.XPATH, '//*[@id="formulario"]/input[4]').send_keys(Municipio)
  driver.find_element(By.XPATH, '//*[@id="formulario"]/input[5]').send_keys(cep)
  driver.find_element(By.XPATH, '//*[@id="formulario"]/select').send_keys(uf)
  driver.find_element(By.XPATH, '//*[@id="formulario"]/input[6]').send_keys(cpf_cnpj)
  driver.find_element(By.XPATH, '//*[@id="formulario"]/input[7]').send_keys(Inscricao)
  driver.find_element(By.XPATH, '//*[@id="formulario"]/input[8]').send_keys(Descricao)
  driver.find_element(By.XPATH, '//*[@id="formulario"]/input[9]').send_keys(Quantidade)
  driver.find_element(By.XPATH, '//*[@id="formulario"]/input[10]').send_keys(ValorUnit)
  driver.find_element(By.XPATH, '//*[@id="formulario"]/input[11]').send_keys(ValorTotal)
  
  
  #Click button
  driver.find_element(By.XPATH, '//*[@id="formulario"]/button').click()
  
  # webdriver closing thread
  close_thread = threading.Thread(target=close_on_enter, args=(driver,))
  close_thread.start()