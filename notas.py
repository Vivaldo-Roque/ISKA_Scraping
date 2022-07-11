# Necessário para usar selenium
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# Minha função
from utils import ConvertTableToExcel
import os

url = "http://41.218.115.14/" # Endereço IP da Secretaria virtual
aproveitamento = "Discentes/Secretaria/Aproveitamento.aspx" # Caminho para aproveitamento
classification = "Discentes/Secretaria/Classificacoes.aspx?s=B3738228BDA70CB1" # Caminho para classificações

table_xpath = "//div[@class='table-responsive-sm']//table" # Xpath para a tabela

# IDs HTML dos elementos de login

username_input_ID = "ContentPlaceHolder1_TxtUtilizador"
password_input_ID = "ContentPlaceHolder1_TxtSenha"
login_button_ID = "ContentPlaceHolder1_BtnEntrar"
error_ID = "ContentPlaceHolder1_lblErro"

delay = 3 # Atraso

# Essa função tem a missão de pegar as notas de aproveitamento e classificação
def PegarMinhasNotas(username, password):

    os.system("cls")
    os.system("color 0A")

    # Configurações
    option = webdriver.ChromeOptions() # Usar o Google Chrome
    option.headless = True # Abrir navegador no background
    option.add_argument('--headless')
    option.add_argument('--log-level=3')

    driver = webdriver.Chrome(options=option) # Inicializar nosso navegador

    print("[*] Iniciado!")

    driver.get(url+"login.aspx") # Navegar para página de login

    # Pegar os elementos de login na página web
    driver.find_element(By.ID, username_input_ID).send_keys(username)
    driver.find_element(By.ID, password_input_ID).send_keys(password)
    driver.find_element(By.ID, login_button_ID).click()

    # Aguardar o estado pronto ser concluído
    WebDriverWait(driver, delay).until(
        lambda driver: driver.execute_script("return document.readyState === 'complete'")
    )

    # Mensagem de erro
    error_message = "Não foi possível autenticar, por favor, tente novamente!"

    # Obter os erros (se houver)
    errors = driver.find_elements(By.ID, error_ID)

    # Se encontrarmos essa mensagem de erro nos erros, o login falhou
    if any(error_message in e.text for e in errors):
        print("[!] Login falhou")
        driver.quit() # Sair do navegador
    else:
        print("[+] Login com sucesso")

        driver.get(url+aproveitamento) # Navegar para página de aproveitamento

        # Localizar todas a tabelas
        elements = WebDriverWait(driver, delay).until(EC.presence_of_all_elements_located((By.XPATH, table_xpath)))

        # Converter tabelas para excel
        ConvertTableToExcel(filename="aproveitamentos", elements= elements)

        driver.get(url+classification) # Navegar para página de classificações

        # Localizar todas a tabelas
        elements = WebDriverWait(driver, delay).until(EC.presence_of_all_elements_located((By.XPATH, table_xpath)))

        # Converter tabelas para excel
        ConvertTableToExcel(filename="classificações", elements= elements)

        driver.quit() # Sair do navegador
        print("[+] Tudo feito")