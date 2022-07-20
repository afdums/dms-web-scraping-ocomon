#Download do ocomon
# https://ocomonphp.sourceforge.io/downloads/


#Instalação da lib do Selenium
#pip3 install selenium
#Instalação BeautifulSoup
#pip3 install beautifulsoup4
#Instalação lxml
#pip3 install lxml
#Download do Geckodriver para Firefox ou ChromeDriver para Chrome
#https://github.com/mozilla/geckodriver/releases
#https://chromedriver.chromium.org/downloads
#Salvar dentro da pasta do python descompactado

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox import options #busca dos elementos na tela
from selenium.webdriver.firefox.options import Options #opções para o navegador
from time import sleep #dar uma pausa entre os comandos
from bs4 import BeautifulSoup #tratar o HTML retornado
import pandas as pd #manipular arquivos, exportar para CSV

options = Options()
options.headless = False #True #executar de forma oculta

navegador = webdriver.Firefox(options=options)

link = "http://localhost:8000/ocomon-4.0RC1/login.php"

navegador.get(url=link)
sleep(1)

inputUsuario = navegador.find_element(By.ID,value="user")
inputUsuario.send_keys("admin")
sleep(1)

inputSenha = navegador.find_element(By.ID,value="pass")
inputSenha.send_keys("admin")
sleep(1)

buttonLogin = navegador.find_element(By.ID,value="bt_login")
buttonLogin.click()
sleep(3) #bom para deixar a página carregar

linkOcorrencias = navegador.find_element(By.ID,value="OCOMON")
linkOcorrencias.click()
sleep(3)

linkFiltroAvancado = navegador.find_element(By.LINK_TEXT,value="Fitro avançado")
linkFiltroAvancado.click()
sleep(3)

#aqui precisamos entrar no frame dos filtros
frameFiltros = navegador.find_element(By.ID,value="iframeMain")
navegador.switch_to.frame(frameFiltros)

sleep(3)

selectMesCorrente = navegador.find_element(By.ID,value="current_month")
selectMesCorrente.click()
sleep(1)

inputDataAbertura = navegador.find_element(By.ID,value="data_abertura_from")
inputDataAbertura.clear() #neste caso nao precisa, mas as vezes o campo tem value já ai o send_keys concatena
inputDataAbertura.send_keys("01/01/1900")
sleep(1)

buttonSearch = navegador.find_element(By.ID,value="idSearch")
buttonSearch.click()
sleep(1)

buttonGerenciarColunas = navegador.find_element(By.XPATH,value="//button/span[contains(text(), 'Gerenciar colunas')]")
buttonGerenciarColunas.click()
sleep(1)

colAbertoPor = navegador.find_element(By.XPATH, value="//button/span[contains(text(), 'Aberto Por')]")
colAbertoPor.click()
sleep(1)

tabelaTicket = navegador.find_element(By.ID,value="table_tickets_queue")

htmlContent = tabelaTicket.get_attribute('outerHTML')

soup = BeautifulSoup(htmlContent,"html.parser")

tickets = soup.find(name="table")

df = pd.read_html(str(tickets))[0]

df.to_csv("tickets.csv", encoding='UTF-8', sep=';', index=False)


navegador.switch_to.default_content() 

navegador.quit()

