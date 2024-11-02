from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from bs4 import BeautifulSoup
import re
import time
from pathlib import Path
from time import sleep
import requests
import subprocess
import pyperclip
import pyautogui
import sys

# criando um caminho para o Webdrive do Google Chrome
PASTA_CHROME_DRIVE = 'C:\\Users\\lucas\\Desktop\\PROJETOS PYTHON\\Captura_de_emails_transportadoras\\cap\\chromedriver.exe'

# entrada de dados de usuário
o_cidade = input('CIDADE ORIGEM: ')
o_estado = input('ESTADO ORIGEM (SIGLA): ')
d_cidade = input('CIDADE DESTINO: ')
d_estado = input('ESTADO DESTINO (SIGLA): ')


# Há um pequeno bug, que ao manter o mouse na caixa de sugestão de cidade, ele selecionara a cidade errada. solução encontrada:
print('*** ATENÇÃO O PROGRAMA ABRIRA O NAVEGADOR, COLOQUE O MOUSE TODO PARA CIMA DA TELA! ***')
sleep(6)

# definindo uma função para utilizar o Chrome drive pelo Selenium
def make_chrome_browser(*options: str) -> webdriver.Chrome:
    chrome_options = webdriver.ChromeOptions()
    if options is not None:
        for option in options:
            chrome_options.add_argument(option)
    chrome_service = Service(
        executable_path=str(PASTA_CHROME_DRIVE),
    )

    browser = webdriver.Chrome(
        service=chrome_service,
        options=chrome_options
    )
    return browser



options = ()
browser = make_chrome_browser(*options)


# acessando o site Transvias, escolhido para este projeto
browser.get('https://www.transvias.com.br/')

# preenchendo os dados de cidades e estados da origem e destino com webdriver
origem = WebDriverWait(browser, 2).until(
    EC.presence_of_element_located(
        (By.NAME, 'OrigemDescricao')
    )
)


origem.send_keys(o_cidade + o_estado)

sleep(1)

origem.send_keys(Keys.TAB)

sleep(2)

destino = WebDriverWait(browser, 2).until(
    EC.presence_of_element_located(
        (By.NAME, 'DestinoDescricao')
    )
)

sleep(2)

destino.send_keys(d_cidade + d_estado)

sleep(2)
destino.send_keys(Keys.TAB)


# Usando o comando com WebDriver para clicar em pesquisar
pesquisar = WebDriverWait(browser, 2).until(
    EC.presence_of_element_located(
        (By.CLASS_NAME, 'm-button--lg')
    )
)

sleep(3)

pesquisar.click()


sleep(3)
elemento = WebDriverWait(browser, 5).until(
    EC.presence_of_element_located(
        (By.TAG_NAME, 'body')
    )
)


# selecionar e copiar a tela inteira

elemento.send_keys(Keys.CONTROL + 'a')

sleep(2)

browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")

sleep(2)

elemento.send_keys(Keys.CONTROL, 'c')


sleep(1)

browser.quit()

sleep(2)

texto_copiado = pyperclip.paste()


# usando expressões regulares para filtrar emails
padrao_email = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'



emails = re.findall(padrao_email, texto_copiado)

mail_valid = []

#ignorar alguns e-mails caso necessarios
for mail in emails:
    if not mail == 'exemplo@exemplo.com':
        mail_valid.append(mail)

emails_str =  ', '.join(mail_valid)



pyperclip.copy(emails_str)

sleep(2)

# usando subprocess para abrir o e-mail (Outlook) no windows
subprocess.Popen(['C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.exe'])

sleep(2)


# comandos usando o pyautogui para simular o teclado

pyautogui.hotkey('ctrl', 'n')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('tab')
sleep(2)

pyautogui.hotkey('tab')

assunto = 'COTAÇÃO DE FRETE'

pyperclip.copy(assunto)

sleep(1)

pyautogui.hotkey('ctrl', 'v')

sleep(2)

pyautogui.hotkey('tab')

sleep(2)

dados_transp = '''Olá, tudo bem?
Gostaria de cotar um frete por favor.

DESTINATÁRIO
CLIENTE DESTINATÁRIO: 

CNPJ: 

CIDADE: 

UF: 

Cep: 

FRETE: FOB

PRODUTO: 
VALOR DA NOTA: R$ 
 VOLUMES
Peso kg


LOCAL DE COLETA (REMETENTE)
EMPRESA: 
ENDEREÇO: 
BAIRRO: 
CEP: 
CIDADE: 
UF: 
CNPJ: 
TELEFONE: 
'''


pyperclip.copy(dados_transp)


pyautogui.hotkey('ctrl', 'v')

