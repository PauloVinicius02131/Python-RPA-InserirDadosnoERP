# Bibliotecas
import selenium
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver

from datetime import date
from datetime import timedelta
import time

import os
import sys

import pandas as pd

# Parâmetros Tabela Excel Input de Dados.
caminho = os.path.abspath(__file__ + "/../../")
caminho = caminho + "\entrada"
nomearquivo = "\dadosEntrada.xlsx"
caminho = caminho + nomearquivo
tabela = "Plan1"
cabecalho = 0
print(caminho)

df = pd.read_excel(caminho, sheet_name=tabela, header=cabecalho)

# Determinar coluna lida e tratar o DataFrame.
df = df['robo'].str.split(';', expand=True)
df.columns = ['Codigo', 'Qnt']

# Visualizar DataFrame de Input.
print(df.head())
print(type(df))

dfSaida = {}

# Encontrar caminho do chrome driver
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

# Variavel do driver
path = application_path + '\chromedriver'
driver = webdriver.Chrome(path)
driver.maximize_window()
# driver.get("http://transnet.grupocsc.com.br/sgtweb/")
driver.get("http://homolog.vicosa.transoft.com.br/sgtweb/")


def fLogin():
    try:
        # Digitar Login.
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "edtLogin"))).send_keys("autobot")
        # Digitar Senha.
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "edtSenha"))).send_keys("grupo2@22")
        # Clicar Login.
        driver.find_element(
            By.XPATH, "/html/body/div/div/div[1]/div[2]/div/div/form/div[4]/div/input").click()
    except:
        print('Erro no Login')
        input('Pressione Enter para continuar tentar novamente...')
        fLogin()


def fNavegacaoCompras():
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, 'Módulos'))).click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, 'Compras'))).click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, 'Processo padrão de compras'))).click()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, 'Solicitação de Compras'))).click()


def fAbrirSolicitacao():
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, 'idSolicitacaoCompra'))).clear()

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, 'idSolicitacaoCompra'))).send_keys('61980')

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, 'Pesquisar'))).click()

    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/section/div[5]/div[3]/form/center/div/center/table/tbody/tr[2]/td/div/table/tbody/tr/td[1]/a'))).click()


def fPreencherItensDaCotacao():
    for index, linha in df.iterrows():
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'input_idItem'))).click()

            time.sleep(1)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'input_idItem'))).clear()

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'input_idItem'))).send_keys(linha['Codigo'])

            time.sleep(1)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'qtInicialItem'))).click()

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'qtInicialItem'))).send_keys(Keys.CONTROL, 'a')

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'qtInicialItem'))).send_keys(Keys.BACKSPACE)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'qtInicialItem'))).send_keys(linha['Qnt'])

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'btnRelacionar'))).click()

            try:
                WebDriverWait(driver, 3).until(EC.alert_is_present())
                alerta = driver.switch_to.alert
                alerta.accept()
                print("Alerta aceitado")
            except:
                print('Não encontrado o  alerta')

            time.sleep(1)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'Gravar'))).click()

            try:
                WebDriverWait(driver, 5).until(EC.alert_is_present())
                alerta = driver.switch_to.alert
                alerta.accept()
                print("Alerta aceitado")
            except:
                print('Não encontrado o  alerta')

            dfSaida.append(linha['Codigo'])

        except:
            print('erro no loop')
