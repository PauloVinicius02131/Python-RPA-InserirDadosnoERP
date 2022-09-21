# Bibliotecas
from unittest import case
from xmlrpc.client import Boolean
import selenium
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium import webdriver

from datetime import date, datetime
from datetime import timedelta
import time

import os
import sys

import pandas as pd

print('💾 Iniciando Programa ... \n ⏱: ' +
      str(datetime.now().strftime("%m/%d %H:%M:%S")))

print('\n 📚 Carregando Tabela de Dados ... \n')

# Parâmetros Tabela Excel para Input de Dados.
caminho = os.path.abspath(__file__ + "/../../")
caminho = caminho + "\entrada"
nomearquivo = "\dadosEntrada.xlsx"
caminho = caminho + nomearquivo

# DataFrame Input de Dados.
tabela = "Dados"
cabecalho = 0
df = pd.read_excel(caminho, sheet_name=tabela, header=cabecalho)
df = df['robo'].str.split(';', expand=True)
df.columns = ['Codigo', 'Qnt']


print(' ✅ Tabela Carregada, Amostra dos dados: \n', '\n', df.head())
# Parametros Cabecalhos para Input de Dados.
tabela = "Cabecalhos"
cabecalho = 0

# Dataframe Cabecalho para Input de Dados.
dic = pd.read_excel(caminho, sheet_name=tabela, header=cabecalho, dtype=str)
almoxarifado = str(dic['Almoxarifado'][0])
empresa = str(dic['Empresa'][0])

print('\n 🚩 A Empresa selecionada é: 📣 ' + empresa)
print('\n 🚩 O Almoxarifado selecionado é: 📣 ' + almoxarifado)

# Parametros de Ambiente para utilização no Transnet.
periodo = date.today().strftime("%d/%m/%Y")
usuario = 'AUTOBOT'

# Parametros de Ambiente para utilização nas Saidas.
dataSaidaCdProduto = []
dataSaidaQnt = []
dataSaidaAlertas = []

# Encontrar Chrome Driver na pasta do projeto.
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)

# Variaveis de configuração do driver.
options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
path = application_path + '\chromedriver'
driver = webdriver.Chrome(executable_path=path, chrome_options=options)
driver.maximize_window()

driver.get("http://transnet.grupocsc.com.br/sgtweb/")
# driver.get("http://homolog.vicosa.transoft.com.br/sgtweb/")


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
        print('Erro no Login - Estouro')
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


def fFiltrosSolicitacao():
    # Listar todos os campos de filtros.
    formulario = driver.find_element(By.NAME, 'formulario')

    # Pegar todos os inputs do formulário.
    formularioInputs = formulario.find_elements(By.TAG_NAME, 'input')

    # Limpar todos os campos e preencher campos necessários.
    for item in formularioInputs:
        driver.execute_script(
            "arguments[0].setAttribute('value',arguments[1])", item, '')
        formInput = (item.get_attribute('name'))
        if formInput == 'dtAbertura':
            driver.execute_script(
                "arguments[0].setAttribute('value',arguments[1])", item, periodo)
        elif formInput == 'dtFechamento':
            driver.execute_script(
                "arguments[0].setAttribute('value',arguments[1])", item, periodo)

    # Pegar todos os selects do formulário.
    formularioSelects = formulario.find_elements(By.TAG_NAME, 'select')

    # Limpar todos os campos e preencher campos necessários.
    for item in formularioSelects:
        switch = (item.get_attribute('name'))
        if switch == 'csTipo':
            Select(driver.find_element(By.NAME, 'csTipo')
                   ).select_by_visible_text('Normal')
        elif switch == 'csStatus':
            Select(driver.find_element(By.NAME, 'csStatus')
                   ).select_by_visible_text('Aberta')
        elif switch == 'idAlmoxarifado':
            Select(driver.find_element(By.NAME, 'idAlmoxarifado')).select_by_visible_text(
                almoxarifado)  # Fazer depara para o almoxarifado selecionado
        elif switch == 'idUsuario':
            Select(driver.find_element(By.NAME, 'idUsuario')
                   ).select_by_visible_text('AUTOBOT')

    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.NAME, 'Pesquisar'))).click()


def fAbrirSolicitacao():
    # Aguardar pesquisa terminar.
    WebDriverWait(driver, 10).until(
        EC.invisibility_of_element((By.ID, 'ajaxLoader')))

    WebDriverWait(driver, 5).until(EC.presence_of_all_elements_located(
        (By.ID, 'registrosvlistacadastrodesolicitacaodecompras')))

    elementoPesquisado = driver.find_element(
        By.ID, 'registrosvlistacadastrodesolicitacaodecompras').find_elements(By.TAG_NAME, 'tbody')[0].text

    if elementoPesquisado == 'Nenhum Registro Localizado':
        try:
            driver.find_element(By.NAME, '<u>I</u>nserir').click()

            formulario = driver.find_element(By.NAME, 'formulario')

            formularioSelects = formulario.find_elements(By.TAG_NAME, 'select')

            for item in formularioSelects:
                switch = (item.get_attribute('name'))
                if switch == 'idAlmoxarifado':
                    Select(driver.find_element(By.NAME, 'idAlmoxarifado')
                           ).select_by_visible_text(almoxarifado)
                elif switch == 'idEmpresa':
                    Select(driver.find_element(By.NAME, 'idEmpresa')
                           ).select_by_visible_text(empresa)
                elif switch == 'csTipo':
                    Select(driver.find_element(By.NAME, 'csTipo')
                           ).select_by_visible_text('Normal')

            WebDriverWait(driver, 10).until(
                EC.invisibility_of_element((By.ID, 'ajaxLoader')))

            driver.find_element(By.NAME, '<u>I</u>nserir').click()

        except:
            print('Erro ao cadastrar uma nova solicitação de compra.')

    else:
        print(
            'Achei uma solicitação ja criada hoje em meu nome com status:\u001b[1m Aberta. \u001b[0m')
        tabela = driver.find_element(
            By.ID, 'registrosvlistacadastrodesolicitacaodecompras').find_element(By.TAG_NAME, 'tbody')
        linha = tabela.find_elements(By.TAG_NAME, 'tr')
        linha[0].find_elements(By.TAG_NAME, 'td')[0].click()


def fPreencherItensDaCotacao():
    time.sleep(1)
    for index, linha in df.iterrows():
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element((By.ID, 'ajaxLoader')))

        driver.find_element(By.NAME, 'input_idItem').clear()

        WebDriverWait(driver, 10).until(lambda driver: len(driver.execute_script(
            'return document.getElementById("input_idItem").value')) == 0)

        driver.find_element(By.NAME, 'input_idItem').send_keys(linha['Codigo'])

        # Esperar encontrar o produto.
        WebDriverWait(driver, 10).until(lambda driver: len(driver.execute_script(
            'return document.getElementById("input_idItem").value')) > 11)

        dataSaidaCdProduto.append(driver.execute_script(
            'return document.getElementById("input_idItem").value'))

        driver.find_element(By.NAME, 'qtInicialItem').click()
        driver.find_element(By.NAME, 'qtInicialItem').send_keys(
            Keys.CONTROL, 'a')
        driver.find_element(By.NAME, 'qtInicialItem').send_keys(Keys.BACKSPACE)
        driver.find_element(By.NAME, 'qtInicialItem').send_keys(linha['Qnt'])

        dataSaidaQnt.append(linha['Qnt'])

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.NAME, 'btnRelacionar'))).click()

        alertaSaida = []
        try:
            while WebDriverWait(driver, 4).until(EC.alert_is_present()):
                alerta = driver.switch_to.alert
                alerta_texto = alerta.text
                alertaSaida.append(alerta_texto)
                alerta.accept()
        except:
            pass

        dataSaidaAlertas.append(alertaSaida)

        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.NAME, 'Gravar'))).click()

        # Colocar while e salvar no append para verificar se o item foi inserido com sucesso.
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alerta = driver.switch_to.alert
        alerta.accept()

        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element((By.ID, 'ajaxLoader')))

    Saidas = {'Código_Produto': dataSaidaCdProduto,
              'Qnt': dataSaidaQnt, 'Alerta': dataSaidaAlertas}

    dfSaida = pd.DataFrame.from_dict(Saidas, orient='index').T
    caminhosaida = os.path.join(os.environ["HOMEPATH"], "Desktop")
    dfSaida.to_csv(caminhosaida + '\\' + 'Cotacao_Eventos_' + datetime.now().strftime(
        "%m_%d_%H_%M_%S") + '.csv', index=False, sep=';', encoding='utf-8')


def fRelatorioProdutosCotacao():
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, 'colunaslista')))

    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, 'lista')))

    # Construção do Cabeçalho do Relatório.
    listaCabecalho = ['Produto', 'Referencia', 'Quantidade',
                      'Saldo Atual', 'Valor Unitário', 'Valor Total']

    # Fazer cabeçalho da tabela de forma dinâmica.
    # listaSolicitacaoCabecalho = driver.find_elements(By.CLASS_NAME, 'colunaslista')
    # for item in listaSolicitacaoCabecalho:
    #     listaCabecalho.append(item.text)

    # Construção da Lista de Itens Solicitados.
    divTabelaSolicitacao = driver.find_element(By.CLASS_NAME, "lista")
    corpoTabelaSolicitacao = divTabelaSolicitacao.find_element(
        By.TAG_NAME, "tbody")
    linhasTabelaSolicitacao = corpoTabelaSolicitacao.find_elements(
        By.TAG_NAME, "tr")

    card = []

    for item in range(len(linhasTabelaSolicitacao)):
        produtoSolicitado = linhasTabelaSolicitacao[item].find_elements(By.TAG_NAME, "td")[
            0]
        referenciaSolicitada = linhasTabelaSolicitacao[item].find_elements(
            By.TAG_NAME, "td")[1]
        saldoAltualSolicitado = linhasTabelaSolicitacao[item].find_elements(
            By.TAG_NAME, "td")[6]
        valorUnitarioSolicitado = linhasTabelaSolicitacao[item].find_elements(
            By.TAG_NAME, "td")[7]
        valorTotalSolicitado = linhasTabelaSolicitacao[item].find_elements(
            By.TAG_NAME, "td")[8]
        quantidadeSolicitada = linhasTabelaSolicitacao[item].find_elements(
            By.TAG_NAME, "input")[1]
        teste = [produtoSolicitado.text, referenciaSolicitada.text, driver.execute_script(
            "return arguments[0].value", quantidadeSolicitada), saldoAltualSolicitado.text, valorUnitarioSolicitado.text, valorTotalSolicitado.text]
        card.append(teste)

    print(card)

    dfteste = pd.DataFrame(card, columns=listaCabecalho)
    caminhosaida = os.path.join(os.environ["HOMEPATH"], "Desktop")
    dfteste.to_csv(caminhosaida + '\\' + 'Cotacao_Solicitados_' + datetime.now().strftime(
        "%m_%d_%H_%M_%S") + '.csv', index=False, sep=';', encoding='utf-8')
