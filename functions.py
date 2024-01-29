from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import random
import pandas as pd
from pathlib import os
from dotenv import load_dotenv

load_dotenv()

def processa_planilha(caminho_planilha, aba_planilha):
    # Abra a planilha do Excel e a aba específica
    df = pd.read_excel(caminho_planilha, sheet_name=aba_planilha)

    # Crie uma lista para armazenar os dicionários
    lista_de_dicionarios = []

    # Itere pelas linhas do DataFrame
    for index, row in df.iterrows():
        # Crie um dicionário para cada linha
        dicionario = {
            row['NOME']: [row['CNPJ'], row['INSTALACAO']]
        }
        # Adicione o dicionário à lista
        lista_de_dicionarios.append(dicionario)
    
    return lista_de_dicionarios


def iniciar_driver():
    chrome_options = Options()
    arguments = ['--lang=pt-BR', '--window-size=1024,780,']
    for argument in arguments:
        chrome_options.add_argument(argument)

    # chrome_options.add_experimental_option('prefs', {
    #     'download.prompt_for_download': False,
    #     'profile.default_content_setting_values.notifications': 2,
    #     'profile.default_content_setting_values.automatic_downloads': 1,
    #     'download.default_directory': caminho_downloads,
    # })

    # driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

    # Insira o JavaScript
    # driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    return driver


def pausa_aleatoria(minimo, maximo):
    # Gera um número decimal aleatório entre minimo e maximo
    tempo_espera = random.uniform(minimo, maximo)
    # Pausa a execução do programa pelo tempo especificado
    time.sleep(tempo_espera)


# Navegando até o site
def navegar_site(driver, site):
    driver.get(site)

def login(driver):
    elemento = driver.find_element(By.XPATH, '//label[text()="São Paulo"]')
    elemento.click()
    campo_email = driver.find_element(By.ID, 'Email')
    campo_senha = driver.find_element(By.ID, 'Senha')
    
    # Limpe os campos antes de inserir os dados
    campo_email.clear()
    campo_senha.clear()
    
    # campo_email.send_keys('contato@villanovacondominios.com.br')
    # campo_senha.send_keys('Villanova1@') 
    campo_email.send_keys(str(os.getenv('EMAIL')))
    campo_senha.send_keys(str(os.getenv('SENHA')))
    
    # Localize o botão de login e clique nele
    botao_login = driver.find_element(By.ID, 'acessar')
    botao_login.click()

def try_login(driver, max_attempts=3):
    attempt = 0
    while attempt < max_attempts:
        try:
            pausa_aleatoria(3, 5)
            login(driver)
            pausa_aleatoria(4, 5)
            # Verificação de sucesso do login
            url_atual = driver.current_url
            if url_atual == 'https://www.edponline.com.br/servicos':
                print('Login efetuado')
                break
        except Exception as e:
            print(f"Falha na tentativa de login {attempt+1}: {str(e)}")
            attempt += 1
            # Se estorou a tentativa vai aparecer o print abaixo
            if attempt == max_attempts:
                print("Todas as tentativas de login falharam.")
                break
            # Informando as tentativas
            print("Tentando novamente...")
            time.sleep(3)


def clicar_minhas_contas(driver, nome_condominio, numero_instalacao):
    try:
        link = driver.find_element(By.XPATH, '//a[contains(text(), "Minhas Contas")]')
        link.click()
        return print('Clique em Minhas Contas bem-sucedido')
    except:
        print(f'Falha ao clicar em Minhas Contas: {nome_condominio} - {numero_instalacao}')
        return print('Falha ao clicar em Minhas Contas')

def clicar_extrato_segunda_via(driver):
    try:
        link2 = driver.find_element(By.PARTIAL_LINK_TEXT, 'Extrato e 2ª Via de Conta')
        link2.click()
        return print('Clique em Extrato e 2ª Via de Conta bem-sucedido')
    except:
        print('Falha ao clicar em Extrato e 2ª Via de Conta')
        return

def clicar_cnpj(driver):
    try:
        element = driver.find_element(By.CSS_SELECTOR, 'label[for="Documento-CNPJ"]')
        element.click()
        return print('Clique em CNPJ bem-sucedido')
    except:
        print('Falha ao clicar no CNPJ')
        return 

def acessar_conta(driver, numero_cnpj, numero_instalacao):
    try:
        campo_cnpj = driver.find_element(By.ID, 'Cnpj')
        campo_instalacao = driver.find_element(By.ID, 'Instalacao')
        campo_cnpj.send_keys(numero_cnpj)
        campo_instalacao.send_keys(numero_instalacao) 

        # Localize o botão de acessar e clique nele
        botao_acessar = driver.find_element(By.ID, 'acessar')
        botao_acessar.click()

        return True  # Indica que o preenchimento e clique foram bem-sucedidos
    except:
        print('Falha ao tentar acessar a conta')
        return False  # Indica que o preenchimento ou clique falhou

def fechar_janela_fatura_por_email(driver):
        fatura_por_email = driver.find_elements(By.CSS_SELECTOR, ".edp-btn.edp-btn--primary.edp-btn--outline.pull-right")
        if fatura_por_email:
            fatura_por_email[0].click()
            print('A janela de Fatura por e-mail foi fechada')
        else:
            print("A janela de Fatura por e-mail não estava ativa.")

def clicar_ver_fatura(driver):
    try:
        element = driver.find_element(By.CSS_SELECTOR, 'i.fa.fa-file-text-o')
        element.click()
        print('Botão "Ver Fatura" clicado')
        return True  # Indica que o clique foi bem-sucedido
    except:
        print('Falha ao carregar o botão "Ver Fatura"')
        print('Tentando novamente')
        try:
            element = driver.find_element(By.CSS_SELECTOR, 'i.fa.fa-file-text-o')
            element.click()
            print('Botão "Ver Fatura" clicado na segunda tentativa')
            return True  # Indica que o clique na segunda tentativa foi bem-sucedido
        except:
            print('Falha novamente')
            return False  # Indica que o clique falhou nas duas tentativas

def clicar_botao_baixar(driver):
    try:
        fin_element = driver.find_element(By.CSS_SELECTOR, ".icon-big.d-block.pe-7s-cloud-download")
        fin_element.click()
        # Aguarde até que o download esteja concluído
        print('Clique no botão baixar')
        return True  # Indica que houve clique no botão baixar
    except: 
        print('Falha ao clicar no botão Baixar')

def esperar_arquivo_baixado(driver, caminho_pasta_downloads, tempo_limite=10):
    try:
        # Aguarde até que o arquivo seja baixado (defina o tempo limite conforme necessário)
        wait = WebDriverWait(driver, tempo_limite)
        arquivo_baixado = wait.until(EC.presence_of_file((By.XPATH, f"//a[contains(@href,'{caminho_pasta_downloads}')]")))

        if arquivo_baixado:
            print('Arquivo baixado com sucesso')
            return True
        else:
            print('Arquivo não foi baixado')
            return False

    except Exception as e:
        print(f'Erro ao esperar o arquivo ser baixado: {e}')
        return False

def clicar_icone_conta(driver):
    try:
        link_icone_conta = driver.find_element(By.CSS_SELECTOR, ".icon-edp-usuario-pin")
        link_icone_conta.click()
        print('Ícone da conta clicado')
        return True  # Indica que o clique foi bem-sucedido
    except:
        print('Falha ao tentar clicar no ícone da conta. Tentando novamente...')
        try:
            # Tentando clicar no ícone da conta novamente
            link_icone_conta = driver.find_element(By.CSS_SELECTOR, ".icon-edp-usuario-pin")
            link_icone_conta.click()
            print('Ícone da conta clicado na segunda tentativa')
            return True  # Indica que o clique na segunda tentativa foi bem-sucedido
        except:
            print('Não foi possível clicar no ícone da conta')
            return False  # Indica que o clique falhou nas duas tentativas

def clicar_botao_sair(driver):
    try:
        link_sair = driver.find_element(By.CSS_SELECTOR, 'a.btn-link[href="/acesso/sair"]')
        link_sair.click()
        print('Botão "Sair" clicado')
        return True  # Indica que o clique foi bem-sucedido
    except:
        print('Erro ao tentar clicar no botão "Sair". Tentando novamente...')
        try:
            # Tentando clicar no botão "Sair" novamente
            link_sair = driver.find_element(By.CSS_SELECTOR, 'a.btn-link[href="/acesso/sair"]')
            link_sair.click()
            print('Botão "Sair" clicado na segunda tentativa')
            return True  # Indica que o clique na segunda tentativa foi bem-sucedido
        except:
            print('Não foi possível clicar no botão "Sair"')
            return False  # Indica que o clique falhou nas duas tentativas

def selecionar_fatura_meses(driver):
    # Encontre o elemento select pelo seu ID
    select_element = Select(driver.find_element(By.ID, 'selectVisualizarFaturas'))
    # Selecione a opção pelo valor
    select_element.select_by_value('12')
    print('Clique para ver 12 meses')

def clicar_ver_mais(driver):
    # Encontre o elemento <p> pela sua classe
    element = driver.find_element(By.CLASS_NAME, 'faturas__showMore')
    # Clique no elemento
    element.click()
    print('Clique no botão ver mais')
