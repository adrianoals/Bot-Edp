# Programa principal verificar status conta
from functions import *
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup

print('Iniciando o programa')

print('')

caminho_planilha = input('Informe o caminho da planilha de onde o robô extrairá a lista de condomínios: ')

print('')

aba_planilha = input('Informe o nome da aba de onde o robô extrairá a lista de condomínios: ')
lista_condominios = processa_planilha(caminho_planilha, aba_planilha)

print('')

print(f'Lista de condomínios que o robo irá acessar:\n\n{lista_condominios}')

print()

# Nome do arquivo Excel e caminho
nome_arquivo_excel = 'dados_condominios.xlsx'

# Tenta carregar um arquivo existente, ou cria um novo se não existir
try:
    df_excel = pd.read_excel(nome_arquivo_excel)
except FileNotFoundError:
    df_excel = pd.DataFrame(columns=['NOME', 'CNPJ', 'INSTALACAO', 'MÊS', 'VALOR', 'STATUS'])


# Iniciando o driver
driver = iniciar_driver()

# # Navegando Site
navegar_site(driver, site='https://www.edponline.com.br/imobiliaria')

# input('Digite algo para iniciar: ')

# # Pausa
pausa_aleatoria(3, 5)

# Ativando cookies
botao = driver.find_element(By.ID, 'onetrust-accept-btn-handler')
botao.click()

# Pausa
pausa_aleatoria(2, 4)

for item in lista_condominios:            
    # iniciado o site novamente
    navegar_site(driver, site='https://www.edponline.com.br/imobiliaria')    

    # Vendo o condominio que puxareos os dados
    print(f'Condomínio que iremos acessar:')
    print(item)

    # Armazenando dados do condomínio
    for condominio, dado in item.items():
        # Armazenar a chave em uma variável
        nome_condominio = condominio
        # Armazenar o primeiro item (item 1) em uma variável
        numero_cnpj = dado[0]
        # Armazenar o segundo item (item 2) em uma variável
        numero_instalacao = dado[1]

    # Fazendo Login
    try_login(driver)
    
    # Pausa
    pausa_aleatoria(5, 8)
    
    # Clique no link Minhas Contas
    clicar_minhas_contas(driver, nome_condominio, numero_instalacao)

    # Pausa
    pausa_aleatoria(5, 8)
    
    # Acessando a conta

    # Clique no segundo link Extrato e 2ª Via de Conta
    clicar_extrato_segunda_via(driver)

    # Pausa
    pausa_aleatoria(5, 8)
    
    # Clicando no CNPJ
    clicar_cnpj(driver)

    # Pausa
    pausa_aleatoria(5, 8)
    
    # Acessando uma instalação
    acessar_conta(driver, numero_cnpj, numero_instalacao)
    
    max_tentativas_acesso = 3
    tentativas_acesso = 0

    while tentativas_acesso < max_tentativas_acesso:
        # Verificando se o acesso foi bem-sucedido
        url_atual = driver.current_url
        if url_atual == 'https://www.edponline.com.br/servicos/extrato-de-contas':
            print('Sucesso ao acessar a instalação')
            break  # Saia do loop se o acesso for bem-sucedido
        else:
            print(f'Falha ao tentar entrar na instalação - Tentativa {tentativas_acesso + 1}')

            # Ações para tentar corrigir o acesso
            clicar_icone_conta(driver)
            pausa_aleatoria(5, 8)
            clicar_botao_sair(driver)
            try_login(driver)
            pausa_aleatoria(5, 8)
            clicar_minhas_contas(driver, nome_condominio, numero_instalacao)
            pausa_aleatoria(5, 8)
            clicar_extrato_segunda_via(driver)
            pausa_aleatoria(5, 8)
            clicar_cnpj(driver)
            pausa_aleatoria(5, 8)
            acessar_conta(driver, numero_cnpj, numero_instalacao)
            
            tentativas_acesso += 1

    # Se o loop atingir o número máximo de tentativas sem sucesso
    else:
        print(f'Não foi possível acessar a instalação após {max_tentativas_acesso} tentativas.')

        # Navegando página de serviços
        navegar_site(driver, site='https://www.edponline.com.br/servicos')

        # Encontrando e clicando no icone da conta
        clicar_icone_conta(driver)

        # Pausa
        pausa_aleatoria(5, 10)
    
        # Localizando o botão sair
        clicar_botao_sair(driver)

        # Pausa
        pausa_aleatoria(5, 10)
        
        print(f'Indo para o próximo condomínio não foi possível entrar na conta do ({nome_condominio} - {numero_instalacao})')

        print()

        continue

    # Pausa
    pausa_aleatoria(5, 10)

    # Se abrir a janela de fatura por email
    fechar_janela_fatura_por_email(driver)
    
    # Pausa
    pausa_aleatoria(5, 10)

    selecionar_fatura_meses(driver)
    
    # Pausa
    pausa_aleatoria(1, 2)
    
    clicar_ver_mais(driver)

    # Pausa
    pausa_aleatoria(1, 2)

    clicar_ver_mais(driver)

    # Pausa
    pausa_aleatoria(1, 2)

    clicar_ver_mais(driver)

    # Pausa
    pausa_aleatoria(1, 2)

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    cards = soup.find_all('div', class_='faturas__card')

    dados = []

    for card in cards:
        mes = card.find('p', class_='faturas__value').text
        valor = card.find('p', class_='faturas__value').find_next('p', class_='faturas__value').text
        status = card.find('div', class_='faturas__status')
        if status is not None:
            status = status.find('p').text
        else:
            status = card.find('div', class_='faturas__contaExtrato').find('p').text
        dados.append([mes, valor, status])

    # print(dados)

    print()

    # Cria um dicionário onde cada chave é um item da lista
    dicionario = {tuple(i): i for i in dados}

    # Converte o dicionário de volta para uma lista
    lista_sem_duplicados = list(dicionario.values())

    print(lista_sem_duplicados)

    # Convertendo a lista para um DataFrame do pandas
    df_novo = pd.DataFrame(lista_sem_duplicados, columns=['MÊS', 'VALOR', 'STATUS'])

    # Adicionando colunas do condomínio
    df_novo['NOME'] = nome_condominio
    df_novo['CNPJ'] = numero_cnpj
    df_novo['INSTALACAO'] = numero_instalacao

    # Concatenando o novo DataFrame ao DataFrame existente
    df_excel = pd.concat([df_excel, df_novo], ignore_index=True)

    print(f'Dados adicionados ao arquivo {nome_arquivo_excel}')
    print()

    # Salvando o DataFrame completo no arquivo Excel
    df_excel.to_excel(nome_arquivo_excel, index=False)


    # Navegando página de serviços
    navegar_site(driver, site='https://www.edponline.com.br/servicos')

    # Encontrando e clicando no icone da conta
    clicar_icone_conta(driver)

    # Pausa
    pausa_aleatoria(5, 6)
    
    # Localizando o botão sair
    clicar_botao_sair(driver)

    # Pausa
    pausa_aleatoria(5, 6)
    
    print()

input('Digite algo para fechar o programa: ')

driver.quit()