# Baixando contas "programa principal"
from functions import *
from selenium.webdriver.common.by import By


print('Iniciando o programa')

print('')

caminho_planilha = input('Informe o caminho da planilha de onde o robô extrairá a lista de condomínios: ')

print('')

aba_planilha = input('Informe o nome da aba de onde o robô extrairá a lista de condomínios: ')
lista_condominios = processa_planilha(caminho_planilha, aba_planilha)

print('')

print(f'Lista de condomínios que o robo irá acessar:\n\n{lista_condominios}')

print()

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

    # Clique no ver fatura
    clicar_ver_fatura(driver)

    # Pausa
    pausa_aleatoria(10, 15) 
    
    # Clicando no baixar
    clicar_botao_baixar(driver)

    # Pausa 
    pausa_aleatoria(10, 15) 


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