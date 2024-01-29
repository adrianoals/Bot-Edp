import pandas as pd

# Abra a planilha do Excel e a aba específica
df = pd.read_excel(r"C:\Users\drili\OneDrive\Área de Trabalho\Bot Edp\planilha\EDP_FÁTIMA.xlsx", sheet_name='VENCIMENTO(10)_1')

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


print(len(lista_de_dicionarios))

print(lista_de_dicionarios)
print(type(lista_de_dicionarios))