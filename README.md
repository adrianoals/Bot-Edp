# Projeto de Automação da EDP

## Descrição
Este projeto consiste em um conjunto de scripts Python desenvolvidos para automatizar a tarefa de acessar o site da EDP, 
realizar login e baixar as contas de luz de uma lista de condomínios fornecida em uma planilha do Excel. Além disso, 
o projeto inclui um script que, ao invés de baixar as contas, grava as informações do status de conta dos últimos 12 meses 
de cada instalação de luz. Este sistema foi desenvolvido especificamente para uma administradora de condomínios, visando facilitar 
as operações diárias da empresa.

## Funcionalidades
- **Baixar Contas de Luz:** Automatiza o download das contas de luz para uma lista de condomínios a partir de uma planilha do Excel.
- **Registro de Status de Conta:** Grava informações em uma planilha excel sobre o status de conta dos últimos 12 meses para cada instalação.

## Instruções de Uso
1. Informe o caminho relativo da planilha que irá extrair a lista de condomínios. Exemplo: `C:\\Users\\usuario\\Downloads`.
2. Informe em qual aba da planilha irá extrair as informações.
3. Aguarde o programa rodar.

**Observações:**
- Ao digitar o caminho, não coloque as aspas `" "`.
- A planilha deve manter uma configuração padrão para que o programa possa ler e entender a lista de condomínios.

## Dependências
O projeto requer as bibliotecas Python, que estão listadas no arquivo `requirements.txt`

