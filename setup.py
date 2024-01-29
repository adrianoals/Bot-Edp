import sys
import os
from cx_Freeze import setup, Executable

# Definir o que deve ser incluído na pasta final
arquivos = ['requirements.txt', 'instruções.txt', 'functions.py']

# Saída dos arquivos
configuracao = Executable(
    script = 'main.py',
    icon = 'account_icon.ico'
)

# Configurar o CX-Freeze(detalhes do programa)
setup(
    name ='Bot EDP',
    version = '1.0',
    description = 'Esse programa baixa as contas no site da EDP',
    author = 'Adriano',
    options = {'build_exe': {'include_files': arquivos}},
    executables = [configuracao] 
)

# python setup.py build