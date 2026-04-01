# OBS1: Aqui, terá um comando para instalar todas as dependências necessárias para rodar o projeto, utilizando o arquivo requirements.txt
# OBS2: Será avaliado se vale mais a pena criar um executavel ou deixar os arquivos python
# OBS3: Função ainda não sendo utilizada...

import subprocess
import sys

def instalar_requisitos():
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "-r", "requirements.txt"
        ])
    except Exception as e:
        print("Erro ao instalar dependências:", e)