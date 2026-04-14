import subprocess
import sys

from exportar.mensagem import abrir_loading, fechar_loading

def instalar_requisitos(caminho_requirements):
    abrir_loading("Instalando Dependências")
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "-r", caminho_requirements
        ])
        print("Bibliotecas instaladas com sucesso!")
    except subprocess.CalledProcessError:
        print("Erro ao instalar bibliotecas.")
    fechar_loading()

caminho = sys.argv[1]
instalar_requisitos(caminho_requirements = caminho)