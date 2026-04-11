import subprocess
import sys

def instalar_requisitos(caminho_requirements):
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "-r", caminho_requirements
        ])
        print("Bibliotecas instaladas com sucesso!")
    except subprocess.CalledProcessError:
        print("Erro ao instalar bibliotecas.")

caminho = sys.argv[1]
instalar_requisitos(caminho_requirements = caminho)