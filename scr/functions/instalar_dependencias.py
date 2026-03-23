import subprocess
import sys

def instalar_requisitos():
    try:
        subprocess.check_call([
            sys.executable, "-m", "pip", "install", "-r", "requirements.txt"
        ])
    except Exception as e:
        print("Erro ao instalar dependências:", e)