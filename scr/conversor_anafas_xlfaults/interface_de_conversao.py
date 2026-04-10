from tkinter import Tk
from tkinter.filedialog import askopenfilename

def selecionar_arquivo():
    root = Tk()
    root.withdraw()  # esconde a janela principal

    caminho = askopenfilename(
        title="Selecione um arquivo ANAFAS",
        filetypes=[("Arquivos ANAFAS", "*.ANA*")]
    )

    return caminho