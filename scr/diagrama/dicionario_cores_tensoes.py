import seaborn as sns

# Essa função tem como objetivo gerar cores para diferentes níveis de tensão
def Gera_Cores_Tensoes(niveis_de_tensoes):  # lista de Tensões
    cores_tensoes = sns.color_palette("tab10", len(niveis_de_tensoes))  # Gera as cores
    dicionario_cores = dict(
        zip(niveis_de_tensoes, cores_tensoes))  # Cria um dicionário com a tensão como chave e a cor como valor
    return dicionario_cores