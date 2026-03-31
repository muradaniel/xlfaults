from diagrama.criacao_elementos import Simbolo_Barra

def Desenho_Barra(d, Barra, posicao_elementos, dicionario_cores):
    for index, row in Barra.iterrows():
        barra = row["Número"]
        tensao = row["Tensão (kV)"]
        posicao = posicao_elementos[barra]
        d += Simbolo_Barra().at(posicao).label(f"Barra {barra}\n{tensao} kV").color(dicionario_cores[tensao])