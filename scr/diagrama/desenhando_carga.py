import math
import schemdraw.elements as elm

def Desenho_Carga(d, Carga, posicao_elementos, dicionario_cores):
    for index, row in Carga.iterrows():
        barra = row["Barra Conectada"]
        tensao = row["Tensão (kV)"]
        nome = row["Nome"]
        posicao_carga = posicao_elementos[nome]
        posicao_barra = posicao_elementos[barra]
        conexao = row['Tipo de Conexão']
        potencia = f"{row['Potência Ativa (%)']} + J{row['Potência Reativa (%)']} MVA"
        teta = math.degrees(math.atan2(posicao_elementos[nome][1] - posicao_elementos[barra][1], posicao_elementos[nome][0] - posicao_elementos[barra][0]))
        d += elm.Line().at(posicao_barra).to(posicao_carga).color(dicionario_cores[tensao])
        d += elm.GroundSignal().at(posicao_carga).theta(teta+90).label(f"{nome}\n{potencia}\n{conexao}", loc='bottom').color(dicionario_cores[tensao]).fill()