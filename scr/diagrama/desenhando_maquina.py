import math
import schemdraw.elements as elm
from diagrama.criacao_elementos import Simbolo_Maquina

def Desenho_Maquina(d, Maquina, posicao_elementos, dicionario_cores):
    for index, row in Maquina.iterrows():
        barra = row["Barra Conectada"]
        tensao = row["Tensão (kV)"]
        potencia = row["Potência Nominal (MVA)"]
        nome = row["Nome"]
        posicao_gerador = posicao_elementos[nome]
        posicao_barra = posicao_elementos[barra]
        conexao = row['Tipo de Conexão']
        
        teta = math.degrees(math.atan2(posicao_elementos[nome][1] - posicao_elementos[barra][1], posicao_elementos[nome][0] - posicao_elementos[barra][0]))
        
        if row["Tipo de Máquina"] == "Gerador":
            type = "G"
        else:
            type = "M"
        maquina = Simbolo_Maquina(cor=dicionario_cores[tensao], type=type).at(posicao_gerador).theta(teta+180).label(f"{nome}\n{tensao} kV\n{conexao}\n{potencia} MVA")
        d += maquina
        ponto_c = maquina.absanchors['C']
        d += elm.Line().at(posicao_barra).to(ponto_c).color(dicionario_cores[tensao])