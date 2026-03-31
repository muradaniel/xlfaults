import schemdraw.elements as elm
import math

def Desenho_Linha(d, Linha, posicao_elementos, dicionario_cores):
    for index, row in Linha.iterrows():
        nome = row["Nome"]
        posicao_barra_de = posicao_elementos[row["Barra de"]]
        posicao_barra_para = posicao_elementos[row["Barra para"]]
        posicao_linha = posicao_elementos[row["Nome"]]
        tensao = row["Tensão (kV)"]
        teta = math.degrees(math.atan2(posicao_barra_para[1] - posicao_barra_de[1], posicao_barra_para[0] - posicao_barra_de[0]))
        d+= elm.Line().at(posicao_barra_de).to(posicao_linha).color(dicionario_cores[tensao])
        d += elm.Dot(radius=0.001).color(dicionario_cores[tensao]).label(f"{nome}", loc='bottom', ofst=-0.85)
        d+= elm.Line().at(posicao_linha).to(posicao_barra_para).color(dicionario_cores[tensao])