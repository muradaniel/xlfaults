import math
import schemdraw.elements as elm
from diagrama.criacao_elementos import Simbolo_Transformador
from conversor_anafas_xlfaults import procv_tensoes



def Ajustes_trafos(posicao_barra_de, posicao_barra_para, posicao_transformador):
        x1, y1 = posicao_barra_de
        x2, y2 = posicao_barra_para
        xm, ym = posicao_transformador
        if (x1 < x2) and (y1 < y2):
            xm = xm - 1.5
            ym = ym - 1.5
        elif (x1 > x2) and (y1 > y2):
            xm = xm + 1.5
            ym = ym + 1.5
        elif (x1 < x2) and (y1 > y2):
            xm = xm - 1.5
            ym = ym + 1.5
        else:
            xm = xm + 1.5
            ym = ym - 1.5
        return xm, ym



def Desenho_Transformador(d, Transformador, posicao_elementos, dicionario_cores, Barra):
    for index, row in Transformador.iterrows():
        nome = row["Nome"]
        posicao_barra_de = posicao_elementos[row["Barra de"]]
        posicao_barra_para = posicao_elementos[row["Barra para"]]
        posicao_transformador = posicao_elementos[row["Nome"]]
        conexao = row['Tipo de Conexão']
        potencia = row["Potência Nominal (MVA)"]
        tensao_primario = row["Tensão Primário (kV)"]
        tensao_secundario = row["Tensão Secundário (kV)"]

        barra_de = row["Barra de"]
        barra_para = row["Barra para"]

        tensao_barra_de = Barra.set_index('Número').loc[barra_de, 'Tensão (kV)']
        
        tensao_barra_para = Barra.set_index('Número').loc[barra_para, 'Tensão (kV)']

        cor_primario = dicionario_cores[tensao_barra_de]
        cor_secundario = dicionario_cores[tensao_barra_para]

        teta = math.degrees(math.atan2(posicao_barra_para[1] - posicao_barra_de[1], posicao_barra_para[0] - posicao_barra_de[0]))
        xm, ym = Ajustes_trafos(posicao_barra_de, posicao_barra_para, posicao_transformador)
        # trafo = Simbolo_Transformador(cor_primario=dicionario_cores[tensao_primario], cor_secundario=dicionario_cores[tensao_secundario], conexao=conexao).at(posicao_transformador).theta(teta).label(f"{nome}\n{tensao_primario}-{tensao_secundario} kV\n{potencia} MVA")
        trafo = Simbolo_Transformador(cor_primario=cor_primario, cor_secundario=cor_secundario, conexao=conexao).at(posicao_transformador).theta(teta).label(f"{nome}\n{tensao_primario}-{tensao_secundario} kV\n{potencia} MVA")
        d += trafo
        ponto_p = trafo.absanchors['p']
        ponto_s = trafo.absanchors['s']
        d += elm.Line().at(ponto_s).to(posicao_barra_para).color(dicionario_cores[tensao_barra_para])
        d += elm.Line().at(posicao_barra_de).to(ponto_p).color(dicionario_cores[tensao_barra_de])