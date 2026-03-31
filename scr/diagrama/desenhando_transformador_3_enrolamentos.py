import math
import itertools
from diagrama.criacao_elementos import Trafo_3_enrolamentos
import schemdraw.elements as elm

def tratamento_trafo_3e(posicao_barra_p, posicao_barra_s, posicao_barra_t, posicao_transformador_3e, d):
    melhor_dist = float('inf')
    melhor_teta = 0
    melhor_flip = False
    melhor_reverse = False

    for teta, flip, reverse in itertools.product(range(360), [False, True], [False, True]):
        # Cria o trafo com as transformações apropriadas
        trafo = (
            Trafo_3_enrolamentos()
            .at(posicao_transformador_3e)
            .theta(teta)
        )
        if flip:
            trafo = trafo.flip()
        if reverse:
            trafo = trafo.reverse()

        # Desenho invisível
        trafo = trafo.color("#FFFFFF00").zorder(0)
        d += trafo

        # Calcula distância total entre os terminais e as barras
        ponto_p = trafo.absanchors['p']
        ponto_s = trafo.absanchors['s']
        ponto_t = trafo.absanchors['t']

        dist_atual = (
            math.dist(posicao_barra_s, ponto_s) +
            math.dist(posicao_barra_t, ponto_t)
        )

        # Atualiza se for melhor
        if dist_atual < melhor_dist:
            melhor_dist = dist_atual
            melhor_teta = teta
            melhor_flip = flip
            melhor_reverse = reverse

    return melhor_teta, melhor_flip, melhor_reverse


def Desenho_Transformador_3_enrolamentos(d, Transformador_3E, posicao_elementos, dicionario_cores):
        
        for index, row in Transformador_3E.iterrows():
            barra_p = row["Barra P"]
            barra_s = row["Barra S"]
            barra_t = row["Barra T"]

            tensao_p = row["Tensão Primário (kV)"]
            tensao_s = row["Tensão Secundário (kV)"]
            tensao_t = row["Tensão Terciário (kV)"]

            posicao_barra_p = posicao_elementos[barra_p]
            posicao_barra_s = posicao_elementos[barra_s]
            posicao_barra_t = posicao_elementos[barra_t]
            posicao_transformador_3e = posicao_elementos[row["Nome"]]
            potencia = row["Potência Nominal (MVA)"]
            conexao = row['Tipo de Conexão']
            nome = row["Nome"]
            
            melhor_teta, melhor_flip, melhor_reverse = tratamento_trafo_3e(posicao_barra_p, posicao_barra_s, posicao_barra_t, posicao_transformador_3e, d)
            trafo = Trafo_3_enrolamentos(cor_primario=dicionario_cores[tensao_p], cor_secundario=dicionario_cores[tensao_s], cor_terciario=dicionario_cores[tensao_t], conexao=conexao).at(posicao_transformador_3e).label(f"{nome}\n{tensao_p}-{tensao_s}-{tensao_t} kV\n{potencia} MVA").theta(melhor_teta)
            if melhor_flip:
                trafo = trafo.flip()

            if melhor_reverse:
                trafo = trafo.reverse()

            d += trafo

            # Capturar as coordenadas absolutas dos pontos de ancoragem
            ponto_p = trafo.absanchors['p']
            ponto_s = trafo.absanchors['s']
            ponto_t = trafo.absanchors['t']

            elm.Line().at(posicao_barra_p).to(ponto_p).color(dicionario_cores[tensao_p]).fill()
            elm.Line().at(posicao_barra_s).to(ponto_s).color(dicionario_cores[tensao_s]).fill()
            elm.Line().at(posicao_barra_t).to(ponto_t).color(dicionario_cores[tensao_t]).fill()