#----------------------------------------------------------------------------------------------------------------------
#------------------------------------------------ BIBLIOTECAS ---------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------

import schemdraw.elements as elm  # Possui os elementos elétricos já prontos
from datetime import datetime
import seaborn as sns # Geração de cores automáticas
import networkx as nx # Organizar barramentos
import xlwings as xw
import pandas as pd  # Trabalhar com tabelas e planilhas
import schemdraw  # Realizar os desenhos elétricos
import itertools
import math # cálculos matemáticos
import ctypes
import os


from functions.leitura_e_tratamento_tabelas import tabelas_de_dados
from diagrama.legenda_tecnica import adicionar_margem_pdf
from diagrama.criacao_elementos import *  # noqa: F403
from diagrama.organizacao_grafo import Grafo

def gerar_diagrama(caminho_excel="xlfaults.xlsm"):

    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------------------- LEITURA DE DADOS DO EXCEL  --------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    Maquina, Carga, Transformador, Linha, Barra, Transformador3E, Configuracoes = tabelas_de_dados(caminho_excel)

    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------- MONTAGEM DO GRAFO PARA ORGANIZAÇÃO DOS ELEMENTOS ----------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    G, posicao_elementos = Grafo(Linha, Transformador, Maquina, Carga, Barra)


    # Essa função tem como objetivo gerar cores para diferentes níveis de tensão
    def Gera_Cores_Tensoes(niveis_de_tensoes):  # lista de Tensões
        cores_tensoes = sns.color_palette("tab10", len(niveis_de_tensoes))  # Gera as cores
        dicionario_cores = dict(
            zip(niveis_de_tensoes, cores_tensoes))  # Cria um dicionário com a tensão como chave e a cor como valor
        return dicionario_cores


    # Essa função tem como objetivo, corrigir a posição central do transformador
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


    # Tratamento de rotação e flipagem dotrasnformador
    # def tratamento_trafo_3e(posicao_barra_p, posicao_barra_s, posicao_barra_t, posicao_transformador_3e, d):
    #     melhor_dist = float('inf')
    #     melhor_teta = 0
    #     melhor_flip = False
    #     melhor_reverse = False

    #     for teta, flip, reverse in itertools.product(range(360), [False, True], [False, True]):
    #         # Cria o trafo com as transformações apropriadas
    #         trafo = (
    #             Trafo_3_enrolamentos()
    #             .at(posicao_transformador_3e)
    #             .theta(teta)
    #         )
    #         if flip:
    #             trafo = trafo.flip()
    #         if reverse:
    #             trafo = trafo.reverse()

    #         # Desenho invisível
    #         trafo = trafo.color("#FFFFFF00").zorder(0)
    #         d += trafo

    #         # Calcula distância total entre os terminais e as barras
    #         ponto_p = trafo.absanchors['p']
    #         ponto_s = trafo.absanchors['s']
    #         ponto_t = trafo.absanchors['t']

    #         dist_atual = (
    #             math.dist(posicao_barra_s, ponto_s) +
    #             math.dist(posicao_barra_t, ponto_t)
    #         )

    #         # Atualiza se for melhor
    #         if dist_atual < melhor_dist:
    #             melhor_dist = dist_atual
    #             melhor_teta = teta
    #             melhor_flip = flip
    #             melhor_reverse = reverse

    #     return melhor_teta, melhor_flip, melhor_reverse


    #----------------------------------------------------------------------------------------------------------------------
    #------------------------------------------------- VARIAVEIS ------------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    nx.draw(G, pos=posicao_elementos, with_labels=True, node_color='skyblue', edge_color='black', node_size=1000, font_size=12)

    dicionario_cores = Gera_Cores_Tensoes(Barra["Tensão (kV)"].tolist())
    caminho_diagrama = fr"{caminho_excel.replace("xlsm", "pdf")}"
    wb = xw.Book(fr"{caminho_excel}")
    sheet = wb.sheets['Oculto']
    nome_caso = sheet.range('A2').value
    potencia_base = sheet.range('A1').value
    tipo_de_curto = "qualquer"
    simulador = "Xlfaults"
    data = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    tensao_falta = complex(1, 0)


    #----------------------------------------------------------------------------------------------------------------------
    #-------------------------------------- INÍCIO DO DIAGRAMA ------------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------


    with schemdraw.Drawing(file=fr"{caminho_diagrama}", dpi=150, theme='grade3', show=False) as d:
        d.config(fontsize=10)


    # Desenhando o Elemento BARRA ****************************************************************************************
        for index, row in Barra.iterrows():
            barra = row["Número"]
            tensao = row["Tensão (kV)"]
            posicao = posicao_elementos[barra]
            d += Desenho_Barra().at(posicao).label(f"Barra {barra}\n{tensao} kV").color(dicionario_cores[tensao])


    # Desenhando o Elemento Maquina **************************************************************************************
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
            maquina = Desenho_Maquina(cor=dicionario_cores[tensao], type=type).at(posicao_gerador).theta(teta+180).label(f"{nome}\n{tensao} kV\n{conexao}\n{potencia} MVA")
            d += maquina
            ponto_c = maquina.absanchors['C']
            d += elm.Line().at(posicao_barra).to(ponto_c).color(dicionario_cores[tensao])

            
    # Desenhando o Elemento Transformador ********************************************************************************
        for index, row in Transformador.iterrows():
            nome = row["Nome"]
            posicao_barra_de = posicao_elementos[row["Barra de"]]
            posicao_barra_para = posicao_elementos[row["Barra para"]]
            posicao_transformador = posicao_elementos[row["Nome"]]
            conexao = row['Tipo de Conexão']
            potencia = row["Potência Nominal (MVA)"]
            tensao_primario = row["Tensão Primário (kV)"]
            tensao_secundario = row["Tensão Secundário (kV)"]

            teta = math.degrees(math.atan2(posicao_barra_para[1] - posicao_barra_de[1], posicao_barra_para[0] - posicao_barra_de[0]))
            xm, ym = Ajustes_trafos(posicao_barra_de, posicao_barra_para, posicao_transformador)
            trafo = Desenho_Transformador(cor_primario=dicionario_cores[tensao_primario], cor_secundario=dicionario_cores[tensao_secundario], conexao=conexao).at(posicao_transformador).theta(teta).label(f"{nome}\n{tensao_primario}-{tensao_secundario} kV\n{potencia} MVA")
            d += trafo
            ponto_p = trafo.absanchors['p']
            ponto_s = trafo.absanchors['s']
            d += elm.Line().at(ponto_s).to(posicao_barra_para).color(dicionario_cores[tensao_secundario])
            d += elm.Line().at(posicao_barra_de).to(ponto_p).color(dicionario_cores[tensao_primario])


    # Desenhando o Elemento Linha ********************************************************************************
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


    # Desenhando o Elemento Carga **************************************************************************************
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


    # Desenhando o Elemento Transformador de 3 Enrolamentos ***************************************************************
        
        # for index, row in Transformador_3E.iterrows():
        #     barra_p = row["Barra P"]
        #     barra_s = row["Barra S"]
        #     barra_t = row["Barra T"]

        #     tensao_p = row["Tensão Primário (kV)"]
        #     tensao_s = row["Tensão Secundário (kV)"]
        #     tensao_t = row["Tensão Terciário (kV)"]

        #     posicao_barra_p = posicao_elementos[barra_p]
        #     posicao_barra_s = posicao_elementos[barra_s]
        #     posicao_barra_t = posicao_elementos[barra_t]
        #     posicao_transformador_3e = posicao_elementos[row["Nome"]]
        #     potencia = row["Potência Nominal (MVA)"]
        #     conexao = row['Tipo de Conexão']
        #     nome = row["Nome"]
            
        #     melhor_teta, melhor_flip, melhor_reverse = tratamento_trafo_3e(posicao_barra_p, posicao_barra_s, posicao_barra_t, posicao_transformador_3e, d)
        #     trafo = Trafo_3_enrolamentos(cor_primario=dicionario_cores[tensao_p], cor_secundario=dicionario_cores[tensao_s], cor_terciario=dicionario_cores[tensao_t], conexao=conexao).at(posicao_transformador_3e).label(f"{nome}\n{tensao_p}-{tensao_s}-{tensao_t} kV\n{potencia} MVA").theta(melhor_teta)
        #     if melhor_flip:
        #         trafo = trafo.flip()

        #     if melhor_reverse:
        #         trafo = trafo.reverse()

        #     d += trafo

        #     # Capturar as coordenadas absolutas dos pontos de ancoragem
        #     ponto_p = trafo.absanchors['p']
        #     ponto_s = trafo.absanchors['s']
        #     ponto_t = trafo.absanchors['t']

        #     elm.Line().at(posicao_barra_p).to(ponto_p).color(dicionario_cores[tensao_p]).fill()
        #     elm.Line().at(posicao_barra_s).to(ponto_s).color(dicionario_cores[tensao_s]).fill()
        #     elm.Line().at(posicao_barra_t).to(ponto_t).color(dicionario_cores[tensao_t]).fill()


    # Desenhando o Elemento Curto-Circuito ***************************************************************
        #Curto().at(posicao_elementos[barra_curto])


    adicionar_margem_pdf(caminho_diagrama, caminho_diagrama.replace(".pdf", "_Final.pdf"), 200, nome_caso, potencia_base, simulador, data, dicionario_cores)
    ctypes.windll.user32.MessageBoxW(0, f"Diagrama gerado com sucesso!\n{caminho_diagrama.replace(".pdf","")}_Legendado.pdf", "Sucesso", 0)
    if os.path.exists(caminho_diagrama):
        os.remove(caminho_diagrama)
        os.remove(caminho_diagrama.replace(".pdf", "_Final.pdf"))

    caminho_final = caminho_diagrama.replace(".pdf", "_Legendado.pdf")

    if os.path.exists(caminho_final):
        os.startfile(caminho_final)

gerar_diagrama(caminho_excel="xlfaults.xlsm")