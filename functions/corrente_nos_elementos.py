import networkx as nx
import pandas as pd
import numpy as np
import cmath
import math


def corrente_nos_elementos(Configuracoes, Linha, Maquina, Carga, Transformador, resultados, G, T012abc):
    for index_conf, row_conf in Configuracoes.iterrows():
        linhas_corrigidas = [] # Acumulador do DataFrame desse curto circuito
        for tabela_elemento in [Linha, Maquina, Carga]:  # Transformador possui algumas propriedades a mais
            for index_elemento, row_elemento in tabela_elemento.iterrows():
                tensoes = resultados['Tensões nas Barras - Não Corrigidas'][index_conf]
                nome_elemento = row_elemento["Nome"]
                barra_curto = row_conf["Barra de Ocorrência"]



                if tabela_elemento is Linha: # Calculo das correntes nas linhas de transmissão
                    numero_barra_de = row_elemento["Barra de"]
                    numero_barra_para = row_elemento["Barra para"]
                    
                    delta_vk0 = (
                        tensoes.loc[tensoes["Barra"] == numero_barra_de, "Seq. (0)"].values[0]
                        - tensoes.loc[tensoes["Barra"] == numero_barra_para, "Seq. (0)"].values[0]
                    )
                    delta_vk1 = (
                        tensoes.loc[tensoes["Barra"] == numero_barra_de, "Seq. (+)"].values[0]
                        - tensoes.loc[tensoes["Barra"] == numero_barra_para, "Seq. (+)"].values[0]
                    )
                    delta_vk2 = (
                        tensoes.loc[tensoes["Barra"] == numero_barra_de, "Seq. (-)"].values[0]
                        - tensoes.loc[tensoes["Barra"] == numero_barra_para, "Seq. (-)"].values[0]
                    )
                    correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=numero_barra_de, weight='weight') # o target pode ser barra para, tanto faz
                    correcao_negativa = nx.shortest_path_length(G, source=numero_barra_de, target=barra_curto, weight='weight')



                elif tabela_elemento is Maquina:  # Calculo das correntes nas máquinas
                    barra_conectada = row_elemento["Barra Conectada"]
                    conexao = row_elemento["Tipo de Conexão"]
                    if conexao == "Yt":
                        delta_vk0 = (0 - tensoes.loc[tensoes["Barra"] == barra_conectada, "Seq. (0)"].values[0])
                    else: # Demais conexões, como Y e D
                        delta_vk0 = 0

                    delta_vk1 = (1 - tensoes.loc[tensoes["Barra"] == barra_conectada, "Seq. (+)"].values[0])
                    delta_vk2 = (0 - tensoes.loc[tensoes["Barra"] == barra_conectada, "Seq. (-)"].values[0])

                    correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=barra_conectada, weight='weight')
                    correcao_negativa = nx.shortest_path_length(G, source=barra_conectada, target=barra_curto, weight='weight')



                else: # Carga Shunt
                    barra_conectada = row_elemento["Barra Conectada"]
                    conexao = row_elemento["Tipo de Conexão"]

                    if conexao == "Yt":
                        delta_vk0 = (0 - tensoes.loc[tensoes["Barra"] == barra_conectada, "Seq. (0)"].values[0])
                    else: # Demais conexões, como Y e D
                        delta_vk0 = 0
                    delta_vk1 = (tensoes.loc[tensoes["Barra"] == barra_conectada, "Seq. (+)"].values[0])
                    delta_vk2 = (tensoes.loc[tensoes["Barra"] == barra_conectada, "Seq. (-)"].values[0])

                    correcao_positiva = nx.shortest_path_length(G, source=barra_curto, target=barra_conectada, weight='weight')
                    correcao_negativa = nx.shortest_path_length(G, source=barra_conectada, target=barra_curto, weight='weight')

                Ic0 = delta_vk0 / row_elemento["Z0 (pu) Base"] # A impedância de seq(0) já contém a impedância de neutro, caso exista
                Ic1 = delta_vk1 / row_elemento["Z1 (pu) Base"]
                Ic2 = delta_vk2 / row_elemento["Z1 (pu) Base"]

                Ic0 = Ic0
                Ic1 = Ic1 * cmath.rect(1, math.radians(correcao_positiva))
                Ic2 = Ic2 * cmath.rect(1, math.radians(correcao_negativa))

                Ic3f = T012abc @ np.array([[Ic0], [Ic1], [Ic2]])

                Ica = Ic3f[0][0]
                Icb = Ic3f[1][0]
                Icc = Ic3f[2][0]

                linhas_corrigidas.append({
                    "Nome": nome_elemento,
                    "Fase A": Ica,
                    "Fase B": Icb,
                    "Fase C": Icc,
                    "Seq. (0)": Ic0,
                    "Seq. (+)": Ic1,
                    "Seq. (-)": Ic2
                })

        tabela_corrente_elementos = pd.DataFrame(linhas_corrigidas)
        resultados['Correntes de Contibuição'].append(tabela_corrente_elementos)
    return resultados     