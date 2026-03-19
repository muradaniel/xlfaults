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



                elif tabela_elemento is Carga: # Carga Shunt
                    barra_conectada = row_elemento["Barra Conectada"]
                    conexao = row_elemento["Tipo de  Conexão"]

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
        for index_elemento, row_elemento in Transformador.iterrows():
            tensoes = resultados['Tensões nas Barras - Não Corrigidas'][index_conf]
            nome_elemento = row_elemento["Nome"]
            barra_curto = row_conf["Barra de Ocorrência"]
            numero_barra_de = row_elemento["Barra de"]
            numero_barra_para = row_elemento["Barra para"]
            conexao = row_elemento["Tipo de Conexão"]

            if conexao == "Yt-Yt":
                delta_vk0 = (
                    tensoes.loc[tensoes["Barra"] == numero_barra_de, "Seq. (0)"].values[0]
                    - tensoes.loc[tensoes["Barra"] == numero_barra_para, "Seq. (0)"].values[0]
                )
            else: # Demais conexões...
                delta_vk0 = 0
            delta_vk1 = (
                tensoes.loc[tensoes["Barra"] == numero_barra_de, "Seq. (+)"].values[0]
                - tensoes.loc[tensoes["Barra"] == numero_barra_para, "Seq. (+)"].values[0]
            )
            delta_vk2 = (
                tensoes.loc[tensoes["Barra"] == numero_barra_de, "Seq. (-)"].values[0]
                - tensoes.loc[tensoes["Barra"] == numero_barra_para, "Seq. (-)"].values[0]
            )
            correcao_positiva_primario = nx.shortest_path_length(G, source=barra_curto, target=numero_barra_de, weight='weight')
            correcao_negativa_primario = nx.shortest_path_length(G, source=numero_barra_de, target=barra_curto, weight='weight')
            correcao_positiva_secundario = nx.shortest_path_length(G, source=barra_curto, target=numero_barra_para, weight='weight')
            correcao_negativa_secundario = nx.shortest_path_length(G, source=numero_barra_para, target=barra_curto, weight='weight')

            Ic0 = delta_vk0 / row_elemento["Z0 (pu) Base"]
            Ic1 = delta_vk1 / row_elemento["Z1 (pu) Base"] 
            Ic2 = delta_vk2 / row_elemento["Z1 (pu) Base"]

            Ic0 = Ic0
            Ic1_p = Ic1 * cmath.rect(1, math.radians(correcao_positiva_primario))
            Ic2_p = Ic2 * cmath.rect(1, math.radians(correcao_negativa_primario))

            Ic0 = Ic0
            Ic1_s = Ic1 * cmath.rect(1, math.radians(correcao_positiva_secundario))
            Ic2_s = Ic2 * cmath.rect(1, math.radians(correcao_negativa_secundario))

            Ic3f_primario = T012abc @ np.array([[Ic0], [Ic1_p], [Ic2_p]])
            Ic3f_secundario = T012abc @ np.array([[Ic0], [Ic1_s], [Ic2_s]])

            Ica_primario = Ic3f_primario[0][0]
            Icb_primario = Ic3f_primario[1][0]
            Icc_primario = Ic3f_primario[2][0]

            Ica_secundario = Ic3f_secundario[0][0]
            Icb_secundario = Ic3f_secundario[1][0]
            Icc_secundario = Ic3f_secundario[2][0]

            linhas_corrigidas.append({
                "Nome": nome_elemento + " (Primário)",
                "Fase A": Ica_primario,
                "Fase B": Icb_primario,
                "Fase C": Icc_primario,
                "Seq. (0)": Ic0,
                "Seq. (+)": Ic1_p,
                "Seq. (-)": Ic2_p
            })

            linhas_corrigidas.append({
                "Nome": nome_elemento + " (Secundário)",
                "Fase A": Ica_secundario,
                "Fase B": Icb_secundario,
                "Fase C": Icc_secundario,
                "Seq. (0)": Ic0,
                "Seq. (+)": Ic1_s,
                "Seq. (-)": Ic2_s
            })

        tabela_corrente_elementos = pd.DataFrame(linhas_corrigidas)
        resultados['Correntes de Contibuição'].append(tabela_corrente_elementos)
    return resultados     