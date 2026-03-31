import networkx as nx
import pandas as pd
import numpy as np
import cmath
import math


def tensoes_barras(Barra, Zbarra12, Zbarra0, resultados, Configuracoes, G, T012abc):
    for index, row in Configuracoes.iterrows(): # Vai iterar a cada curto circuito
        linhas_nao_corrigidas = []  # acumulador do DataFrame desse curto
        linhas_corrigidas = []  # acumulador do DataFrame desse curto
        for barra in Barra['Número'].tolist():
            Vka0 = 0 - Zbarra0.loc[row["Barra de Ocorrência"], barra] * resultados['Correntes de Falta'][index]["Seq. (0)"].iloc[0]
            Vka1 = 1 - Zbarra12.loc[row["Barra de Ocorrência"], barra] * resultados['Correntes de Falta'][index]["Seq. (+)"].iloc[0]
            Vka2 = 0 - Zbarra12.loc[row["Barra de Ocorrência"], barra] * resultados['Correntes de Falta'][index]["Seq. (-)"].iloc[0]
            
            linhas_nao_corrigidas.append({
                "Barra": barra,
                "Seq. (0)": Vka0,
                "Seq. (+)": Vka1,
                "Seq. (-)": Vka2
            })

            correcao_positiva = nx.shortest_path_length(G, source=row["Barra de Ocorrência"], target=barra, weight='weight') # +30
            correcao_negativa = nx.shortest_path_length(G, source=barra, target=row["Barra de Ocorrência"], weight='weight') # -30

            Vka0 = Vka0
            Vka1 = Vka1 * cmath.rect(1, math.radians(correcao_positiva))
            Vka2 = Vka2 * cmath.rect(1, math.radians(correcao_negativa))

            Vk3f = T012abc @ np.array([
                [Vka0],
                [Vka1],
                [Vka2]
            ])

            Vka = Vk3f[0][0]
            Vkb = Vk3f[1][0]
            Vkc = Vk3f[2][0]

            linhas_corrigidas.append({
                "Barra": barra,
                "Fase A": Vka,
                "Fase B": Vkb,
                "Fase C": Vkc,
                "Seq. (0)": Vka0,
                "Seq. (+)": Vka1,
                "Seq. (-)": Vka2
            })

        tabela_tensoes_barras_corrigidas = pd.DataFrame(linhas_corrigidas)
        tabela_tensoes_barras_nao_corrigidas = pd.DataFrame(linhas_nao_corrigidas)
    
        resultados['Tensões nas Barras'].append(tabela_tensoes_barras_corrigidas)
        resultados['Tensões nas Barras - Não Corrigidas'].append(tabela_tensoes_barras_nao_corrigidas)
    print("Tensões nas barras calculadas com sucesso!")
    return resultados
 
