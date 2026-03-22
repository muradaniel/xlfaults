# OBS1: Abaixo, foi utilizado o método das submatrizes de admitâncias
# OBS2: As linhas que o somatório complex(0, 0) podemos apagar, deixei mais a carater ilustrativo
# OBS3: Será Necessário realizar um tratamento para o caso de barras isoladas, para evitar erros na inversão da matriz Ybarra0

import pandas as pd
import numpy as np
from functions.regularizar_Ybarra0 import regularizar_Ybarra0

def calcular_matrizes(Maquina, Transformador, Linha, Carga, Barra):

    # -------------------------- CÁLCULOS DA YBARRA (+) (-) -------------------------------
    Ybarra12 = pd.DataFrame(np.zeros((len(Barra['Número']), len(Barra['Número'])), dtype=complex), index=Barra['Número'].tolist(), columns=Barra['Número'].tolist())

    for index, row in Maquina.iterrows():
        Ybarra12.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] =  (1 / row["Z1 (pu) Base"]) + Ybarra12.loc[int(row["Barra Conectada"]), int(row["Barra Conectada"])]
    for index, row in Carga.iterrows():
        Ybarra12.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] =  (1 / row["Z1 (pu) Base"]) + Ybarra12.loc[int(row["Barra Conectada"]), int(row["Barra Conectada"])]
    for index, row in Linha.iterrows():
        Ybarra12.at[int(row["Barra de"]), int(row["Barra de"])] =  (1 / row["Z1 (pu) Base"]) + Ybarra12.loc[int(row["Barra de"]), int(row["Barra de"])]
        Ybarra12.at[int(row["Barra para"]), int(row["Barra para"])] = (1 / row["Z1 (pu) Base"]) + Ybarra12.loc[int(row["Barra para"]), int(row["Barra para"])]
        Ybarra12.at[int(row["Barra de"]), int(row["Barra para"])] = (-1 / row["Z1 (pu) Base"]) + Ybarra12.loc[int(row["Barra de"]), int(row["Barra para"])]
        Ybarra12.at[int(row["Barra para"]), int(row["Barra de"])] =  (-1 / row["Z1 (pu) Base"]) + Ybarra12.loc[int(row["Barra para"]), int(row["Barra de"])] 
    for index, row in Transformador.iterrows():
        Ybarra12.at[int(row["Barra de"]), int(row["Barra de"])] = (1 / row["Z1 (pu) Base"]) + Ybarra12.loc[int(row["Barra de"]), row["Barra de"]]
        Ybarra12.at[int(row["Barra para"]), int(row["Barra para"])] = (1 / row["Z1 (pu) Base"]) + Ybarra12.loc[int(row["Barra para"]), row["Barra para"]]
        Ybarra12.at[int(row["Barra para"]), int(row["Barra de"])] = (-1 / row["Z1 (pu) Base"]) + Ybarra12.loc[int(row["Barra para"]), row["Barra de"]]
        Ybarra12.at[int(row["Barra de"]), int(row["Barra para"])] = (-1 / row["Z1 (pu) Base"]) + Ybarra12.loc[int(row["Barra de"]), row["Barra para"]]

    # ---------------------------- CÁLCULOS DA ZBARRA (+), (-) ----------------------------
    Zbarra12 = pd.DataFrame(np.linalg.inv(Ybarra12.values), index = Barra['Número'].tolist(), columns = Barra['Número'].tolist())

    # ---------------------------- CÁLCULOS DA YBARRA (ZERO) ------------------------------
    Ybarra0 = pd.DataFrame(np.zeros((len(Barra['Número']), len(Barra['Número'])), dtype=complex), index=Barra['Número'].tolist(), columns=Barra['Número'].tolist())
    for index, row in Maquina.iterrows():
        if row["Tipo de Conexão"] == "Yt":
            Ybarra0.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0.loc[row["Barra Conectada"], row["Barra Conectada"]]
        elif row["Tipo de Conexão"] == "Y":
            Ybarra0.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] = complex(0, 0) + Ybarra0.loc[row["Barra Conectada"], row["Barra Conectada"]]
        elif row["Tipo de Conexão"] == "D":
            Ybarra0.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] = complex(0, 0) + Ybarra0.loc[row["Barra Conectada"], row["Barra Conectada"]]
    for index, row in Carga.iterrows():
        if row["Tipo de Conexão"] == "Yt":
            Ybarra0.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0.loc[row["Barra Conectada"], row["Barra Conectada"]]
        else:
            Ybarra0.at[int(row["Barra Conectada"]), int(row["Barra Conectada"])] = complex(0, 0) + Ybarra0.loc[row["Barra Conectada"], row["Barra Conectada"]]
    for index, row in Linha.iterrows():
        if True: # Apenas para deixar o código formatado
            Ybarra0.at[int(row["Barra de"]), int(row["Barra de"])] =  (1 / row["Z0 (pu) Base"]) + Ybarra0.loc[int(row["Barra de"]), row["Barra de"]]
            Ybarra0.at[int(row["Barra para"]), int(row["Barra para"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0.loc[int(row["Barra para"]), row["Barra para"]]
            Ybarra0.at[int(row["Barra de"]), int(row["Barra para"])] = (-1 / row["Z0 (pu) Base"]) + Ybarra0.loc[int(row["Barra de"]), row["Barra para"]]
            Ybarra0.at[int(row["Barra para"]), int(row["Barra de"])] =  (-1 / row["Z0 (pu) Base"]) + Ybarra0.loc[int(row["Barra para"]), row["Barra de"]]
    for index, row in Transformador.iterrows():
        if row["Tipo de Conexão"] == "Yt-Yt":
            Ybarra0.at[int(row["Barra de"]), int(row["Barra de"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0.loc[int(row["Barra de"]), row["Barra de"]]
            Ybarra0.at[int(row["Barra para"]), int(row["Barra para"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0.loc[int(row["Barra para"]), row["Barra para"]]
            Ybarra0.at[int(row["Barra de"]), int(row["Barra para"])] = (-1 / row["Z0 (pu) Base"]) + Ybarra0.loc[int(row["Barra de"]), row["Barra para"]]
            Ybarra0.at[int(row["Barra para"]), int(row["Barra de"])] = (-1 / row["Z0 (pu) Base"]) + Ybarra0.loc[int(row["Barra para"]), row["Barra de"]]
        elif (row["Tipo de Conexão"] == "Yt-D"):
            Ybarra0.at[int(row["Barra de"]), int(row["Barra de"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0.loc[int(row["Barra de"]), row["Barra de"]]
            Ybarra0.at[int(row["Barra para"]), int(row["Barra para"])] = complex(0, 0) + Ybarra0.loc[int(row["Barra para"]), row["Barra para"]]
            Ybarra0.at[int(row["Barra de"]), int(row["Barra para"])] = complex(0, 0) + Ybarra0.loc[int(row["Barra de"]), row["Barra para"]]
            Ybarra0.at[int(row["Barra para"]), int(row["Barra de"])] = complex(0, 0) + Ybarra0.loc[int(row["Barra para"]), row["Barra de"]]
        elif (row["Tipo de Conexão"] == "D-Yt"):
            Ybarra0.at[int(row["Barra de"]), int(row["Barra de"])] = complex(0, 0) + Ybarra0.loc[int(row["Barra de"]), row["Barra de"]]
            Ybarra0.at[int(row["Barra para"]), int(row["Barra para"])] = (1 / row["Z0 (pu) Base"]) + Ybarra0.loc[int(row["Barra para"]), row["Barra para"]]
            Ybarra0.at[int(row["Barra de"]), int(row["Barra para"])] = complex(0, 0) + Ybarra0.loc[int(row["Barra de"]), row["Barra para"]]
            Ybarra0.at[int(row["Barra para"]), int(row["Barra de"])] = complex(0, 0) + Ybarra0.loc[int(row["Barra para"]), row["Barra de"]]
        else:
            Ybarra0.at[int(row["Barra de"]), int(row["Barra de"])] = complex(0, 0) + Ybarra0.loc[int(row["Barra de"]), row["Barra de"]]
            Ybarra0.at[int(row["Barra para"]), int(row["Barra para"])] = complex(0, 0) + Ybarra0.loc[int(row["Barra para"]), row["Barra para"]]
            Ybarra0.at[int(row["Barra de"]), int(row["Barra para"])] = complex(0, 0) + Ybarra0.loc[int(row["Barra de"]), row["Barra para"]]
            Ybarra0.at[int(row["Barra para"]), int(row["Barra de"])] = complex(0, 0) + Ybarra0.loc[int(row["Barra para"]), row["Barra de"]]

    # ---------------------------- CÁLCULOS DA ZBARRA (+), (-) ----------------------------


    Zbarra0 = regularizar_Ybarra0(Ybarra0, Barra) # Regulariza a matriz Ybarra0 e obtém Zbarra0, tratando o caso de barras isoladas

    return Ybarra12, Zbarra12, Ybarra0, Zbarra0#, isoladas