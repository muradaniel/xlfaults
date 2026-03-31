import numpy as np
import pandas as pd

def correntes_injetadas(Ybarra12, Ybarra0, resultados, Configuracoes, Barra, T012abc):
    
    if 'Correntes Injetadas nos Barramentos' not in resultados:
        resultados['Correntes Injetadas nos Barramentos'] = []

    for index_conf, row_config in Configuracoes.iterrows():
        
        tensoes = resultados['Tensões nas Barras - Não Corrigidas'][index_conf]

        # 🔥 Garantir alinhamento correto
        tensoes = tensoes.set_index("Barra").loc[Barra["Número"]]

        # Sequência
        I0injetada = Ybarra0 @ tensoes["Seq. (0)"].values
        I1injetada = Ybarra12 @ tensoes["Seq. (+)"].values
        I2injetada = Ybarra12 @ tensoes["Seq. (-)"].values

        # Vetor 012
        I012 = np.vstack([I0injetada, I1injetada, I2injetada])

        # Conversão para abc
        Iabc_injetada = T012abc @ I012

        # Fases
        Ica, Icb, Icc = Iabc_injetada

        # DataFrame
        df = pd.DataFrame({
            "Barra": Barra["Número"].values,
            "Fase A": Ica,
            "Fase B": Icb,
            "Fase C": Icc,
            "Seq. (0)": I0injetada,
            "Seq. (+)": I1injetada,
            "Seq. (-)": I2injetada
        })

        resultados['Correntes Injetadas nos Barramentos'].append(df)

    print("Correntes injetadas nos barramentos calculadas com sucesso!")
    # Retornar dict novamente, não a lista
    return resultados
