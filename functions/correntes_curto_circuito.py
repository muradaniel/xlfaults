def correntes_curto(resultados, df_configuracoes, Zbarra12, Zbarra0, T012abc, Barra, potencia_base):
    import numpy as np
    import pandas as pd
    #from functions.conversao_complexa import cartesiano_polar
    
    for index, row in df_configuracoes.iterrows():

        tipo_de_curto = row["Tipo de Curto Circuito"]
        barra_curto = row["Barra de Ocorrência"]
        impedancia_falta = row["Z (pu) Base"]

        if tipo_de_curto == "Trifásico":
            Ifa0 = 0
            Ifa1 = (1 / (Zbarra12.loc[barra_curto, barra_curto] + 3 * impedancia_falta))
            Ifa2 = 0
            IF = Ifa1
            If3f = T012abc @ np.array([[Ifa0],[Ifa1],[Ifa2]])

        elif tipo_de_curto == "Monofásico":
            Ifa0 = (1 / (2 * Zbarra12.loc[barra_curto, barra_curto] + Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta)) # Adicionar a Zf
            Ifa1 = (1 / (2 * Zbarra12.loc[barra_curto, barra_curto] + Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta)) # Adicionar a Zf
            Ifa2 = (1 / (2 * Zbarra12.loc[barra_curto, barra_curto] + Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta)) # Adicionar a Zf
            IF = 3 * Ifa0
            If3f = T012abc @ np.array([[Ifa0], [Ifa1], [Ifa2]])

        elif tipo_de_curto == "Bifásico":
            Ifa0 = 0
            Ifa1 = (1 / 2 * Zbarra12.loc[barra_curto, barra_curto] + impedancia_falta) # Adicionar a Zf
            Ifa2 = -(1 / 2 * Zbarra12.loc[barra_curto, barra_curto] + impedancia_falta) # Adicionar a Zf
            If3f = T012abc @ np.array([[Ifa0], [Ifa1], [Ifa2]])
            IF = If3f[1][0] # Fase B

        elif tipo_de_curto == "Bifásico - Terra":
            Ifa1 = 1 / (
                    Zbarra12.loc[barra_curto, barra_curto] +
                    (
                            Zbarra12.loc[barra_curto, barra_curto] *
                            (Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta)
                    ) / (
                            Zbarra12.loc[barra_curto, barra_curto] +
                            Zbarra0.loc[barra_curto, barra_curto] +
                            3 * impedancia_falta
                    )
            )
            Ifa0 = - Ifa1 * (Zbarra12.loc[barra_curto, barra_curto] / (Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta + Zbarra12.loc[barra_curto, barra_curto]))
            Ifa2 = - Ifa1 * ((Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta) / (Zbarra0.loc[barra_curto, barra_curto] + 3 * impedancia_falta + Zbarra12.loc[barra_curto, barra_curto]))
            IF = 3 * Ifa0
            If3f = T012abc @ np.array([[Ifa0], [Ifa1], [Ifa2]])
        else:
            print("Ocorreu Algum Erro, Conferir...")


        Vb = Barra.loc[Barra["Número"] == barra_curto, "Tensão (kV)"].values[0]
        IF *= (potencia_base / (np.sqrt(3) * Vb))
        print(f"IF (pu) = {IF}")
        Ifa = If3f[0][0] #* (potencia_base / (np.sqrt(3) * Vb))
        Ifb = If3f[1][0] #* (potencia_base / (np.sqrt(3) * Vb))
        Ifc = If3f[2][0] #* (potencia_base / (np.sqrt(3) * Vb))

        tabela_curto_circuito = pd.DataFrame([{
            "IF": IF,
            "Fase A": Ifa,
            "Fase B": Ifb,
            "Fase C": Ifc,
            "Seq. (0)": Ifa0,
            "Seq. (+)": Ifa1,
            "Seq. (-)": Ifa2
        }])

        resultados['Correntes de Falta'].append(tabela_curto_circuito)

    return resultados

