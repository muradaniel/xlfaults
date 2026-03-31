import numpy as np

def valores_reais(resultados, Barra, Configuracoes, potencia_base, Maquinas, Carga, Linha, Transformador):
    for index, row in Configuracoes.iterrows():

        # Para Correntes de Curto...        
        Vb = Barra.loc[Barra["Número"] == row["Barra de Ocorrência"], "Tensão (kV)"].values[0]
        Ib = potencia_base / (np.sqrt(3) * Vb)
        df = resultados['Correntes de Falta'][index]
        df = df.map(lambda x: x * Ib)
        resultados['Correntes de Falta'][index] = df

        
        map_maquinas = Maquinas.set_index("Nome")["Barra Conectada"]
        map_cargas = Carga.set_index("Nome")["Barra Conectada"]
        map_linhas = Linha.set_index("Nome")["Barra de"] # ou "Barra para", tanto faz, pois a corrente é a mesma
        map_transformadores_p = Transformador.set_index("Nome")["Barra de"]
        map_transformadores_s = Transformador.set_index("Nome")["Barra para"]
        map_vb = Barra.set_index("Número")["Tensão (kV)"]


        # Para Correntes de Contribuição...
        df = resultados['Correntes de Contribuição'][index]
        colunas = df.columns.difference(["Nome"])

        for index_contribuicao, row_contribuicao in df.iterrows():
            nome = row_contribuicao["Nome"]
            
            if nome in map_maquinas:
                barra = map_maquinas[nome]
                barra = int(np.real(barra))
                Vb = map_vb[barra]
                Ib = potencia_base / (np.sqrt(3) * Vb)

                df.loc[index_contribuicao, colunas] *= Ib
            elif nome in map_cargas:
                barra = map_cargas[nome]
                barra = int(np.real(barra))
                Vb = map_vb[barra]
                Ib = potencia_base / (np.sqrt(3) * Vb)

                df.loc[index_contribuicao, colunas] *= Ib

            elif nome in map_linhas:
                barra = map_linhas[nome]
                barra = int(np.real(barra))
                Vb = map_vb[barra]
                Ib = potencia_base / (np.sqrt(3) * Vb)

                df.loc[index_contribuicao, colunas] *= Ib

            elif nome.replace(" (Primário)", "") in map_transformadores_p:
                barra = map_transformadores_p[nome.replace(" (Primário)", "")]
                barra = int(np.real(barra))
                Vb = map_vb[barra]
                Ib = potencia_base / (np.sqrt(3) * Vb)

                df.loc[index_contribuicao, colunas] *= Ib
            elif nome.replace(" (Secundário)", "") in map_transformadores_s:
                barra = map_transformadores_s[nome.replace(" (Secundário)", "")]
                barra = int(np.real(barra))
                Vb = map_vb[barra]
                Ib = potencia_base / (np.sqrt(3) * Vb)

                df.loc[index_contribuicao, colunas] *= Ib

        resultados['Correntes de Contribuição'][index] = df

        # Para Correntes Injetadas nos Barramentos...
        df = resultados['Correntes Injetadas nos Barramentos'][index]
        for index_injetadas, row_injetadas in df.iterrows():
            barra = row_injetadas["Barra"]
            barra = int(np.real(barra))
            Vb = map_vb[barra]
            Ib = potencia_base / (np.sqrt(3) * Vb)

            colunas_injetadas = df.columns.difference(["Barra"])
            df.loc[index_injetadas, colunas_injetadas] *= Ib
        resultados['Correntes Injetadas nos Barramentos'][index] = df


        # Para Tensões nas Barras...
        df = resultados['Tensões nas Barras'][index]
        for index_tensoes, row_tensoes in df.iterrows():
            barra = row_tensoes["Barra"] # por algum motivo ta retornando complexo, então converti pra inteiro
            barra = int(np.real(barra))
            Vb = map_vb[barra]

            colunas_tensoes = df.columns.difference(["Barra"])
            df.loc[index_tensoes, colunas_tensoes] *= (Vb / np.sqrt(3))
        
        resultados['Tensões nas Barras'][index] = df
    print("Valores PU x REAL convertidos com sucesso!")
    return resultados