import pandas as pd

def exportar_resultados(Ybarra12, Zbarra12, Ybarra0, Zbarra0, resultados, Configuracoes):
    
    chaves_desejadas = [
        'Correntes de Falta',
        'Tensões nas Barras',
        'Correntes de Contribuição',
        'Correntes Injetadas nos Barramentos'
    ]
    with pd.ExcelWriter("resultados.xlsx") as writer:

        # 🔹 Exporta matrizes
        Ybarra12.to_excel(writer, sheet_name="Ybarra (+) (-)", index=True)
        Zbarra12.to_excel(writer, sheet_name="Zbarra (+) (-)", index=True)
        Ybarra0.to_excel(writer, sheet_name="Ybarra (0)", index=True)
        Zbarra0.to_excel(writer, sheet_name="Zbarra (0)", index=True)

        # 🔹 Loop por cenário de curto
        for i, (_, row) in enumerate(Configuracoes.iterrows()):

            dfs_empilhados = []

            for chave in chaves_desejadas:
                
                lista_dfs = resultados.get(chave, [])

                # Segurança
                if i >= len(lista_dfs):
                    continue

                df = lista_dfs[i].copy()
                df["Tipo"] = chave
                dfs_empilhados.append(df)

            if not dfs_empilhados:
                continue

            # 🔹 Empilha tudo
            df_final = pd.concat(dfs_empilhados, ignore_index=True)

            # 🔹 Criar coluna "Elemento / IF"
            def definir_elemento(row_df):
                if pd.notna(row_df.get("IF")) and row_df.get("IF") != "":
                    return row_df["IF"]  # valor do curto
                elif pd.notna(row_df.get("Barra")):
                    return f"Barra {int(row_df['Barra'])}"
                elif pd.notna(row_df.get("Nome")):
                    return row_df["Nome"]
                else:
                    return ""

            df_final["Elemento / IF"] = df_final.apply(definir_elemento, axis=1)

            # 🔹 Remover colunas desnecessárias (mantém IF!)
            colunas_remover = [c for c in ["Barra", "Nome"] if c in df_final.columns]
            df_final = df_final.drop(columns=colunas_remover)

            # 🔹 Reordenar colunas
            colunas = df_final.columns.tolist()
            colunas = ["Tipo", "Elemento / IF"] + [c for c in colunas if c not in ["Tipo", "Elemento / IF"]]
            df_final = df_final[colunas]
            df_final.drop(columns=["IF"], inplace=True)

            # 🔹 Nome da aba (limite Excel = 31)
            nome_aba = f"Curto {row['Tipo de Curto Circuito']} - Barra {row['Barra de Ocorrência']}"
            nome_aba = nome_aba[:31]

            # 🔹 Exporta
            df_final.to_excel(writer, sheet_name=nome_aba, index=False)
