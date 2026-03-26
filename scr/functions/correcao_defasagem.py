import networkx as nx

def correcao_defasagem(Linha, Transformador):
    G = nx.DiGraph()
    for tabela in [Linha, Transformador]: # Unicos elementos que conectam os barramentos
        for index, row in tabela.iterrows():
            if tabela is Transformador: # Se tabela for transformador
                defasagem = row['Defasagem (°)']
                G.add_edge(row['Barra de'], row['Barra para'], weight=defasagem)
                G.add_edge(row['Barra para'], row['Barra de'], weight=-defasagem)
                #tipo = row["Tipo de Conexão"]
                #if tipo in ["Yt-D", "Y-D"]:
                    #G.add_edge(row['Barra de'], row['Barra para'], weight=-30)
                    #G.add_edge(row['Barra para'], row['Barra de'], weight=30)
                #elif tipo in ["D-Yt", "D-Y"]:
                    #G.add_edge(row['Barra de'], row['Barra para'], weight=30)
                    #G.add_edge(row['Barra para'], row['Barra de'], weight=-30)
                #else:
                    #G.add_edge(row['Barra de'], row['Barra para'], weight=0)
                    #G.add_edge(row['Barra para'], row['Barra de'], weight=0)
            else: # Se a tabela for linha
                G.add_edge(row['Barra de'], row['Barra para'], weight=0)
                G.add_edge(row['Barra para'], row['Barra de'], weight=0)
    return G