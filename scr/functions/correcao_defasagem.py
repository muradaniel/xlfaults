# OBS: Este cálculo serve para gerar um grafo direcionado e ponderado,
# onde as arestas representam as conexões entre os barramentos e os pesos representam a defasagem associada a cada conexão.

import networkx as nx

def correcao_defasagem(Linha, Transformador):
    G = nx.DiGraph()
    for tabela in [Linha, Transformador]: # Unicos elementos que conectam os barramentos
        for index, row in tabela.iterrows():
            if tabela is Transformador: # Se tabela for transformador
                defasagem = row['Defasagem (°)']
                G.add_edge(row['Barra de'], row['Barra para'], weight=defasagem)
                G.add_edge(row['Barra para'], row['Barra de'], weight=-defasagem)
            else: # Se a tabela for linha
                G.add_edge(row['Barra de'], row['Barra para'], weight=0)
                G.add_edge(row['Barra para'], row['Barra de'], weight=0)
    print("Grafo de defasagem montado com sucesso!")
    return G