import pandas as pd
import networkx as nx

# Essa função tem como objetivo resumir as ligações entre todos os componentes com o objetivo de organizar a posição das barras.
def Grafo(Linha, Transformador, Maquina, Carga, Barra):
    ligacao_entre_elementos = pd.DataFrame(columns=["Elemento", "Barra conectada"])
    for tabela in [Linha, Transformador]:
        for index, row in tabela.iterrows():
            ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra de"])]
            ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra para"])]
    for tabela in [Maquina, Carga]:
        for index, row in tabela.iterrows():
            ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra Conectada"])]
    # for index, row in Transformador_3E.iterrows():
    #     try:
    #         ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra P"])]
    #         ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra S"])]
    #         ligacao_entre_elementos.loc[len(ligacao_entre_elementos)] = [row["Nome"], int(row["Barra T"])]
    #     except (ValueError, TypeError):
    #         print(f"Aviso: erro ao processar linha {index} — Nome: {row['Nome']}")
    #         continue

    # Criando o Grafo:
    G = nx.Graph()
    G.add_nodes_from(Barra["Número"].tolist()) # Cria os nós dos Barramentos
    G.add_nodes_from(Transformador["Nome"].tolist()) # Cria os nós dos Transformadores
    G.add_nodes_from(Linha["Nome"].tolist()) # Cria os nós das Linhas de Transmissões
    G.add_nodes_from(Maquina["Nome"].tolist()) # Cria os nós das Máquinas (Motores e Geradores)
    G.add_nodes_from(Carga["Nome"].tolist()) # Cria os nós das Cargas sem Rotações
    #G.add_nodes_from(Transformador_3E["Nome"].tolist()) # Cria os nós dos Transformadores de 3 Enrolamentos

    
    G.add_edges_from(list(zip(ligacao_entre_elementos["Elemento"].tolist(), ligacao_entre_elementos["Barra conectada"].tolist()))) # Cria as ligações entre os nós (arestas)
    # Tipos de Layout
    #posicao_elementos = nx.spring_layout(G, scale=50, iterations=300000, threshold=1e-9) # Esolher a tipo de Organização do Grafo, espaçamento entre os nós, e número de iterações
    #posicao_elementos = pos = nx.nx_agraph.graphviz_layout(G, prog="sfdp")
    posicao_elementos = nx.kamada_kawai_layout(G, scale=10)
    # posicao_elementos = nx.planar_layout(G, scale=15)

    #posicao_elementos = nx.spectral_layout(G, scale=70)
    return G, posicao_elementos