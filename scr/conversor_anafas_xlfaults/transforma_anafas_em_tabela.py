import pandas as pd
from io import StringIO

def Anafas_para_tabelas(caminho):
    with open(caminho, 'r', encoding='latin-1') as f:
        texto = f.read()

    # =========================
    # 🔵 DBAR
    # =========================

    posicao_dbar = texto.find('DBAR') + 5 # não quero o DBAR, por isso o +5
    posicao_99999 = texto.find('99999', posicao_dbar) - 1 # Não pegar ultima linha em branco, por isso o -1
    dbar = texto[posicao_dbar:posicao_99999]

    # remover cabeçalhos
    linhas = dbar.split('\n')[2:]  # pula 2 primeiras linhas
    linhas = [l for l in linhas if l.strip() != '']

    texto_limpo = '\n'.join(linhas)

    # definir colunas (posição inicial, final)
    colspecs = [
        (0, 5), # NB 'bus number'
        (9, 20), # BN 'bus name'
        (31, 35) # VBAS 'base voltage'
    ]

    nomes = ['Número', 'Nome', 'Tensão (kV)']

    df_dbar = pd.read_fwf(StringIO(texto_limpo), colspecs=colspecs, names=nomes)
    df_dbar["Número"] = df_dbar["Número"].astype(int)
    df_dbar["Nome"] = df_dbar["Nome"].astype(str)
    df_dbar["Tensão (kV)"] = df_dbar["Tensão (kV)"].astype(float)


    # =========================
    # 🔴 DCIR
    # =========================

    # corta o texto depois do DBAR
    texto_restante = texto[posicao_99999:]
    
    posicao_dcir = texto_restante.find('DCIR')
    posicao_99999 = texto_restante.find('99999', posicao_dcir) - 1 # Não pegar ultima linha em branco, por isso o -1

    dcir = texto_restante[posicao_dcir:posicao_99999]

    # remover cabeçalhos
    linhas = dcir.split('\n')[3:]  # pula 3 primeiras linhas
    linhas = [l for l in linhas if l.strip() != '']

    texto_limpo = '\n'.join(linhas)

    # definir colunas (posição inicial, final)
    colspecs = [
        (0, 5), # BF 'bus from'
        (7, 12), # BT 'bus to
        (16, 17), # T 'Type'
        (17, 23), # R1 'resistance (+) (-)'
        (23, 29), # X1 'reactance (+) (-)'
        (29, 35), # R0 'resistance (0)'
        (35, 41), # X0 'reactance (0)'
        (41, 47), # CN 'Name'
        (72, 75), #DEF 'defasagem
        (80, 82), # CD 'conexao barra de
        (82, 88), # RNDE 'resistence neutro barra de'
        (88, 94), # XNDE 'reactance neutro barra de'
        (94, 96), # CP 'conexao barra para terra'
        (96, 102), # RNPA 'resistence para terra'
        (102, 108), # XNPA 'reactance para terra'
        (177, 184), # Sbase
    ]

    nomes = ['Barra de',
             'Barra para',
             'Tipo',
             'R1 (%)',
             'X1 (%)',
             'R0 (%)',
             'X0 (%)',
             'Nome',
             'Defasagem (°)',
             'Tipo de Conexão P',
             'RN P (%)',
             'XN P (%)',
             'Tipo de Conexão S',
             'RN S (%)',
             'XN S (%)',
             'Potência Nominal (MVA)'
    ]


    df_dcir = pd.read_fwf(StringIO(texto_limpo), colspecs=colspecs, names=nomes)

    df_dcir["Barra de"] = df_dcir["Barra de"].astype(int)
    df_dcir["Barra para"] = df_dcir["Barra para"].astype(int)
    df_dcir["Tipo"] = df_dcir["Tipo"].astype(str)
    df_dcir["R1 (%)"] = df_dcir["R1 (%)"].astype(float) / 100
    df_dcir["X1 (%)"] = df_dcir["X1 (%)"].astype(float) / 100
    df_dcir["R0 (%)"] = df_dcir["R0 (%)"].astype(float) / 100
    df_dcir["X0 (%)"] = df_dcir["X0 (%)"].astype(float) / 100
    df_dcir["Nome"] = df_dcir["Nome"].astype(str)
    df_dcir["Defasagem (°)"] = df_dcir["Defasagem (°)"].astype(float)
    df_dcir["Tipo de Conexão P"] = df_dcir["Tipo de Conexão P"].astype(str)
    df_dcir["RN P (%)"] = df_dcir["RN P (%)"].astype(float) / 100
    df_dcir["XN P (%)"] = df_dcir["XN P (%)"].astype(float) / 100
    df_dcir["Tipo de Conexão S"] = df_dcir["Tipo de Conexão S"].astype(str)
    df_dcir["RN S (%)"] = df_dcir["RN S (%)"].astype(float) / 100
    df_dcir["XN S (%)"] = df_dcir["XN S (%)"].astype(float) / 100
    df_dcir["Potência Nominal (MVA)"] = df_dcir["Potência Nominal (MVA)"].astype(int)


    return df_dbar, df_dcir