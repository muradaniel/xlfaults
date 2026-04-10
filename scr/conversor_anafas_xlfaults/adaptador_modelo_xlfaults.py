from .procv_tensoes import ProcV_tensoes


def remover_colunas(df, colunas):
    for coluna in colunas:
        if coluna in df.columns:
            df = df.drop(coluna, axis=1)
    return df

def adaptador_de_tabela(dbar, dcir):

    Barra = dbar
    Linha = dcir[dcir["Tipo"] == "L"]
    Maquina = dcir[dcir["Tipo"] == "G"]
    Transformador = dcir[dcir["Tipo"] == "T"]
    Carga = dcir[dcir["Tipo"] == "C"]

    

    # Barra:
    cols = ['Tensão (kV)']
    Barra[cols] = Barra[cols].round(1)
    


    # Linha:
    Linha = remover_colunas(Linha, ['Tipo de Conexão P', 'Tipo de Conexão S', 'RN P (%)', 'XN P (%)', 'Tipo de Conexão S', 'RN S (%)', 'XN S (%)', 'Defasagem (°)', 'Tipo'])
    cols = ['R1 (%)', 'X1 (%)', 'R0 (%)', 'X0 (%)']
    Linha[cols] = Linha[cols].round(2)
    Linha = Linha.fillna(0)



    # Maquina:
    Maquina = remover_colunas(Maquina, ['Tipo de Conexão S', 'RN S (%)', 'XN S (%)', 'Defasagem (°)', 'Tipo', 'Barra para'])
    Maquina["Tipo de Máquina"] = "Gerador"
    Maquina['Tipo de Conexão P'] = Maquina['Tipo de Conexão P'].replace('YN', 'Yt')
    Maquina = Maquina.rename(columns={'Tipo de Conexão P': 'Tipo de Conexão'})
    Maquina = Maquina.rename(columns={'Barra de': 'Barra Conectada'})
    Maquina = Maquina.rename(columns={'RN P (%)': 'RN (%)'})
    Maquina = Maquina.rename(columns={'XN P (%)': 'XN (%)'})
    Maquina = Maquina.fillna(0)
    cols = ['R1 (%)', 'X1 (%)', 'R0 (%)', 'X0 (%)', 'RN (%)', 'XN (%)']
    Maquina[cols] = Maquina[cols].round(2)
    


    # Transformador:
    Transformador = Transformador.copy()
    Transformador['Tipo de Conexão P'] = Transformador['Tipo de Conexão P'].replace('YN', 'Yt')
    Transformador['Tipo de Conexão S'] = Transformador['Tipo de Conexão S'].replace('YN', 'Yt')
    Transformador['Tipo de Conexão'] = (Transformador['Tipo de Conexão P'] + '-' + Transformador['Tipo de Conexão S'])
    Transformador = remover_colunas(Transformador, ['Tipo', 'Tipo de Conexão P', 'Tipo de Conexão S'])
    Transformador = Transformador.fillna(0)
    cols = ['R1 (%)', 'X1 (%)', 'R0 (%)', 'X0 (%)', 'RN P (%)', 'XN P (%)', 'RN S (%)', 'XN S (%)', 'Defasagem (°)']
    Transformador[cols] = Transformador[cols].round(2)


    # Reordenar as colunas das tabelas
    Linha, Maquina, Transformador, Carga = ProcV_tensoes(Barra, Linha, Maquina, Transformador, Carga)
    Barra = Barra[['Nome', 'Número', 'Tensão (kV)']]
    Linha = Linha[['Nome', 'Barra de', 'Barra para', 'Tensão (kV)', 'Potência Nominal (MVA)', 'R1 (%)', 'R0 (%)', 'X1 (%)', 'X0 (%)']]
    Maquina = Maquina[['Nome', 'Tipo de Máquina','Barra Conectada', 'Potência Nominal (MVA)', 'Tensão (kV)', 'R1 (%)', 'R0 (%)', 'RN (%)', 'X1 (%)', 'X0 (%)', 'XN (%)', 'Tipo de Conexão']]
    Transformador = Transformador[['Nome', 'Potência Nominal (MVA)', 'Barra de', 'Tensão Primário (kV)', 'Barra para', 'Tensão Secundário (kV)', 'R1 (%)', 'R0 (%)', 'RN P (%)', 'RN S (%)', 'X1 (%)', 'X0 (%)', 'XN P (%)', 'XN S (%)', 'Tipo de Conexão', 'Defasagem (°)']]

    return Barra, Linha, Maquina, Transformador, Carga