import sys

from conversor_anafas_xlfaults.interface_de_conversao import selecionar_arquivo
from conversor_anafas_xlfaults.transforma_anafas_em_tabela import Anafas_para_tabelas
from conversor_anafas_xlfaults.adaptador_modelo_xlfaults import adaptador_de_tabela
from conversor_anafas_xlfaults.transferir_dados import transferir_dados_para_excel


def main(caminho):
    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------------------- SELECIONAR CAMINHO ANAFAS  --------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    caminho_anafas = selecionar_arquivo()

    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------------------- TRANSFORMAR ANAFAS EM TABELAS  ----------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    df_dbar, df_dcir = Anafas_para_tabelas(caminho_anafas)

    #----------------------------------------------------------------------------------------------------------------------
    #------------------------------------------- ADPTAR AO MODELO XLFAULTS ------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    Barra, Linha, Maquina, Transformador, Carga = adaptador_de_tabela(df_dbar, df_dcir)

    #----------------------------------------------------------------------------------------------------------------------
    #----------------------------------------- TRANSFERIR DADOS PARA EXCEL ------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    transferir_dados_para_excel(caminho, "Barramentos", Barra)
    transferir_dados_para_excel(caminho, "Transformadores", Transformador)
    transferir_dados_para_excel(caminho, "Máquinas", Maquina)
    transferir_dados_para_excel(caminho, "Carga", Carga)
    transferir_dados_para_excel(caminho, "Linhas de Transmissão", Linha)

caminho = sys.argv[1]
main(caminho=caminho)
print("Conversão Realizada com Sucesso!")