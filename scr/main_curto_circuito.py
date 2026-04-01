import math # Cálculos Matemáticos
import cmath # Cálculos com Números Complexos
import numpy as np # Manipulação de Matrizes

from functions.correntes_curto_circuito import correntes_curto
from functions.tensoes_nas_barras import tensoes_barras
from functions.matrizes_y_e_z_barra import calcular_matrizes
from functions.correcao_defasagem import correcao_defasagem
from functions.corrente_nos_elementos import corrente_nos_elementos
from functions.correntes_injetadas import correntes_injetadas
from functions.converter_fasores import formatar_fasores
from functions.exportar_matrizes import exportar_matrizes
from functions.conversao_valores_reais import valores_reais
from functions.exportar_resultados import exportar_resultados
from functions.leitura_e_tratamento_tabelas import tabelas_de_dados
from functions.leitura_variaveis_sistemas import variaveis_sistema


def main(caminho):

    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------------------- LEITURA DE DADOS DO EXCEL  --------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    # LEITURA DOS ELEMENTOS DO SISTEMA
    Maquina, Carga, Transformador, Linha, Barra, Transformador3E, Configuracoes = tabelas_de_dados(caminho)

    # LEITURA DAS CONFIGURAÇÕES DO SISTEMA 
    potencia_base, nome_caso_estudo, unidade = variaveis_sistema(caminho, nome_planilha="Oculto")

    #----------------------------------------------------------------------------------------------------------------------
    #--------------------------------------------- VARIAVEIS  -------------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------
    
    # CRIAÇÃO DAS TABELAS DE RESULTADOS
    resultados = {'Correntes de Falta': [],
                  'Tensões nas Barras': [],
                  'Correntes de Contribuição': [],
                  'Correntes Injetadas nos Barramentos': [],
                  'Tensões nas Barras - Não Corrigidas': [],
                  #'Correntes de Contibuição - Não Corrigidas': [],
                  }

    # MATRIZES DE TRANSFORMAÇÕES
    alfa = cmath.rect(1, math.radians(120)) # 1<120°

    T012abc =  np.array([
        [1,      1,               1  ],
        [1,    alfa**2,          alfa],
        [1,     alfa,          alfa**2]
    ])

    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------------------- CÁLCULO DAS MATRIZES --------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    Ybarra12, Zbarra12, Ybarra0, Zbarra0 = calcular_matrizes(Maquina, Transformador, Linha, Carga, Barra)

    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------------- CORRENTES DE CURTO CIRCUITO -------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    resultados = correntes_curto(resultados, Configuracoes, Zbarra12, Zbarra0, T012abc, Barra, potencia_base)
    
    #---------------------------------------------------------------------------------------------------------------------
    #------------------------------------- TENSÕES NOS BARRAMENTOS -------------------------------------------------------
    #---------------------------------------------------------------------------------------------------------------------
    
    G = correcao_defasagem(Linha, Transformador) # Gera o grafo de correção de defasagem, por conta dos transformadores
    resultados = tensoes_barras(Barra, Zbarra12, Zbarra0, resultados, Configuracoes, G, T012abc)

    #---------------------------------------------------------------------------------------------------------------------
    #------------------------------------- CORRENTE NOS ELEMENTOS --------------------------------------------------------
    #---------------------------------------------------------------------------------------------------------------------
    
    resultados = corrente_nos_elementos(Configuracoes, Linha, Maquina, Carga, Transformador, resultados, G, T012abc)
    
    #----------------------------------------------------------------------------------------------------------------------
    #----------------------------------- CÁLCULOS DAS CORRENTES INJETADAS -------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------
    
    resultados = correntes_injetadas(Ybarra12, Ybarra0, resultados, Configuracoes, Barra, T012abc)

    #---------------------------------------------------------------------------------------------------------------------
    #----------------------------------- UNIDADE DE SAÍDA (REAL OU PU) ---------------------------------------------------
    #---------------------------------------------------------------------------------------------------------------------

    if unidade == "Real":
        resultados = valores_reais(resultados, Barra, Configuracoes, potencia_base, Maquina, Carga, Linha, Transformador)
    resultados = formatar_fasores(resultados)

    #---------------------------------------------------------------------------------------------------------------------
    #------------------------------------- EXPORTAR RESULTADOS -----------------------------------------------------------
    #---------------------------------------------------------------------------------------------------------------------

    exportar_matrizes(caminho, Ybarra12, Zbarra12, Ybarra0, Zbarra0)
    exportar_resultados(Ybarra12, Zbarra12, Ybarra0, Zbarra0, resultados, Configuracoes)


print("Simulação Finalizada")
main(caminho = "xlfaults.xlsm")