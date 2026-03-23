import xlwings as xw # Leitura e Inserção de Dados no Excel
import pandas as pd # Leitura e Manipulação de Dados do Excel
from datetime import datetime # Para Controle de Tempo de Execução
import math # Cálculos Matemáticos
import cmath # Cálculos com Números Complexos
import numpy as np # Manipulação de Matrizes
import sys

from functions.correntes_curto_circuito import correntes_curto
from functions.tensoes_nas_barras import tensoes_barras
from functions.matrizes_y_e_z_barra import calcular_matrizes
from functions.correcao_defasagem import correcao_defasagem
from functions.corrente_nos_elementos import corrente_nos_elementos
from functions.correntes_injetadas import correntes_injetadas
from functions.converter_fasores import formatar_fasores
from functions.exportar_matrizes import exportar_matrizes
from functions.conversao_valores_reais import valores_reais


def curto_circuito():

    #----------------------------------------------------------------------------------------------------------------------
    #-------------------------------------- LEITURA DOS DADOS E TABELAS  --------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    # ABERTURA DO ARQUIVO PELO XLWINGS
    caminho = "xlfaults.xlsm" # O arquivo excel deve estar na mesma pasta do código Python
    wb = xw.Book(fr"{caminho}")

    # LEITURA DOS ELEMENTOS DO SISTEMA
    Maquina = pd.read_excel(fr"{caminho}", sheet_name = "Máquinas")
    Carga = pd.read_excel(fr"{caminho}", sheet_name = "Carga")
    Transformador = pd.read_excel(fr"{caminho}", sheet_name = "Transformadores")
    Linha = pd.read_excel(fr"{caminho}", sheet_name= "Linhas de Transmissão")
    Barra = pd.read_excel(fr"{caminho}", sheet_name= "Barramentos")
    Transformador3E = pd.read_excel(fr"{caminho}", sheet_name = "Transformadores 3 Enrolamentos") # Ainda não implementado...
    #print(Maquina, Carga, Transformador, Linha, Barra, Transformador3E) # Até aqui a leitura funcionou bem...
    print(Maquina)
    print(Carga)
    print(Transformador)
    print(Linha)
    print(Barra)

    # LEITURA DAS CONFIGURAÇÕES DA SIMULAÇÃO
    ws = wb.sheets['Oculto']
    potencia_base = ws.range('A1').value
    nome_caso_estudo = ws.range('A2').value
    unidade = ws.range('A3').value
    Configuracoes = pd.read_excel(fr"{caminho}", sheet_name= "Configurações") # Aqui poderemos rodar várias simulações
    #print(potencia_base, configuracoes)



    #----------------------------------------------------------------------------------------------------------------------
    #--------------------------------------------- VARIAVEIS  -------------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------
    
    # CRIAÇÃO DAS TABELAS DE RESULTADOS
    tabela_tensoes_barras = pd.DataFrame(columns=["Barra", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Utilizado nos cálculos
    tabela_tensoes_barras_corrigidos = pd.DataFrame(columns=["Barra", "Fase A", "Fase B", "Fase C", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Resultado Final

    tabela_correntes_contribuicao = pd.DataFrame(columns=["Trecho", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Utilizado nos cálculos
    tabela_correntes_contribuicao_corrigidos = pd.DataFrame(columns=["Trecho", "Fase A", "Fase B", "Fase C", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Resultado Final

    tabela_correntes_injetadas = pd.DataFrame(columns=["Elemento", "Barra", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Utilizado nos cálculos
    tabela_correntes_injetadas_corrigidos = pd.DataFrame(columns=["Elemento", "Barra", "Fase A", "Fase B", "Fase C", "Seq. (0)", "Seq. (+)", "Seq. (-)"]) # Resultado Final


    # DADOS EXTRAS/COMPLEMENTARES
    data = datetime.now().strftime('%d/%m/%Y %H:%M:%S')


    # MATRIZES DE TRANSFORMAÇÕES
    alfa = cmath.rect(1, math.radians(120)) # 1<120°

    T012abc =  np.array([
        [1,      1,               1  ],
        [1,    alfa**2,          alfa],
        [1,     alfa,          alfa**2]
    ])

    Tabc012 = (1/3) * np.array([
        [1,       1,            1   ],
        [1,     alfa,        alfa**2],
        [1,     alfa**2,       alfa ]
    ])

    resultados = {'Correntes de Falta': [],
                  'Tensões nas Barras': [],
                  'Correntes de Contribuição': [],
                  'Correntes Injetadas nos Barramentos': [],
                  'Tensões nas Barras - Não Corrigidas': [],
                  'Correntes de Contibuição - Não Corrigidas': [],
                  }

    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------------- CÁLCULO DAS IMPEDÂNCIAS -----------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    Maquina["Z1 (pu) Base"] = Maquina["R1 (pu) Base"] + 1j *  Maquina["X1 (pu) Base"]
    Maquina["ZN (pu) Base"] = Maquina["RN (pu) Base"] + 1j *  Maquina["XN (pu) Base"]
    Maquina["Z0 (pu) Base"] = Maquina["R0 (pu) Base"] + 1j * (Maquina["X0 (pu) Base"] + 3 * Maquina["ZN (pu) Base"])
    

    Transformador["Z1 (pu) Base"] = 0 + 1j * Transformador["X1 (pu) Base"]
    Transformador["Z0 (pu) Base"] = 0 + 1j * (Transformador["X0 (pu) Base"] + 3 * Transformador["XN P (pu) Base"] + 3 * Transformador["XN S (pu) Base"])

    Linha["Z1 (pu) Base"] = Linha["R1 (pu) Base"] + 1j * Linha["X1 (pu) Base"]
    Linha["Z0 (pu) Base"] = Linha["R0 (pu) Base"] + 1j * Linha["X0 (pu) Base"]

    Carga["Z1 (pu) Base"] = Carga["R1 (pu) Base"] + 1j * Carga["X1 (pu) Base"]
    Carga["ZN (pu) Base"] = Carga["RN (pu) Base"] + 1j * Carga["XN (pu) Base"]
    Carga["Z0 (pu) Base"] = Carga["R1 (pu) Base"] + 1j * (Carga["X1 (pu) Base"] + 3 * Carga["ZN (pu) Base"])

    Configuracoes["Z (pu) Base"] = Configuracoes["Impedância de Falta R (pu)"] + 1j * Configuracoes["Impedância de Falta X (pu)"]
    # Futuramente, para o Transformador 3 Enrolamentos...

    

    Ybarra12, Zbarra12, Ybarra0, Zbarra0 = calcular_matrizes(Maquina, Transformador, Linha, Carga, Barra)

    print(Ybarra12)
    print(Ybarra0)
    print(Zbarra12)
    print(Zbarra0)

    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------------- CORRENTES DE CURTO CIRCUITO -------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    resultados = correntes_curto(resultados, Configuracoes, Zbarra12, Zbarra0, T012abc, Barra, potencia_base)
    
    #---------------------------------------------------------------------------------------------------------------------
    #------------------------------------- TENSÕES NOS BARRAMENTOS -------------------------------------------------------
    #---------------------------------------------------------------------------------------------------------------------
    
    G = correcao_defasagem(Linha, Transformador) # Gera o grafo de correção de defasagem, por conta dos transformadores
    resultados = tensoes_barras(Barra, Zbarra12, Zbarra0, resultados, Configuracoes, G, T012abc)
    #resultados = tensoes_barras(Barra, Zbarra12, Zbarra0, resultados, Configuracoes, G, T012abc)


    #---------------------------------------------------------------------------------------------------------------------
    #------------------------------------- CORRENTE NOS ELEMENTOS --------------------------------------------------------
    #---------------------------------------------------------------------------------------------------------------------
    
    resultados = corrente_nos_elementos(Configuracoes, Linha, Maquina, Carga, Transformador, resultados, G, T012abc)
    
    #----------------------------------------------------------------------------------------------------------------------
    #----------------------------------- CÁLCULOS DAS CORRENTES INJETADAS -------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------
    
    resultados = correntes_injetadas(Ybarra12, Ybarra0, resultados, Configuracoes, Barra, T012abc)

    
    #---------------------------------------------------------------------------------------------------------------------
    #------------------------------------- EXPORTAR RESULTADOS -----------------------------------------------------------
    #---------------------------------------------------------------------------------------------------------------------
    if unidade == "Real":
        resultados = valores_reais(resultados, Barra, Configuracoes, potencia_base, Maquina, Carga, Linha, Transformador)
    resultados = formatar_fasores(resultados)
    print(resultados['Correntes de Falta'][0])


    exportar_matrizes(wb,Ybarra12, Zbarra12, Ybarra0, Zbarra0)
    
    with pd.ExcelWriter("resultados.xlsx") as writer:

        resultados['Correntes de Falta'][0].to_excel(
            writer, sheet_name="Correntes de Falta", index=False
        )
        resultados['Tensões nas Barras'][0].to_excel(
            writer, sheet_name="Tensões nas Barras", index=False
        )
        resultados['Correntes de Contribuição'][0].to_excel(
            writer, sheet_name="Correntes de Contribuição", index=False
        )
        resultados['Correntes Injetadas nos Barramentos'][0].to_excel(
            writer, sheet_name="Correntes Injetadas", index=False
        )

curto_circuito()