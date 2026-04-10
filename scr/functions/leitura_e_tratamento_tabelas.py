# OBS1: Abaixo, foi implementado os valores complexos das impedâncias, no excel é seprado em Real e Imagninário

import pandas as pd
import numpy as np

def tabelas_de_dados(caminho):

    # Leitura dos dados das tabelas
    Maquina = pd.read_excel(fr"{caminho}", sheet_name = "Máquinas")
    Carga = pd.read_excel(fr"{caminho}", sheet_name = "Carga")
    Transformador = pd.read_excel(fr"{caminho}", sheet_name = "Transformadores")
    Linha = pd.read_excel(fr"{caminho}", sheet_name= "Linhas de Transmissão")
    Barra = pd.read_excel(fr"{caminho}", sheet_name= "Barramentos")
    #Transformador3E = pd.read_excel(fr"{caminho}", sheet_name = "Transformadores 3 Enrolamentos") # Ainda não implementado...
    Configuracoes = pd.read_excel(fr"{caminho}", sheet_name= "Configurações") # Aqui poderemos rodar várias simulações
    
    # Cálculos de Impedâncias na base do sistema
    Maquina["Z1 (pu) Base"] = Maquina["R1 (pu) Base"] + 1j *  Maquina["X1 (pu) Base"]
    Maquina["ZN (pu) Base"] = Maquina["RN (pu) Base"] + 1j *  Maquina["XN (pu) Base"]
    Maquina["Z0 (pu) Base"] = Maquina["R0 (pu) Base"] + 1j * (Maquina["X0 (pu) Base"]) + 3 * Maquina["ZN (pu) Base"]
    
    Transformador["Z1 (pu) Base"] = 0 + 1j * Transformador["X1 (pu) Base"]
    Transformador["ZN P (pu) Base"] = Transformador["RN P (pu) Base"] + 1j * Transformador["XN P (pu) Base"]
    Transformador["ZN S (pu) Base"] = Transformador["RN S (pu) Base"] + 1j * Transformador["XN S (pu) Base"]
    Transformador["Z0 (pu) Base"] = (Transformador["R0 (pu) Base"] + 1j * (Transformador["X0 (pu) Base"])) + 3 * (Transformador["ZN P (pu) Base"] + Transformador["ZN S (pu) Base"])

    Linha["Z1 (pu) Base"] = Linha["R1 (pu) Base"] + 1j * Linha["X1 (pu) Base"]
    Linha["Z0 (pu) Base"] = Linha["R0 (pu) Base"] + 1j * Linha["X0 (pu) Base"]

    Carga["Z1 (pu) Base"] = Carga["R1 (pu) Base"] + 1j * Carga["X1 (pu) Base"]
    Carga["Z0 (pu) Base"] = Carga["R1 (pu) Base"] + 1j * (Carga["X1 (pu) Base"])

    Configuracoes["Z (pu) Base"] = Configuracoes["Resistência de Falta (pu)"] + 1j * Configuracoes["Reatância de Falta (pu)"]

    modulo_tensao = Configuracoes["Módulo Tensão Pré-Falta (pu)"]
    angulo_tensao = np.deg2rad(Configuracoes["Ângulo Tensão Pré-Falta (graus)"])
    Configuracoes["Tensão Pré-Falta (pu)"] = modulo_tensao * (np.cos(angulo_tensao) + 1j * np.sin(angulo_tensao)) # Cria uma coluna complexa a partir do módulo e do ângulo em graus

     
    print("Tabelas de dados lidas e tratadas com sucesso!")
    return Maquina, Carga, Transformador, Linha, Barra, Configuracoes