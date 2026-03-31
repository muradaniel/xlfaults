from datetime import datetime
import schemdraw  # Realizar os desenhos elétricos

from functions.leitura_e_tratamento_tabelas import tabelas_de_dados
from diagrama.legenda_tecnica import adicionar_margem_pdf
from diagrama.criacao_elementos import *  # noqa: F403
from diagrama.organizacao_grafo import Grafo
from functions.leitura_variaveis_sistemas import variaveis_sistema
from diagrama.dicionario_cores_tensoes import Gera_Cores_Tensoes
from diagrama.desenhando_barramento import Desenho_Barra
from diagrama.desenhando_maquina import Desenho_Maquina
from diagrama.desenho_transformador import Desenho_Transformador
from diagrama.desenhando_linha import Desenho_Linha
from diagrama.desenhando_carga import Desenho_Carga
#from diagrama.desenhando_transformador_3_enrolamentos import Desenho_Transformador_3_enrolamentos # Futuro...


def gerar_diagrama(caminho_excel="xlfaults.xlsm"):

    #----------------------------------------------------------------------------------------------------------------------
    #------------------------------- LEITURA DE DADOS DO EXCEL & VARIAVEIS  -----------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    Maquina, Carga, Transformador, Linha, Barra, Transformador3E, Configuracoes = tabelas_de_dados(caminho_excel)
    potencia_base, nome_caso_estudo, unidade = variaveis_sistema(caminho_excel, nome_planilha="Oculto")
    caminho_diagrama = fr"{caminho_excel.replace("xlsm", "pdf")}"
    simulador = "Xlfaults" # Nome do Simulador
    data = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------- MONTAGEM DO GRAFO PARA ORGANIZAÇÃO DOS ELEMENTOS ----------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    G, posicao_elementos = Grafo(Linha, Transformador, Maquina, Carga, Barra)

    #----------------------------------------------------------------------------------------------------------------------
    #---------------------------- GERA DICIONÁRIO DE CORES POR NÍVEL DE TENSÃO --------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    dicionario_cores = Gera_Cores_Tensoes(Barra["Tensão (kV)"].tolist())

    #----------------------------------------------------------------------------------------------------------------------
    #-------------------------------------- INÍCIO DO DIAGRAMA ------------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    with schemdraw.Drawing(file=fr"{caminho_diagrama}", dpi=150, theme='grade3', show=False) as d:
        d.config(fontsize=10) # Configura o tamanho da fonte para os rótulos dos elementos

    #----------------------------------------------------------------------------------------------------------------------
    #-------------------------------------- INSERE OS BARRAMENTOS ---------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

        Desenho_Barra(d, Barra, posicao_elementos, dicionario_cores)

    #----------------------------------------------------------------------------------------------------------------------
    #-------------------------------------- INSERE AS MÁQUINAS ------------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

        Desenho_Maquina(d, Maquina, posicao_elementos, dicionario_cores)

    #----------------------------------------------------------------------------------------------------------------------
    #-------------------------------------- INSERE OS TRANSFORMADORES -----------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

        Desenho_Transformador(d, Transformador, posicao_elementos, dicionario_cores)

    #----------------------------------------------------------------------------------------------------------------------
    #----------------------------------- INSERE AS LINHAS DE TRANSMISSÃO --------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

        Desenho_Linha(d, Linha, posicao_elementos, dicionario_cores)

    #----------------------------------------------------------------------------------------------------------------------
    #------------------------------------------ INSERE AS CARGAS ----------------------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------
        
        Desenho_Carga(d, Carga, posicao_elementos, dicionario_cores)

    #----------------------------------------------------------------------------------------------------------------------
    #----------------------------------- INSERE OS TRANSFORMADORES DE 3 ENROLAMENTOS --------------------------------------
    #----------------------------------------------------------------------------------------------------------------------
        
        #Desenho_Transformador_3_enrolamentos(d, Transformador_3E, posicao_elementos, dicionario_cores) # Futuro...

    #----------------------------------------------------------------------------------------------------------------------
    #----------------------------------- INSERE A LEGENDA NO DESENHO TÉCNICO ----------------------------------------------
    #----------------------------------------------------------------------------------------------------------------------

    adicionar_margem_pdf(
    caminho_diagrama,
    caminho_diagrama.replace(".pdf", "_final.pdf"),
    200,
    nome_caso_estudo,
    potencia_base,
    simulador,
    data,
    dicionario_cores
)


gerar_diagrama(caminho_excel="xlfaults.xlsm")
print("Diagrama Finalizado")