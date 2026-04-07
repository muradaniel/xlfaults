import xlwings as xw
import pandas as pd
from xlwings.utils import col_name


def formatar_complexo(val):
    """Garante 4 casas decimais: 1.0000+0.0000j"""
    try:
        # Tenta tratar como número complexo/float
        real = float(val.real)
        imag = float(val.imag)
        y = f"{real:.4f}{imag:+.4f}j"
        # A ideia é boa, mas perde muito a formatação da matriz, principalmente a Ybarra
        # if real == 0 and imag == 0:
        #     y = "0"
        # elif real == 0 and imag != 0:
        #     y = f"{imag:.4f}j"
        # elif real != 0 and imag == 0:
        #     y = f"{real:.4f}"
        return y
    except (AttributeError, ValueError):
        return str(val)

def exportar_matrizes(caminho, Ybarra12, Zbarra12, Ybarra0, Zbarra0):
    nomes_planilhas = {
        "Y Barra (+) (-)": Ybarra12,
        "Z Barra (+) (-)": Zbarra12,
        "Y Barra (0)": Ybarra0,
        "Z Barra (0)": Zbarra0
    }

    wb = xw.Book(fr"{caminho}")
    for nome, df in nomes_planilhas.items():
        
        ws = wb.sheets[nome]
        
        # 1. APAGAR TUDO: Limpa conteúdo e formatações anteriores (cores, negrito, etc)
        ws.clear() 

        # Preparar os dados com o método moderno .map()
        df_formatado = df.map(formatar_complexo)

        ws.cells.api.HorizontalAlignment = -4108  # xlCenter

        # 2. INSERIR DADOS
        # index=True e header=True garantem que os nomes das barras apareçam
        alvo = ws.range('A1')
        alvo.options(pd.DataFrame, index=True, header=True).value = df_formatado
        ws.range('A1').value = f"Barramento"

        

        ultima_coluna = col_name(len(df.columns) + 1)
        ultima_linha = len(df.index) + 1

        alvo = ws.range(f"A1:{ultima_coluna}{ultima_linha}")  # Ajusta o intervalo para incluir o índice e o cabeçalho
        for border_id in range(7, 13):  # todas as bordas
            alvo.api.Borders(border_id).LineStyle = 1  # linha contínua
            alvo.api.Borders(border_id).Weight = 1     # espessura (opcional)

        # 3. NEGRITO NA PRIMEIRA LINHA E PRIMEIRA COLUNA
        # 'expand()' detecta automaticamente o tamanho da tabela inserida
        tabela = alvo.expand()
        tabela.row_height = 20    # altura confortável
        tabela.rows[0].api.Font.Bold = True     # Primeira linha (cabeçalho)
        tabela.columns[0].api.Font.Bold = True  # Primeira coluna (índice)
        tabela.rows[0].api.HorizontalAlignment = -4108  # Alinha o cabeçalho à direita para melhor visualização
        tabela.columns[0].api.HorizontalAlignment = -4108  # Alinha o índice à direita para melhor visualização
        # 4. AUTOAJUSTE DE LARGURA
        tabela.columns.autofit()

    print("Matrizes exportadas com sucesso para o Excel!")