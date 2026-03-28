import xlwings as xw
import pandas as pd

def formatar_complexo(val):
    """Garante 4 casas decimais: 1.0000+0.0000j"""
    try:
        # Tenta tratar como número complexo/float
        real = float(val.real)
        imag = float(val.imag)
        return f"{real:.4f}{imag:+.4f}j"
    except (AttributeError, ValueError):
        return str(val)

def exportar_matrizes(caminho, Ybarra12, Zbarra12, Ybarra0, Zbarra0):
    nomes_planilhas = {
        "Y Barra (+) (-)": Ybarra12,
        "Z Barra (+) (-)": Zbarra12,
        "Y Barra (0)": Ybarra0,
        "Z Barra (0)": Zbarra0
    }

    for nome, df in nomes_planilhas.items():
        wb = xw.Book(fr"{caminho}")
        ws = wb.sheets[nome]
        
        # 1. APAGAR TUDO: Limpa conteúdo e formatações anteriores (cores, negrito, etc)
        ws.clear() 

        # Preparar os dados com o método moderno .map()
        df_formatado = df.map(formatar_complexo)

        # 2. INSERIR DADOS
        # index=True e header=True garantem que os nomes das barras apareçam
        alvo = ws.range('A1')
        alvo.options(pd.DataFrame, index=True, header=True).value = df_formatado

        # 3. NEGRITO NA PRIMEIRA LINHA E PRIMEIRA COLUNA
        # 'expand()' detecta automaticamente o tamanho da tabela inserida
        tabela = alvo.expand()
        tabela.rows[0].api.Font.Bold = True     # Primeira linha (cabeçalho)
        tabela.columns[0].api.Font.Bold = True  # Primeira coluna (índice)
        tabela.rows[0].api.HorizontalAlignment = -4108  # Alinha o cabeçalho à direita para melhor visualização

        # 4. AUTOAJUSTE DE LARGURA
        tabela.columns.autofit()

    print("Planilhas limpas, matrizes exportadas e formatadas com sucesso!")