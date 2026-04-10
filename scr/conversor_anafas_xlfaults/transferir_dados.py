import xlwings as xw

def transferir_dados_para_excel(caminho, aba, tabela):

    # Abrir o Excel
    wb = xw.Book(caminho)  # ou xw.Book('arquivo.xlsx')
    ws = wb.sheets[aba]

    # escrever linha a linha
    linha_inicial = 2

    for i, (_, row) in enumerate(tabela.iterrows()):
        ws.range(f"A{linha_inicial + i}").value = row.tolist()