import xlwings as xw

def variaveis_sistema(caminho, nome_planilha):
    wb = xw.Book(fr"{caminho}")
    ws = wb.sheets[nome_planilha]
    potencia_base = ws.range('A1').value
    nome_caso_estudo = ws.range('A2').value
    unidade = ws.range('A3').value

    return potencia_base, nome_caso_estudo, unidade