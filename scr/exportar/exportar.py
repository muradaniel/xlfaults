import xlwings as xw


def Gera_subtitulo(sheet, ultima_linha, unidade, titulo):
    sheet.range(f"B{ultima_linha}:H{ultima_linha}").merge() # Mescla as células para o título
    titulo_caso_estudo = sheet.range(f"B{ultima_linha}") # Define a célula onde o título do caso de estudo será inserido
    if unidade == "Real" and "Correntes" in titulo:
            titulo_caso_estudo.value = f"{titulo} (kA)"
    elif unidade == "Real" and "Tensões" in titulo:
            titulo_caso_estudo.value = f"{titulo} (kV)"
    else:
        titulo_caso_estudo.value = f"{titulo} (PU)"
    titulo_caso_estudo.font.bold = True # Deixa o título em negrito
    titulo_caso_estudo.font.size = 12 # Define o tamanho da fonte do título
    titulo_caso_estudo.api.HorizontalAlignment = -4108  # xlCenter
    titulo_caso_estudo.color = (217, 217, 217) # pintando de Cinza
    ultima_linha += 1
    return ultima_linha


def Gera_cabecalhos(sheet, colunas, ultima_linha, cor):
    cabecalho_curto = sheet.range(f"B{ultima_linha}:H{ultima_linha}")
    cabecalho_curto.value = colunas
    cabecalho_curto.font.size = 12
    cabecalho_curto.api.HorizontalAlignment = -4108  # xlCenter
    cabecalho_curto.color = (0, 0, 0) # pintando de Preto
    cabecalho_curto.font.color = cor
    cabecalho_curto.font.bold = True # Deixa o cabeçalho em negrito
    ultima_linha += 1
    return ultima_linha


def Inserir_tabela(sheet, resultados, nome_tabela, index_configuracoes, ultima_linha):
    sheet.range(f"B{ultima_linha}").options(index=False, header=False).value = resultados[nome_tabela][index_configuracoes]
    ultima_linha += len(resultados[nome_tabela][index_configuracoes]) # Incrementa a última linha com o número de linhas da tabela mais uma linha de espaçamento
    return ultima_linha

def gerar_relatorio_completo(caminho, resultados, Configuracoes, potencia_base, nome_caso_estudo, unidade, data):

    wb = xw.Book(caminho)
    sheet = wb.sheets["Relatório"]

    sheet.api.Cells.Clear() # Limpa o conteúdo da planilha antes de gerar o relatório
    sheet.cells.api.HorizontalAlignment = -4108  # xlCenter
    sheet.range('B:B').font.bold = True
    
    # Título do Relatório
    sheet.range('B2:H2').merge() # Mescla as células para o título
    titulo = sheet.range('B2:H2') # Define a célula onde o título será inserido
    titulo.value = nome_caso_estudo # Insere o nome do caso de estudo como título
    titulo.font.bold = True # Deixa o título em negrito
    titulo.font.size = 26 # Define o tamanho da fonte do título
    titulo.api.HorizontalAlignment = -4108  # xlCenter
    # Borda superior (simples)
    titulo.api.Borders(8).LineStyle = 1
    titulo.api.Borders(8).Weight = 2
    # Borda inferior (dupla)
    titulo.api.Borders(9).LineStyle = -4119
    titulo.api.Borders(9).Weight = 4


    # Data da Simulação
    data_atual = sheet.range('B4')
    data_atual.value = f"Data: {data}" # Insere a data atual # Alterar no futuro
    data_atual.font.size = 12 # Define o tamanho da fonte da data
    data_atual.font.bold = True # Deixa a data em negrito
    data_atual.api.HorizontalAlignment = -4131  # xlLeft

    # Caso de Estudo
    ultima_linha = 5
    for index_configuracoes, row_configuracao in Configuracoes.iterrows():

        # Gerar Título do Caso de Estudo
        sheet.range(f"B{ultima_linha}:H{ultima_linha}").merge() # Mescla as células para o título do caso de estudo
        titulo_caso_estudo = sheet.range(f"B{ultima_linha}") # Define a célula onde o título do caso de estudo será inserido
        texto = f"Curto Circuito {row_configuracao['Tipo de Curto Circuito']}"
        texto += f" na Barra {row_configuracao['Barra de Ocorrência']}"
        texto += f" - Tensão Pré Falta {row_configuracao['Módulo Tensão Pré-Falta (pu)']} ∠ {row_configuracao['Ângulo Tensão Pré-Falta (graus)']}° PU"
        texto += f" - Impedância de falta {(complex(round(row_configuracao['Z (pu) Base'].real, 2), round(row_configuracao['Z (pu) Base'].imag, 2)))} PU" 
        titulo_caso_estudo.value = texto # Insere o título do caso de estudo
        titulo_caso_estudo.font.bold = True # Deixa o título em negrito
        titulo_caso_estudo.font.size = 18 # Define o tamanho da fonte do título
        titulo_caso_estudo.api.HorizontalAlignment = -4108  # xlCenter
        titulo_caso_estudo.color = (217, 217, 217) # pintando de Cinza
        ultima_linha += 1

        # Gerar Subtítulo das Correntes de Curto Circuito
        ultima_linha = Gera_subtitulo(sheet, ultima_linha, unidade, "Correntes de Curto Circuito")
        # Gerar Cabeçalho da tabela de Curto circuito
        ultima_linha = Gera_cabecalhos(sheet, ['IF"', 'Fase A', 'Fase B', 'Fase C', 'Seq. (0)', 'Seq. (+)', 'Seq. (-)'], ultima_linha, (142, 217, 115))
        # Insere os dados da tabela de Curto Circuito
        sheet.range(f"B{ultima_linha}").options(index=False, header=False).value = resultados['Correntes de Falta'][index_configuracoes]
        ultima_linha = Inserir_tabela(sheet, resultados, 'Correntes de Falta', index_configuracoes, ultima_linha)


        # Gerar Subtítulo das Tensões nos Barramentos
        ultima_linha = Gera_subtitulo(sheet, ultima_linha, unidade, "Tensões nos Barramentos")
        # Gerar Cabeçalho da tabela de Tensões nos Barramentos
        ultima_linha = Gera_cabecalhos(sheet, ['Barra', 'Fase A', 'Fase B', 'Fase C', 'Seq. (0)', 'Seq. (-)', 'Seq. (+)'], ultima_linha, (166, 200, 236))
        # Insere os dados da tabela de Tensões nos Barramentos
        ultima_linha = Inserir_tabela(sheet, resultados, 'Tensões nas Barras', index_configuracoes, ultima_linha)


        # Gerar Subtítulo das Correntes nos Elementos
        ultima_linha = Gera_subtitulo(sheet, ultima_linha, unidade, "Correntes de Contribuição")
        # Gerar Cabeçalho da tabela de Correntes de Contribuição
        ultima_linha = Gera_cabecalhos(sheet, ['Elemento', 'Fase A', 'Fase B', 'Fase C', 'Seq. (0)', 'Seq. (-)', 'Seq. (+)'], ultima_linha, (255, 113, 79))
        # Insere os dados da tabela de Correntes de Contribuição
        ultima_linha = Inserir_tabela(sheet, resultados, 'Correntes de Contribuição', index_configuracoes, ultima_linha)


        # Gerar Subtítulo das CorrentesInjetadas nos Barramentos
        ultima_linha = Gera_subtitulo(sheet, ultima_linha, unidade, "Correntes Injetadas nos Barramentos")
        # Gerar Cabeçalho da tabela de Correntes Injetadas nos Barramentos
        ultima_linha = Gera_cabecalhos(sheet, ['Barra', 'Fase A', 'Fase B', 'Fase C', 'Seq. (0)', 'Seq. (-)', 'Seq. (+)'], ultima_linha, (255, 255, 0))
        # Insere os dados da tabela de Correntes Injetadas nos Barramentos
        ultima_linha = Inserir_tabela(sheet, resultados, 'Correntes Injetadas nos Barramentos', index_configuracoes, ultima_linha)

        ultima_linha += 3
    




