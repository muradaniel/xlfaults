import fitz  # PyMuPDF
import os


def adicionar_margem_pdf(
    arquivo_entrada,
    arquivo_saida,
    margem_pt,
    nome_caso,
    potencia_base,
    simulador,
    data,
    dicionario_cores
):

    # 🔹 Arquivo temporário (evita conflito)
    arquivo_temp = arquivo_saida.replace(".pdf", "_temp.pdf")

    # =========================
    # ETAPA 1: Criar PDF com margem
    # =========================
    with fitz.open(arquivo_entrada) as doc_original:
        novo_doc = fitz.open()

        for page in doc_original:
            largura, altura = page.rect.width, page.rect.height

            nova_pagina = novo_doc.new_page(
                width=largura + 2 * margem_pt,
                height=altura + 2 * margem_pt
            )

            nova_pagina.show_pdf_page(
                fitz.Rect(margem_pt, margem_pt, margem_pt + largura, margem_pt + altura),
                doc_original,
                page.number
            )

        # salva arquivo temporário
        if os.path.exists(arquivo_temp):
            os.remove(arquivo_temp)

        novo_doc.save(arquivo_temp)
        novo_doc.close()

    # =========================
    # ETAPA 2: Adicionar legenda
    # =========================
    doc = fitz.open(arquivo_temp)
    pagina = doc[0]

    largura = pagina.rect.width
    altura = pagina.rect.height

    # 📐 Layout
    distanciaH = 550
    distanciaV = 125

    x0 = largura - distanciaH
    y0 = altura - distanciaV

    # 🔲 Molduras
    pagina.draw_rect(
        fitz.Rect(x0, y0, largura - 20, altura - 20),
        color=(0, 0, 0),
        width=3,
        fill=(1, 1, 1)
    )

    pagina.draw_rect(
        fitz.Rect(20, 20, largura - 20, altura - 20),
        color=(0, 0, 0),
        width=2
    )

    # 🔹 Linhas horizontais
    pagina.draw_line((x0, y0 + 35), (largura - 20, y0 + 35), color=(0, 0, 0))
    pagina.draw_line((x0, y0 + 70), (largura - 20, y0 + 70), color=(0, 0, 0))
    pagina.draw_line((x0, y0 + 105), (largura - 20, y0 + 105), color=(0, 0, 0))

    # 🔹 Linha 1
    pagina.insert_text((x0 + 5, y0 + 10), "Nome do Caso:", fontsize=10)
    pagina.insert_text((x0 + 5, y0 + 32), nome_caso, fontsize=18)

    # 🔹 Linha 2
    pagina.insert_text((x0 + 5, y0 + 45), "Tipo de Estudo Elétrico:", fontsize=10)
    pagina.insert_text((x0 + 5, y0 + 67), "Análise de Curto Circuito", fontsize=18)

    pagina.insert_text((x0 + 230, y0 + 45), "Potência Base:", fontsize=10)
    pagina.insert_text((x0 + 230, y0 + 67), f"{potencia_base} MVA", fontsize=18)

    pagina.insert_text((x0 + 390, y0 + 45), "Software:", fontsize=10)
    pagina.insert_text((x0 + 390, y0 + 67), simulador, fontsize=18)

    # 🔹 Linha 3
    pagina.insert_text((x0 + 5, y0 + 80), "Data:", fontsize=10)
    pagina.insert_text((x0 + 5, y0 + 100), data, fontsize=18)

    pagina.insert_text((x0 + 230, y0 + 80), "Tensão Pré-Falta:", fontsize=10)
    pagina.insert_text((x0 + 230, y0 + 100), "1 < 0° pu", fontsize=18)

    pagina.insert_text((x0 + 345, y0 + 80), "Tipo de Representação:", fontsize=10)
    pagina.insert_text((x0 + 345, y0 + 100), "Diagrama Unifilar", fontsize=18)

    # 🔹 Linhas verticais
    pagina.draw_line((largura - 330, y0 + 35), (largura - 330, y0 + 105), color=(0, 0, 0))
    pagina.draw_line((largura - 170, y0 + 35), (largura - 170, y0 + 70), color=(0, 0, 0))
    pagina.draw_line((largura - 210, y0 + 70), (largura - 210, y0 + 105), color=(0, 0, 0))

    # 🎨 Legenda de cores
    inicio_x = x0 - 80
    inicio_y = y0 + 60
    espacamento = 140

    shape = pagina.new_shape()

    for i, (chave, valor) in enumerate(dicionario_cores.items()):
        x = inicio_x - i * espacamento

        pagina.insert_text((x, inicio_y + 8), f"{chave} kV", fontsize=18)

        shape.draw_circle(fitz.Point(x - 25, inicio_y), 10)
        shape.finish(color=valor, fill=valor)

    shape.commit()

    pagina.insert_text(
        (inicio_x - (len(dicionario_cores) * espacamento) - 80, inicio_y + 8),
        "Níveis de Tensão:",
        fontsize=20
    )

    # =========================
    # ETAPA 3: Salvar FINAL
    # =========================
    arquivo_saida = f"{nome_caso} - Diagrama Unifilar.pdf"

    if os.path.exists(arquivo_saida):
        os.remove(arquivo_saida)

    doc.save(arquivo_saida)
    doc.close()

    # 🔥 Remove temporário
    if os.path.exists(arquivo_temp):
        os.remove(arquivo_temp)

    # 🔥 Remove temporário
    if os.path.exists(arquivo_temp):
        os.remove(arquivo_temp)

    # 🔥 Remove PDF original (gerado pelo schemdraw)
    if os.path.exists(arquivo_entrada):
        os.remove(arquivo_entrada)
