import numpy as np

def regularizar_Ybarra0(Ybarra0, tol=1e-10, eps=1e-10):
    
    isoladas = []
    Ynova = Ybarra0.copy()

    # 🔹 1. Detecta barras totalmente isoladas (linha e coluna zero)
    for barra in Ynova.index:
        linha_zero = np.all(np.abs(Ynova.loc[barra, :]) < tol)
        coluna_zero = np.all(np.abs(Ynova.loc[:, barra]) < tol)

        if linha_zero and coluna_zero:
            isoladas.append(barra)

    # 🔹 2. Regulariza barras isoladas
    for barra in isoladas:
        Ynova.loc[barra, :] = 0
        Ynova.loc[:, barra] = 0
        Ynova.loc[barra, barra] = eps

    # 🔹 3. Detecta e corrige ilhas de sequência zero (sem referência ao terra)
    for barra in Ynova.index:
        soma_linha = np.sum(Ynova.loc[barra, :])
        
        if abs(soma_linha) < tol:
            Ynova.loc[barra, barra] += eps

    return Ynova, isoladas
