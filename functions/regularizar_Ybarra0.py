import numpy as np
import pandas as pd


def regularizar_Ybarra0(Ybarra0, tol=1e-6):
    return

def ilhas(Ybarra0, tol=1e-6):
    return

def barra_isolada(Ybarra0, Barra, tol=1e-6):
    print(Ybarra0)
    Y = Ybarra0.values
    indices_isoladas = []
    for i in range(Y.shape[0]):
        soma_linha = np.sum(Y[i, :])
        soma_coluna = np.sum(Y[:, i])
        elemento_diagonal = Y[i, i]
        total = soma_linha + soma_coluna - 2 * elemento_diagonal
        if total <= tol: # Verifica se tem barra isolada
            indices_isoladas.append(i)
    manter = [i for i in range(Y.shape[0]) if i not in indices_isoladas]
    Ybarra_reg = Y[np.ix_(manter, manter)]
    return Ybarra_reg, indices_isoladas



Ybarra0 = pd.DataFrame(np.array([
                    [4, 0, 0],
                    [0, 2, 9], 
                    [0, 9, 9]]), 
                    index=[1, 2, 3], 
                    columns=[1, 2, 3])
print(barra_isolada(Ybarra0))