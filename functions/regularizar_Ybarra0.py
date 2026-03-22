import numpy as np
import pandas as pd

# --------------------------------------------------
# Detectar barras isoladas (linha e coluna zeradas)
# --------------------------------------------------
def barras_isoladas(Y, tol=1e-12):
    linhas_zero = np.all(np.abs(Y) < tol, axis=1)
    colunas_zero = np.all(np.abs(Y) < tol, axis=0)
    return np.where(linhas_zero & colunas_zero)[0]


# --------------------------------------------------
# Detectar ilhas (componentes conexos)
# --------------------------------------------------
def encontrar_ilhas(Y, tol=1e-12):
    n = Y.shape[0]
    visitado = [False] * n
    ilhas = []

    def dfs(i, componente):
        visitado[i] = True
        componente.append(i)

        for j in range(n):
            if not visitado[j] and i != j and abs(Y[i, j]) > tol:
                dfs(j, componente)

    for i in range(n):
        if not visitado[i]:
            componente = []
            dfs(i, componente)
            ilhas.append(componente)

    return ilhas


# --------------------------------------------------
# Função principal: regularizar Ybarra0 e obter Zbarra0
# --------------------------------------------------
def regularizar_Ybarra0(Ybarra0, Barra, tol=1e-12):

    Y = Ybarra0.values.astype(complex)

    # Detectar estrutura do sistema
    isoladas = barras_isoladas(Y, tol)
    ilhas = encontrar_ilhas(Y, tol)

    # Logs informativos
    if len(isoladas) > 0:
        print(f"⚠️ Barras isoladas: {isoladas}")

    if len(ilhas) > 1:
        print(f"⚠️ Ilhas detectadas: {ilhas}")

    # --------------------------------------------------
    # Escolha do método
    # --------------------------------------------------
    if len(ilhas) == 1 and len(isoladas) == 0:
        # Sistema normal
        Z = np.linalg.inv(Y)
    else:
        # Sistema com problema estrutural
        Z = np.linalg.pinv(Y)

    # --------------------------------------------------
    # Correção física para barras isoladas
    # --------------------------------------------------
    for k in isoladas:
        Z[k, :] = 0
        Z[:, k] = 0
        Z[k, k] = np.inf   # impede corrente infinita

    # --------------------------------------------------
    # Converter para DataFrame
    # --------------------------------------------------
    Zbarra0 = pd.DataFrame(
        Z,
        index=Barra['Número'].tolist(),
        columns=Barra['Número'].tolist()
    )

    return Zbarra0
