def formatar_fasores(resultados, tol=1e-6):
    import numpy as np
    import pandas as pd

    def formatar(z):
        if isinstance(z, complex):
            mag = abs(z)
            if mag < tol:
                return "0.0000 ∠ 0.00°"
            ang = np.angle(z, deg=True)
            return f"{mag:.4f} ∠ {ang:.2f}°"
        return str(z)

    for chave in resultados:
        nova_lista = []

        for df in resultados[chave]:
            if isinstance(df, pd.DataFrame):
                df_formatado = df.map(formatar)  # ✅ cria NOVO DataFrame
                nova_lista.append(df_formatado)
            else:
                nova_lista.append(df)

        resultados[chave] = nova_lista  # ✅ substitui a lista inteira

    return resultados
