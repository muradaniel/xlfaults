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
        return str(z)  # garante que valores reais também viram string

    for chave in resultados:
        for i in range(len(resultados[chave])):
            df = resultados[chave][i]
            if isinstance(df, pd.DataFrame):
                resultados[chave][i] = df.applymap(formatar)

    return resultados