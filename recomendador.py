import pandas as pd
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.preprocessing import MultiLabelBinarizer


def construir_matriz_clientes(df):

    matriz = pd.crosstab(
        df["NOMBRE SN"],
        df["DESCRIPCION"]
    )

    matriz = (matriz > 0).astype(int)

    return matriz


def recomendar_productos(
    df,
    clientes_objetivo,
    cantidad_similares=10,
    cantidad_recomendaciones=10
):

    matriz = construir_matriz_clientes(df)

    clientes_existentes = [
        c for c in clientes_objetivo
        if c in matriz.index
    ]

    resultados = []

    similitud = cosine_similarity(matriz)

    similitud_df = pd.DataFrame(
        similitud,
        index=matriz.index,
        columns=matriz.index
    )

    for cliente in clientes_existentes:

        similares = (
            similitud_df[cliente]
            .drop(cliente)
            .sort_values(ascending=False)
            .head(cantidad_similares)
        )

        productos_cliente = set(
            matriz.loc[cliente]
            [matriz.loc[cliente] > 0]
            .index
        )

        productos_similares = matriz.loc[
            similares.index
        ].sum().sort_values(
            ascending=False
        )

        recomendaciones = (
            productos_similares
            .drop(labels=productos_cliente, errors="ignore")
            .head(cantidad_recomendaciones)
        )

        for producto, puntuacion in recomendaciones.items():

            resultados.append({
                "CLIENTE": cliente,
                "PRODUCTO_RECOMENDADO": producto,
                "PUNTUACION": puntuacion,
                "CLIENTES_SIMILARES": len(similares)
            })

    return pd.DataFrame(resultados)
