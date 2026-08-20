import pandas as pd


def clientes_inactivos_2026(df):

    historico = df[df["AÑO"].between(2019, 2025)]

    clientes_historicos = set(
        historico["NOMBRE SN"].unique()
    )

    clientes_2026 = set(
        df[df["AÑO"] == 2026]["NOMBRE SN"].unique()
    )

    inactivos = clientes_historicos - clientes_2026

    df_inactivos = df[
        df["NOMBRE SN"].isin(inactivos)
    ]

    ventas_anuales = pd.pivot_table(
        df_inactivos,
        index="NOMBRE SN",
        columns="AÑO",
        values="VALOR_VENTA",
        aggfunc="sum",
        fill_value=0
    )

    for año in range(2019, 2026):

        if año not in ventas_anuales.columns:
            ventas_anuales[año] = 0

    ventas_anuales = ventas_anuales[
        list(range(2019, 2026))
    ]

    ventas_anuales["TOTAL_HISTORICO"] = (
        ventas_anuales.sum(axis=1)
    )

    ventas_anuales["ULTIMO_AÑO_COMPRA"] = (
        df_inactivos
        .groupby("NOMBRE SN")["AÑO"]
        .max()
    )

    ventas_anuales = (
        ventas_anuales
        .sort_values(
            "TOTAL_HISTORICO",
            ascending=False
        )
        .reset_index()
    )

    return ventas_anuales.head(20)


def productos_clientes_inactivos(df, clientes):

    datos = df[
        df["NOMBRE SN"].isin(clientes)
        & (df["AÑO"] <= 2025)
    ]

    resultado = (
        datos
        .groupby(
            [
                "NOMBRE SN",
                "FABRICANTE",
                "DESCRIPCION"
            ]
        )["VALOR_VENTA"]
        .sum()
        .reset_index()
    )

    resultado = resultado.sort_values(
        "VALOR_VENTA",
        ascending=False
    )

    return resultado
