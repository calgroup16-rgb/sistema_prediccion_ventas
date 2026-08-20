import pandas as pd


PALABRAS_ANIMAL = [
    "ANIMAL",
    "ALIMENTO ANIMAL",
    "ALIMENTACION ANIMAL",
    "CONCENTRADO",
    "PREMEZCLA",
    "PREMIX",
    "VITAMINA",
    "MINERAL",
    "MINERALES",
    "ADITIVO",
    "ADITIVOS",
    "AMINOACIDO",
    "AMINOACIDOS",
    "PROBIOTICO",
    "PROBIOTICOS",
    "ENZIMA",
    "ENZIMAS",
    "LEVADURA",
    "NUTRICION",
    "NUTRICIONAL",
    "FORRAJE",
    "PIENSO",
    "SUPLEMENTO",
    "SUPLEMENTOS"
]


def identificar_productos_animal(df):

    patron = "|".join(PALABRAS_ANIMAL)

    mask = df["DESCRIPCION"].str.contains(
        patron,
        case=False,
        na=False
    )

    return df[mask].copy()


def ranking_clientes_animal(df):

    datos = identificar_productos_animal(df)

    ranking = (
        datos
        .groupby("NOMBRE SN")
        .agg(
            VENTAS_ANIMAL=("VALOR_VENTA", "sum"),
            PRODUCTOS=("DESCRIPCION", "nunique"),
            COMPRAS=("DESCRIPCION", "count")
        )
        .reset_index()
    )

    ranking["SCORE_ANIMAL"] = (
        ranking["VENTAS_ANIMAL"] * 0.6
        + ranking["PRODUCTOS"] * 0.2
        + ranking["COMPRAS"] * 0.2
    )

    ranking = ranking.sort_values(
        "SCORE_ANIMAL",
        ascending=False
    )

    return ranking.head(10)


def productos_animal_por_cliente(df, clientes):

    datos = identificar_productos_animal(df)

    datos = datos[
        datos["NOMBRE SN"].isin(clientes)
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

    return resultado.sort_values(
        "VALOR_VENTA",
        ascending=False
    )
