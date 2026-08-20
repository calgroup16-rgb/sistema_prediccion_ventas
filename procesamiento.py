import pandas as pd
import numpy as np


def cargar_excel(archivo):
    df = pd.read_excel(archivo)

    # Normalizar nombres de columnas
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.upper()
    )

    columnas_requeridas = [
        "NOMBRE SN",
        "FABRICANTE",
        "DESCRIPCION"
    ]

    faltantes = [
        c for c in columnas_requeridas
        if c not in df.columns
    ]

    if faltantes:
        raise ValueError(
            f"Faltan estas columnas: {', '.join(faltantes)}"
        )

    # Limpiar campos principales
    for columna in ["NOMBRE SN", "FABRICANTE", "DESCRIPCION"]:
        df[columna] = (
            df[columna]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    return df


def detectar_columna_anio(df):

    posibles = [
        "AÑO",
        "ANO",
        "YEAR",
        "FECHA",
        "DATE"
    ]

    for columna in posibles:
        if columna in df.columns:
            return columna

    return None


def preparar_anio(df):

    columna = detectar_columna_anio(df)

    if columna is None:
        raise ValueError(
            "No se encontró una columna de año o fecha."
        )

    if columna in ["FECHA", "DATE"]:

        df["AÑO"] = pd.to_datetime(
            df[columna],
            errors="coerce"
        ).dt.year

    else:

        df["AÑO"] = pd.to_numeric(
            df[columna],
            errors="coerce"
        )

    df["AÑO"] = df["AÑO"].astype("Int64")

    return df


def detectar_columna_valor(df):

    posibles = [
        "TOTAL LINEA",
        "VALOR",
        "VENTAS",
        "VENTA",
        "TOTAL",
        "VALOR VENTA",
        "VALOR DE VENTA",
        "NETO",
        "IMPORTE"
    ]

    for columna in posibles:
        if columna in df.columns:
            return columna

    return None


def preparar_valor(df):

    columna = detectar_columna_valor(df)

    if columna is None:
        raise ValueError(
            "No se encontró una columna de valor de venta."
        )

    df["VALOR_VENTA"] = (
        df[columna]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    df["VALOR_VENTA"] = pd.to_numeric(
        df["VALOR_VENTA"],
        errors="coerce"
    ).fillna(0)

    return df
