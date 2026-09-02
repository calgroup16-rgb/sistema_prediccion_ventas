import pandas as pd
import numpy as np
import pandas as pd

def obtener_top20_inactivos_2026(df):
    """Obtiene los 20 clientes con mayores ventas entre 2019 y 2025 que NO compraron en 2026."""
    if 'FECHA' not in df.columns or 'VALOR_VENTA' not in df.columns or 'CLIENTE' not in df.columns:
        return pd.DataFrame()

    df_temp = df.copy()
    df_temp['FECHA'] = pd.to_datetime(df_temp['FECHA'], errors='coerce')
    df_temp['ANIO'] = df_temp['FECHA'].dt.year

    # Clientes que compraron en 2026
    clientes_2026 = set(df_temp[df_temp['ANIO'] == 2026]['CLIENTE'].unique())

    # Ventas 2019 - 2025
    df_historico = df_temp[(df_temp['ANIO'] >= 2019) & (df_temp['ANIO'] <= 2025)]

    # Excluir los que volvieron a comprar en 2026
    df_inactivos = df_historico[~df_historico['CLIENTE'].isin(clientes_2026)]

    # Top 20 por monto total acumulado
    top20 = (
        df_inactivos.groupby('CLIENTE')['VALOR_VENTA']
        .sum()
        .nlargest(20)
        .reset_index()
    )
    return top20

def obtener_top10_alimentacion_animal(df):
    """Obtiene el Top 10 de clientes activos en 2026 que compran productos de alimentación animal."""
    if 'FECHA' not in df.columns or 'PRODUCTO' not in df.columns:
        return pd.DataFrame()

    df_temp = df.copy()
    df_temp['FECHA'] = pd.to_datetime(df_temp['FECHA'], errors='coerce')
    df_temp['ANIO'] = df_temp['FECHA'].dt.year

    # Palabras clave del sector
    keywords = ['ALIMENTO', 'BALANCEADO', 'FORRAJE', 'GANADO', 'AVICOLA', 'PORCINO', 'PREMEZCLA', 'HARINA', 'MASCOTAS']
    patron = '|'.join(keywords)

    # Filtrar activos en 2026 que contengan palabras clave en productos
    df_2026 = df_temp[df_temp['ANIO'] == 2026]
    df_animal = df_2026[df_2026['PRODUCTO'].str.contains(patron, case=False, na=False)]

    top10 = (
        df_animal.groupby('CLIENTE')['VALOR_VENTA']
        .sum()
        .nlargest(10)
        .reset_index()
    )
    return top10

def cargar_excel(archivo):
    df = pd.read_excel(archivo)
# Filtrar solo FACTURAS para todo el análisis global
    if "TIPO DOC" in df.columns:
        df = df[df["TIPO DOC"].astype(str).str.upper().str.contains("FACTURA")]
        
    return df
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
