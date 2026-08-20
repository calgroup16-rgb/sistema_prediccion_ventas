import streamlit as st
import pandas as pd
from procesamiento import (
    cargar_excel,
    preparar_anio,
    preparar_valor
)
from analisis_clientes import (
    clientes_inactivos_2026,
    productos_clientes_inactivos
)
from recomendador import (
    recomendar_productos
)
from alimentacion_animal import (
    ranking_clientes_animal,
    productos_animal_por_cliente
)


st.set_page_config(
    page_title="Sistema de Predicción Comercial",
    page_icon="📊",
    layout="wide"
)


st.title("📊 Sistema de Predicción Comercial")

st.write(
    "Analice clientes, productos, ventas y oportunidades comerciales."
)


archivo = st.file_uploader(
    "Suba su archivo Excel",
    type=["xlsx", "xls"]
)


if archivo is not None:

    try:

        df = cargar_excel(archivo)
        df = preparar_anio(df)
        df = preparar_valor(df)

        st.success(
            f"Archivo cargado correctamente: "
            f"{len(df):,} registros"
        )

        st.subheader("Vista previa")

        st.dataframe(
            df.head(20),
            use_container_width=True
        )

        # =========================================
        # CLIENTES INACTIVOS
        # =========================================

        st.header(
            "🔴 Clientes que no compraron en 2026"
        )

        top20 = clientes_inactivos_2026(df)

        st.dataframe(
            top20,
            use_container_width=True
        )

        csv_top20 = top20.to_csv(
            index=False
        ).encode("utf-8")

        st.download_button(
            "⬇️ Descargar Top 20 inactivos",
            csv_top20,
            "clientes_inactivos_2026.csv",
            "text/csv"
        )

        # =========================================
        # PRODUCTOS DE LOS INACTIVOS
        # =========================================

        st.header(
            "🛒 Productos comprados históricamente"
        )

        clientes = top20["NOMBRE SN"].tolist()

        productos = productos_clientes_inactivos(
            df,
            clientes
        )

        st.dataframe(
            productos,
            use_container_width=True
        )

        # =========================================
        # CLIENTES OBJETIVO
        # =========================================

        st.header(
            "🎯 Clientes objetivo"
        )

        clientes_disponibles = sorted(
            df["NOMBRE SN"].unique()
        )

        seleccionados = st.multiselect(
            "Seleccione los clientes que desea atacar",
            clientes_disponibles
        )

        if seleccionados:

            recomendaciones = recomendar_productos(
                df,
                seleccionados
            )

            st.subheader(
                "🤖 Productos recomendados"
            )

            st.dataframe(
                recomendaciones,
                use_container_width=True
            )

            csv = recomendaciones.to_csv(
                index=False
            ).encode("utf-8")

            st.download_button(
                "⬇️ Descargar recomendaciones",
                csv,
                "recomendaciones.csv",
                "text/csv"
            )

        # =========================================
        # ALIMENTACION ANIMAL
        # =========================================

        st.header(
            "🐄 Oportunidad: Alimentación Animal"
        )

        ranking = ranking_clientes_animal(df)

        st.subheader(
            "Top 10 clientes con mayor potencial"
        )

        st.dataframe(
            ranking,
            use_container_width=True
        )

        clientes_animal = ranking[
            "NOMBRE SN"
        ].tolist()

        productos_animal = (
            productos_animal_por_cliente(
                df,
                clientes_animal
            )
        )

        st.subheader(
            "Productos relacionados"
        )

        st.dataframe(
            productos_animal,
            use_container_width=True
        )

        csv_animal = ranking.to_csv(
            index=False
        ).encode("utf-8")

        st.download_button(
            "⬇️ Descargar oportunidades alimentación animal",
            csv_animal,
            "oportunidades_alimentacion_animal.csv",
            "text/csv"
        )

    except Exception as e:

        st.error(
            f"Se produjo un error: {e}"
        )
