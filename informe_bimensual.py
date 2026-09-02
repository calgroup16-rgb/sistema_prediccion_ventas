import os
import io
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

# Filtrar únicamente los registros que correspondan a Facturas
        if "TIPO DOC" in df_copy.columns:
            df_copy = df_copy[df_copy["TIPO DOC"].astype(str).str.upper().str.contains("FACTURA")]

def generar_grafica_comparativa(labels, valores_cliente):
    """Genera una gráfica de barras para la comparación bimensual."""
    fig, ax = plt.subplots(figsize=(6, 3.5))
    
    ax.bar(labels, valores_cliente, color=['#1f77b4', '#aec7e8', '#ff7f0e'], width=0.5)
    ax.set_ylabel('Ventas ($)')
    ax.set_title('Comparación de Ventas por Período')
    plt.xticks(rotation=15)
    plt.tight_layout()

    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', dpi=200)
    img_buf.seek(0)
    plt.close(fig)
    return img_buf


def render_modulo_informe(df):
    st.header("📄 Generación de Informe Bimensual")

    RUTA_PLANTILLA = os.path.join("templates", "BIMENSUAL MAY-JUN.docx")

    if not os.path.exists(RUTA_PLANTILLA):
        st.error(f"No se encontró la plantilla en `{RUTA_PLANTILLA}`. Revisa que subiste el archivo a la carpeta 'templates'.")
        return

    # 1. Filtros principales
    modo_seleccion = st.radio("Método de Selección:", ["Por VENDEDOR (Asesor)", "Lista Manual de Clientes"])

    if modo_seleccion == "Por VENDEDOR (Asesor)":
        vendedores = sorted(df["VENDEDOR"].dropna().unique().tolist()) if "VENDEDOR" in df.columns else []
        vendedor_sel = st.selectbox("Selecciona el Asesor:", vendedores)
        clientes_sel = df[df["VENDEDOR"] == vendedor_sel]["CLIENTE"].dropna().unique().tolist()
    else:
        vendedor_sel = "Selección Manual"
        todos_clientes = sorted(df["CLIENTE"].dropna().unique().tolist())
        clientes_sel = st.multiselect("Selecciona los Clientes:", todos_clientes)

    # 2. Selección de Períodos
    col1, col2 = st.columns(2)
    with col1:
        anio_actual = st.number_input("Año del Informe", min_value=2020, max_value=2030, value=2026)
    with col2:
        opciones_bimestre = {
            "Enero - Febrero": ([1, 2], "Noviembre - Diciembre", [11, 12]),
            "Marzo - Abril": ([3, 4], "Enero - Febrero", [1, 2]),
            "Mayo - Junio": ([5, 6], "Marzo - Abril", [3, 4]),
            "Julio - Agosto": ([7, 8], "Mayo - Junio", [5, 6]),
            "Septiembre - Octubre": ([9, 10], "Julio - Agosto", [7, 8]),
            "Noviembre - Diciembre": ([11, 12], "Septiembre - Octubre", [9, 10])
        }
        bimestre_nombre = st.selectbox("Selecciona el Bimestre Actual:", list(opciones_bimestre.keys()))

    if st.button("🚀 Generar Informe Word"):
        if not clientes_sel:
            st.warning("Debes seleccionar al menos un cliente o asesor.")
            return

        meses_curr, prev_bim_nombre, meses_prev = opciones_bimestre[bimestre_nombre]

        df_copy = df.copy()
        if "FECHA" in df_copy.columns:
            df_copy["FECHA"] = pd.to_datetime(df_copy["FECHA"], errors='coerce')
            df_copy["ANIO"] = df_copy["FECHA"].dt.year
            df_copy["MES"] = df_copy["FECHA"].dt.month

        # Filtrar datos de la lista de clientes seleccionados
        df_clientes = df_copy[df_copy["CLIENTE"].isin(clientes_sel)]

        # --- Cálculos Período 1: Actual ---
        p1 = df_clientes[(df_clientes["ANIO"] == anio_actual) & (df_clientes["MES"].isin(meses_curr))]
        v_p1 = p1["VALOR_VENTA"].sum()

        # --- Cálculos Período 2: Mismo bimestre año anterior ---
        p2 = df_clientes[(df_clientes["ANIO"] == (anio_actual - 1)) & (df_clientes["MES"].isin(meses_curr))]
        v_p2 = p2["VALOR_VENTA"].sum()

        # --- Cálculos Período 3: Bimestre anterior del mismo año ---
        anio_prev_bim = anio_actual if bimestre_nombre != "Enero - Febrero" else anio_actual - 1
        p3 = df_clientes[(df_clientes["ANIO"] == anio_prev_bim) & (df_clientes["MES"].isin(meses_prev))]
        v_p3 = p3["VALOR_VENTA"].sum()

        # Generar gráfica comparativa
        labels = [f"{bimestre_nombre} {anio_actual}", f"{bimestre_nombre} {anio_actual-1}", f"{prev_bim_nombre} {anio_prev_bim}"]
        valores = [v_p1, v_p2, v_p3]
        grafica_buf = generar_grafica_comparativa(labels, valores)

        # Cargar la plantilla Word y rellenar los valores
        doc = DocxTemplate(RUTA_PLANTILLA)

        contexto = {
            "vendedor": vendedor_sel,
            "bimestre": bimestre_nombre,
            "anio": str(anio_actual),
            "total_p1": f"${v_p1:,.2f}",
            "total_p2": f"${v_p2:,.2f}",
            "total_p3": f"${v_p3:,.2f}",
            "grafica_ventas": InlineImage(doc, grafica_buf, width=Inches(5.5))
        }

        doc.render(contexto)

        out_buffer = io.BytesIO()
        doc.save(out_buffer)
        out_buffer.seek(0)

        st.success("¡Informe generado correctamente!")
        st.download_button(
            label="📥 Descargar Informe Editado (.docx)",
            data=out_buffer,
            file_name=f"Informe_{vendedor_sel}_{bimestre_nombre}_{anio_actual}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
