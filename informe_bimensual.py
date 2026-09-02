import os
import io
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

def generar_grafica_comparativa(labels, valores_cliente, valores_proveedor):
    """Genera una gráfica de barras comparativa en memoria para el Word."""
    fig, ax = plt.subplots(figsize=(6.5, 3.8))
    x = range(len(labels))
    width = 0.35

    ax.bar([p - width/2 for p in x], valores_cliente, width, label='Clientes Seleccionados', color='#1f77b4')
    ax.bar([p + width/2 for p in x], valores_proveedor, width, label='Ventas de Proveedores (FABRICANTE)', color='#ff7f0e')

    ax.set_ylabel('Ventas ($)')
    ax.set_title('Comparativo Bimensual de Ventas', fontsize=12, fontweight='bold')
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, rotation=10)
    ax.legend()
    plt.tight_layout()

    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', dpi=200)
    img_buf.seek(0)
    plt.close(fig)
    return img_buf


def render_modulo_informe(df):
    st.header("📄 Generador de Informe Bimensual (Word)")

    # Ruta a tu plantilla en la carpeta 'templates'
    RUTA_PLANTILLA = os.path.join("templates", "BIMENSUAL MAY-JUN.docx")

    if not os.path.exists(RUTA_PLANTILLA):
        st.error(f"❌ No se encontró la plantilla en `{RUTA_PLANTILLA}`. Revisa que el archivo exista en la carpeta templates.")
        return

    st.subheader("1. Selección de Filtros (Cliente o Asesor)")
    modo_seleccion = st.radio("Criterio de selección:", ["Por VENDEDOR (Asesor)", "Selección Manual de Clientes"])

    if modo_seleccion == "Por VENDEDOR (Asesor)":
        vendedores = sorted(df["VENDEDOR"].dropna().unique().tolist()) if "VENDEDOR" in df.columns else []
        vendedor_sel = st.selectbox("Selecciona el Vendedor / Asesor:", vendedores)
        clientes_sel = df[df["VENDEDOR"] == vendedor_sel]["CLIENTE"].dropna().unique().tolist()
        st.info(f"Clientes asignados a {vendedor_sel}: {len(clientes_sel)}")
    else:
        vendedor_sel = "Selección Manual"
        todos_clientes = sorted(df["CLIENTE"].dropna().unique().tolist()) if "CLIENTE" in df.columns else []
        clientes_sel = st.multiselect("Selecciona uno o varios Clientes:", todos_clientes)

    st.subheader("2. Selección de Períodos")
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

    if st.button("🚀 Generar Informe en Word"):
        if not clientes_sel:
            st.warning("Debes seleccionar al menos un cliente o asesor con clientes asociados.")
            return

        meses_curr, prev_bim_nombre, meses_prev = opciones_bimestre[bimestre_nombre]

        df_copy = df.copy()

        # --- FILTRO CLAVE: Solo tomar documentos de tipo FACTURA ---
        if "TIPO DOC" in df_copy.columns:
            df_copy = df_copy[df_copy["TIPO DOC"].astype(str).str.upper().str.contains("FACTURA")]

        if "FECHA" in df_copy.columns:
            df_copy["FECHA"] = pd.to_datetime(df_copy["FECHA"], errors='coerce')
            df_copy["ANIO"] = df_copy["FECHA"].dt.year
            df_copy["MES"] = df_copy["FECHA"].dt.month

        # Columna de valor numérico
        col_valor = "VALOR_VENTA" if "VALOR_VENTA" in df_copy.columns else "VALOR"
        df_copy[col_valor] = pd.to_numeric(df_copy[col_valor], errors='coerce').fillna(0)

        # Filtrar facturas de los clientes seleccionados
        df_clientes = df_copy[df_copy["CLIENTE"].isin(clientes_sel)]

        # --- Período 1: Bimestre Actual (ej: Jul-Ago 2026) ---
        p1 = df_clientes[(df_clientes["ANIO"] == anio_actual) & (df_clientes["MES"].isin(meses_curr))]
        v_p1_cli = p1[col_valor].sum()

        # --- Período 2: Mismo bimestre año anterior (ej: Jul-Ago 2025) ---
        p2 = df_clientes[(df_clientes["ANIO"] == (anio_actual - 1)) & (df_clientes["MES"].isin(meses_curr))]
        v_p2_cli = p2[col_valor].sum()

        # --- Período 3: Bimestre anterior mismo año (ej: May-Jun 2026) ---
        anio_prev_bim = anio_actual if bimestre_nombre != "Enero - Febrero" else anio_actual - 1
        p3 = df_clientes[(df_clientes["ANIO"] == anio_prev_bim) & (df_clientes["MES"].isin(meses_prev))]
        v_p3_cli = p3[col_valor].sum()

        # Totales Proveedores
        v_p1_prov = p1[col_valor].sum()
        v_p2_prov = p2[col_valor].sum()
        v_p3_prov = p3[col_valor].sum()

        labels = [
            f"{bimestre_nombre}\n{anio_actual}",
            f"{bimestre_nombre}\n{anio_actual-1}",
            f"{prev_bim_nombre}\n{anio_prev_bim}"
        ]
        valores_cli = [v_p1_cli, v_p2_cli, v_p3_cli]
        valores_prov = [v_p1_prov, v_p2_prov, v_p3_prov]

        grafica_buf = generar_grafica_comparativa(labels, valores_cli, valores_prov)

        # Cargar plantilla Word e inyectar valores
        doc = DocxTemplate(RUTA_PLANTILLA)

        contexto = {
            "vendedor": vendedor_sel,
            "bimestre": bimestre_nombre,
            "anio": str(anio_actual),
            "total_cliente_p1": f"${v_p1_cli:,.2f}",
            "total_cliente_p2": f"${v_p2_cli:,.2f}",
            "total_cliente_p3": f"${v_p3_cli:,.2f}",
            "grafica_ventas": InlineImage(doc, grafica_buf, width=Inches(5.8))
        }

        doc.render(contexto)

        out_buffer = io.BytesIO()
        doc.save(out_buffer)
        out_buffer.seek(0)

        st.success("✅ ¡Informe generado con éxito!")
        st.download_button(
            label="📥 Descargar Informe Word Editado (.docx)",
            data=out_buffer,
            file_name=f"Informe_{vendedor_sel}_{bimestre_nombre}_{anio_actual}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
