import os
import io
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

# Diccionario oficial de Bimestres (B1 a B6)
MAPA_BIMESTRES = {
    "B1": {"nombre": "B1 (Ene - Feb)", "meses": [1, 2], "previo": "B6", "meses_prev": [11, 12]},
    "B2": {"nombre": "B2 (Mar - Abr)", "meses": [3, 4], "previo": "B1", "meses_prev": [1, 2]},
    "B3": {"nombre": "B3 (May - Jun)", "meses": [5, 6], "previo": "B2", "meses_prev": [3, 4]},
    "B4": {"nombre": "B4 (Jul - Ago)", "meses": [7, 8], "previo": "B3", "meses_prev": [5, 6]},
    "B5": {"nombre": "B5 (Sep - Oct)", "meses": [9, 10], "previo": "B4", "meses_prev": [7, 8]},
    "B6": {"nombre": "B6 (Nov - Dic)", "meses": [11, 12], "previo": "B5", "meses_prev": [9, 10]},
}

def generar_grafica_comparativa(labels, valores_totales):
    """Genera una gráfica de barras de los totales acumulados para el informe."""
    fig, ax = plt.subplots(figsize=(6.5, 3.5))
    colores = ['#1f77b4', '#ff7f0e', '#2ca02c']
    
    bars = ax.bar(labels, valores_totales, color=colores, width=0.5)
    ax.set_ylabel('Ventas ($)', fontsize=10)
    ax.set_title('Comparativo de Ventas Totales por Período', fontsize=12, fontweight='bold')
    
    # Formato de moneda en las barras
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'${height:,.0f}',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),  
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=9, fontweight='bold')

    plt.tight_layout()

    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', dpi=200)
    img_buf.seek(0)
    plt.close(fig)
    return img_buf


def redactar_analisis(v_actual, v_anio_ant, v_bim_ant, nom_b_actual, nom_b_prev, anio_actual):
    """Genera una redacción ejecutiva basada en los resultados numéricos."""
    var_vs_anio = ((v_actual - v_anio_ant) / v_anio_ant * 100) if v_anio_ant > 0 else 0
    var_vs_bim = ((v_actual - v_bim_ant) / v_bim_ant * 100) if v_bim_ant > 0 else 0

    tend_anio = "un crecimiento" if var_vs_anio >= 0 else "un decrecimiento"
    tend_bim = "un incremento" if var_vs_bim >= 0 else "una disminución"

    texto = (
        f"Durante el período {nom_b_actual} de {anio_actual}, se alcanzaron ventas totales de ${v_actual:,.2f}. "
        f"Al comparar este resultado con el mismo período del año anterior ({nom_b_actual} {anio_actual-1}), "
        f"se evidencia {tend_anio} del {abs(var_vs_anio):.2f}% (frente a ${v_anio_ant:,.2f}). "
        f"Por otra parte, en relación con el bimestre inmediatamente anterior ({nom_b_prev}), el comportamiento registró "
        f"{tend_bim} del {abs(var_vs_bim):.2f}% respecto a los ${v_bim_ant:,.2f} facturados previamente."
    )
    return texto


def render_modulo_informe(df):
    st.header("📄 Generador de Informe Bimensual (Word)")

    RUTA_PLANTILLA = os.path.join("templates", "BIMENSUAL MAY-JUN.docx")

    if not os.path.exists(RUTA_PLANTILLA):
        st.error(f"❌ No se encontró la plantilla en `{RUTA_PLANTILLA}`.")
        return

    col_a, col_b = st.columns(2)
    
    with col_a:
        st.subheader("1. Selección de Clientes")
        modo_cli = st.radio("Método para clientes:", ["Excel de Clientes", "Por VENDEDOR / Asesor", "Selección Manual"])
        
        clientes_sel = []
        vendedor_nom = "Varios / Lista Propia"

        if modo_cli == "Excel de Clientes":
            file_cli = st.file_uploader("Subir listado de Clientes (Excel)", type=["xlsx", "xls"], key="inf_cli")
            if file_cli:
                df_c = pd.read_excel(file_cli)
                col_c = 'NOMBRE SN' if 'NOMBRE SN' in df_c.columns else df_c.columns[0]
                clientes_sel = df_c[col_c].dropna().astype(str).str.strip().unique().tolist()
                st.success(f"Cargados {len(clientes_sel)} clientes.")
        elif modo_cli == "Por VENDEDOR / Asesor":
            if "VENDEDOR" in df.columns:
                vendedores = sorted(df["VENDEDOR"].dropna().unique().tolist())
                vendedor_nom = st.selectbox("Selecciona Vendedor:", vendedores)
                clientes_sel = df[df["VENDEDOR"] == vendedor_nom]["CLIENTE"].dropna().unique().tolist()
                st.info(f"Clientes de {vendedor_nom}: {len(clientes_sel)}")
        else:
            todos_c = sorted(df["CLIENTE"].dropna().unique().tolist()) if "CLIENTE" in df.columns else []
            clientes_sel = st.multiselect("Clientes:", todos_c)

    with col_b:
        st.subheader("2. Selección de Fabricantes / Proveedores")
        file_prov = st.file_uploader("Subir listado de Fabricantes (Excel)", type=["xlsx", "xls"], key="inf_prov")
        proveedores_sel = []
        if file_prov:
            df_p = pd.read_excel(file_prov)
            col_p = df_p.columns[0]
            for col in df_p.columns:
                if 'FABRICANTE' in col.upper() or 'PROVEEDOR' in col.upper() or 'MARCA' in col.upper():
                    col_p = col
                    break
            proveedores_sel = df_p[col_p].dropna().astype(str).str.strip().unique().tolist()
            st.success(f"Cargados {len(proveedores_sel)} fabricantes.")
        else:
            st.caption("Si no subes archivo, se evaluarán todos los proveedores del catálogo.")

    st.markdown("---")
    st.subheader("3. Período a Reportar")
    c1, c2 = st.columns(2)
    with c1:
        codigo_bimestre = st.selectbox("Seleccione Bimestre Actual:", list(MAPA_BIMESTRES.keys()), 
                                       format_func=lambda x: MAPA_BIMESTRES[x]["nombre"])
    with c2:
        anio_actual = st.number_input("Año del Informe", min_value=2020, max_value=2030, value=2026)

    if st.button("🚀 Generar Informe Word Bimensual"):
        if not clientes_sel:
            st.warning("⚠️ Debes definir al menos un cliente o subir la lista de clientes.")
            return

        # 1. Copia y filtro estricto: SOLO FACTURAS
        df_work = df.copy()
        if "TIPO DOC" in df_work.columns:
            df_work = df_work[df_work["TIPO DOC"].astype(str).str.upper().str.contains("FACTURA")]

        # 2. Fechas y formatos
        if "FECHA" in df_work.columns:
            df_work["FECHA"] = pd.to_datetime(df_work["FECHA"], errors='coerce')
            df_work["ANIO"] = df_work["FECHA"].dt.year
            df_work["MES"] = df_work["FECHA"].dt.month

        col_val = "VALOR_VENTA" if "VALOR_VENTA" in df_work.columns else "VALOR"
        df_work[col_val] = pd.to_numeric(df_work[col_val], errors='coerce').fillna(0)

        # Columna de proveedor/fabricante
        col_prov = None
        for cp in ["FABRICANTE", "PROVEEDOR", "MARCA"]:
            if cp in df_work.columns:
                col_prov = cp
                break
        if not col_prov:
            col_prov = "FABRICANTE"
            df_work[col_prov] = "GENERAL"

        # Filtrar por clientes elegidos
        df_work = df_work[df_work["CLIENTE"].isin(clientes_sel)]

        # Filtrar por proveedores elegidos si se subió lista
        if proveedores_sel:
            df_work = df_work[df_work[col_prov].isin(proveedores_sel)]

        # Definir configuraciones de períodos
        cfg_b = MAPA_BIMESTRES[codigo_bimestre]
        meses_act = cfg_b["meses"]
        
        cod_prev = cfg_b["previo"]
        cfg_prev = MAPA_BIMESTRES[cod_prev]
        meses_prev = cfg_b["meses_prev"]
        
        anio_prev_bim = anio_actual if codigo_bimestre != "B1" else (anio_actual - 1)

        # Filtros de DataFrames por periodo
        df_p1 = df_work[(df_work["ANIO"] == anio_actual) & (df_work["MES"].isin(meses_act))] # B actual 2026
        df_p2 = df_work[(df_work["ANIO"] == (anio_actual - 1)) & (df_work["MES"].isin(meses_act))] # B actual 2025
        df_p3 = df_work[(df_work["ANIO"] == anio_prev_bim) & (df_work["MES"].isin(meses_prev))] # B previo

        # Nombrado de Encabezados de Tabla
        head_b_act = f"VENTA {codigo_bimestre} {anio_actual}"
        head_b_ant_anio = f"VENTA {codigo_bimestre} {anio_actual - 1}"
        head_b_prev = f"VENTA {cod_prev} {anio_prev_bim}"

        # --- TABLA 1: PROVEEDORES ---
        prov_p1 = df_p1.groupby(col_prov)[col_val].sum()
        prov_p2 = df_p2.groupby(col_prov)[col_val].sum()
        prov_p3 = df_p3.groupby(col_prov)[col_val].sum()

        todos_provs = sorted(list(set(prov_p1.index).union(set(prov_p2.index)).union(set(prov_p3.index))))
        
        tabla_proveedores = []
        for pr in todos_provs:
            tabla_proveedores.append({
                "PROVEEDOR": pr,
                "v_act": f"${prov_p1.get(pr, 0):,.2f}",
                "v_p2": f"${prov_p2.get(pr, 0):,.2f}",
                "v_p3": f"${prov_p3.get(pr, 0):,.2f}"
            })

        # --- TABLA 2: CLIENTES ---
        cli_p1 = df_p1.groupby("CLIENTE")[col_val].sum()
        cli_p2 = df_p2.groupby("CLIENTE")[col_val].sum()
        cli_p3 = df_p3.groupby("CLIENTE")[col_val].sum()

        todos_clis = sorted(list(set(cli_p1.index).union(set(cli_p2.index)).union(set(cli_p3.index))))

        tabla_clientes = []
        for cl in todos_clis:
            tabla_clientes.append({
                "CLIENTE": cl,
                "v_act": f"${cli_p1.get(cl, 0):,.2f}",
                "v_p2": f"${cli_p2.get(cl, 0):,.2f}",
                "v_p3": f"${cli_p3.get(cl, 0):,.2f}"
            })

        # Totales Consolidados
        tot_act = df_p1[col_val].sum()
        tot_p2 = df_p2[col_val].sum()
        tot_p3 = df_p3[col_val].sum()

        # Gráfica
        labels_graph = [f"{codigo_bimestre} {anio_actual}", f"{codigo_bimestre} {anio_actual-1}", f"{cod_prev} {anio_prev_bim}"]
        buf_grafica = generar_grafica_comparativa(labels_graph, [tot_act, tot_p2, tot_p3])

        # Redacción del análisis sintético
        texto_analisis = redactar_analisis(
            tot_act, tot_p2, tot_p3,
            cfg_b["nombre"], cfg_prev["nombre"], anio_actual
        )

        # Inyección en plantilla Word
        doc = DocxTemplate(RUTA_PLANTILLA)

        contexto = {
            "vendedor": vendedor_nom,
            "bimestre": cfg_b["nombre"],
            "anio": str(anio_actual),
            "col_b_act": head_b_act,
            "col_b_ant": head_b_ant_anio,
            "col_b_prev": head_b_prev,
            "tabla_proveedores": tabla_proveedores,
            "tabla_clientes": tabla_clientes,
            "total_b_act": f"${tot_act:,.2f}",
            "total_b_ant": f"${tot_p2:,.2f}",
            "total_b_prev": f"${tot_p3:,.2f}",
            "analisis_texto": texto_analisis,
            "grafica_ventas": InlineImage(doc, buf_grafica, width=Inches(5.8))
        }

        doc.render(contexto)

        out = io.BytesIO()
        doc.save(out)
        out.seek(0)

        st.success("✅ ¡Informe generado con tablas de Fabricantes, Clientes y Análisis automático!")
        st.download_button(
            label="📥 Descargar Informe Word (.docx)",
            data=out,
            file_name=f"Informe_{codigo_bimestre}_{anio_actual}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
