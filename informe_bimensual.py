import os
import io
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- MAPA REGLA FIJA DE BIMESTRES ---
MAPA_BIMESTRES = {
    "B1": {"nombre": "B1 (Ene - Feb)", "texto_titulo": "Enero - Febrero", "meses": [1, 2], "previo": "B6", "meses_prev": [11, 12]},
    "B2": {"nombre": "B2 (Mar - Abr)", "texto_titulo": "Marzo - Abril", "meses": [3, 4], "previo": "B1", "meses_prev": [1, 2]},
    "B3": {"nombre": "B3 (May - Jun)", "texto_titulo": "Mayo - Junio", "meses": [5, 6], "previo": "B2", "meses_prev": [3, 4]},
    "B4": {"nombre": "B4 (Jul - Ago)", "texto_titulo": "Julio - Agosto", "meses": [7, 8], "previo": "B3", "meses_prev": [5, 6]},
    "B5": {"nombre": "B5 (Sep - Oct)", "texto_titulo": "Septiembre - Octubre", "meses": [9, 10], "previo": "B4", "meses_prev": [7, 8]},
    "B6": {"nombre": "B6 (Nov - Dic)", "texto_titulo": "Noviembre - Diciembre", "meses": [11, 12], "previo": "B5", "meses_prev": [9, 10]},
}

def preparar_dataframe(df):
    if df is None or df.empty:
        return df

    df = df.copy()
    
    # 1. Normalización de encabezados
    df.columns = [str(col).strip() for col in df.columns]

    # Identificar columnas fijas AÑO y MES
    col_ano = next((c for c in df.columns if str(c).strip().upper() in ["AÑO", "ANO", "ANIO", "YEAR"]), None)
    col_mes = next((c for c in df.columns if str(c).strip().upper() in ["MES", "MONTH"]), None)

    if not col_ano or not col_mes:
        col_fecha = next((c for c in df.columns if "FECHA" in str(c).strip().upper()), None)
        if col_fecha:
            df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
            if not col_mes:
                df["MES"] = df[col_fecha].dt.month
            if not col_ano:
                df["AÑO"] = df[col_fecha].dt.year

    if col_ano and col_ano != "AÑO":
        df["AÑO"] = df[col_ano]
    if col_mes and col_mes != "MES":
        df["MES"] = df[col_mes]

    # Forzar AÑO y MES a enteros (Limpia valores decimales tipo 2025.0)
    df["AÑO"] = pd.to_numeric(df["AÑO"], errors="coerce").fillna(0).astype(int)
    df["MES"] = pd.to_numeric(df["MES"], errors="coerce").fillna(0).astype(int)

    return df

def aplicar_neteo_total_linea(df):
    """
    Regla de Neteo Obligatoria (Filtro Global)
    Si TIPO DOC es NOTA CREDITO, resta a TOTAL LINEA.
    """
    if df is None or df.empty:
        return df, "TOTAL LINEA"

    df_res = df.copy()

    col_val = next((c for c in df_res.columns if str(c).strip().upper() == "TOTAL LINEA"), None)
    if not col_val:
        col_val = next((c for c in df_res.columns if "TOTAL" in str(c).upper() and "LINEA" in str(c).upper()), None)
    if not col_val:
        col_val = next((c for c in df_res.columns if "TOTAL" in str(c).upper()), df_res.columns[-1])

    df_res[col_val] = pd.to_numeric(df_res[col_val], errors='coerce').fillna(0.0)

    col_tipo_doc = next((c for c in df_res.columns if str(c).strip().upper() in ["TIPO DOC", "TIPO_DOC", "TIPO DOCUMENTO", "DOCUMENTO"]), None)

    if col_tipo_doc:
        # Detecta "NOTA CREDITO", "NOTA", "NC", "CREDITO"
        es_nc = df_res[col_tipo_doc].astype(str).str.strip().str.upper().str.contains("NOTA CREDITO|NOTA|NC|CREDITO|CRÉDITO", na=False)
        df_res.loc[es_nc & (df_res[col_val] > 0), col_val] = -1 * df_res.loc[es_nc & (df_res[col_val] > 0), col_val]

    return df_res, col_val

def generar_grafica_comparativa_fabricantes(labels, valores_totales):
    fig, ax = plt.subplots(figsize=(6.5, 3.5))
    colores = ['#1f77b4', '#ff7f0e', '#2ca02c']
    
    bars = ax.bar(labels, valores_totales, color=colores, width=0.45)
    ax.set_ylabel('Ventas ($)', fontsize=10, fontfamily='sans-serif')
    ax.set_title('Gráfica 1. Comparativo Bimensual de Ventas Totales', fontsize=11, fontweight='bold', fontfamily='sans-serif')
    
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'${height:,.0f}',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 4),  
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=9, fontweight='bold', fontfamily='sans-serif')

    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', dpi=200)
    img_buf.seek(0)
    plt.close(fig)
    return img_buf

def generar_analisis_y_aspectos(v_act, v_p2, v_p3, df_p1, df_p3, nom_b_act, nom_b_prev, anio_actual, col_val, col_cliente):
    var_vs_anio = ((v_act - v_p2) / v_p2 * 100) if v_p2 > 0 else 0
    var_vs_bim = ((v_act - v_p3) / v_p3 * 100) if v_p3 > 0 else 0

    tend_anio = "un crecimiento" if var_vs_anio >= 0 else "un decrecimiento"
    tend_bim = "un incremento" if var_vs_bim >= 0 else "una disminución"

    col_desc = next((c for c in df_p1.columns if str(c).strip().upper() in ["DESCRIPCION", "DESCRIPCIÓN", "PRODUCTO", "CONCEPTO", "LINEA"]), None)

    prod_subieron_str, prod_bajaron_str, texto_insumos = "", "", ""

    if col_desc and col_val in df_p1.columns:
        v_p1_prod = df_p1.groupby(col_desc)[col_val].sum() if not df_p1.empty else pd.Series(dtype=float)
        v_p3_prod = df_p3.groupby(col_desc)[col_val].sum() if not df_p3.empty else pd.Series(dtype=float)

        todos_prods = list(set(v_p1_prod.index).union(set(v_p3_prod.index)))
        diff_list = []
        for p in todos_prods:
            v_actual = float(v_p1_prod.get(p, 0.0))
            v_previo = float(v_p3_prod.get(p, 0.0))
            diff = v_actual - v_previo
            pct = ((diff / v_previo) * 100) if v_previo > 0 else (100.0 if v_actual > 0 else 0.0)
            diff_list.append({"producto": str(p).strip(), "diff": diff, "pct": pct})

        df_diff = pd.DataFrame(diff_list)

        if not df_diff.empty:
            subieron = df_diff[df_diff["diff"] > 0].sort_values(by="diff", ascending=False).head(3)
            bajaron = df_diff[df_diff["diff"] < 0].sort_values(by="diff", ascending=True).head(3)

            if not subieron.empty:
                prod_subieron_str = ", ".join([f"'{row['producto']}' (+${row['diff']:,.0f})" for _, row in subieron.iterrows()])
            if not bajaron.empty:
                prod_bajaron_str = ", ".join([f"'{row['producto']}' (-${abs(row['diff']):,.0f})" for _, row in bajaron.iterrows()])

            if prod_subieron_str or prod_bajaron_str:
                texto_insumos = " Destacando en el portafolio: "
                if prod_subieron_str:
                    texto_insumos += f"crecimiento en {prod_subieron_str}."
                if prod_bajaron_str:
                    texto_insumos += f" Reducción en {prod_bajaron_str}."

    top_clis = df_p1.groupby(col_cliente)[col_val].sum().sort_values(ascending=False).index.tolist() if col_cliente and col_cliente in df_p1.columns and not df_p1.empty else []
    cli_1 = str(top_clis[0]) if len(top_clis) > 0 else "clientes clave"
    cli_2 = str(top_clis[1]) if len(top_clis) > 1 else "segundo grupo de cuentas"

    texto_analisis = (
        f"Durante el período {nom_b_act} de {anio_actual}, se alcanzaron ventas totales netas de ${v_act:,.2f}. "
        f"Al comparar este resultado con el mismo período del año anterior ({nom_b_act} {anio_actual-1}), "
        f"se evidencia {tend_anio} del {abs(var_vs_anio):.2f}% (frente a ${v_p2:,.2f}). "
        f"Asimismo, en relación con el bimestre inmediatamente anterior ({nom_b_prev}), se registró "
        f"{tend_bim} del {abs(var_vs_bim):.2f}% respecto a los ${v_p3:,.2f} facturados."
        f"{texto_insumos}"
    )

    aspectos = [
        {"titulo": "Análisis de líneas e insumos con mayor variación", "descripcion": "Impulsar estrategias para revertir caídas en productos afectados y mantener abastecimiento constante."},
        {"titulo": "Recuperación de negocios pendientes", "descripcion": f"Fortalecer seguimiento a clientes estratégicos como {cli_1} y {cli_2}."},
        {"titulo": "Disponibilidad de inventario", "descripcion": "Garantizar disponibilidad continua en insumos de alta rotación."},
        {"titulo": "Líneas de mayor rentabilidad", "descripcion": "Aumentar participación comercial de las líneas con mayor margen."},
        {"titulo": "Venta cruzada y cartera", "descripcion": "Aprovechar cuentas activas para introducir nuevos productos y acelerar recaudo."}
    ]

    texto_cierre = "Se mantiene el enfoque en potenciar la venta de las líneas principales y asegurar el cumplimiento de metas comerciales."

    return texto_analisis, aspectos, texto_cierre

def dar_formato_tabla(table, headers, rows_data):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9.5)
            
        shading = OxmlElement('w:shd')
        shading.set(qn('w:val'), 'clear')
        shading.set(qn('w:color'), 'auto')
        shading.set(qn('w:fill'), '1F4E78')
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading)

    tot_rows = len(rows_data)
    for r_idx, row in enumerate(rows_data):
        row_cells = table.add_row().cells
        es_ultima_fila = (r_idx == tot_rows - 1)
        
        for c_idx, val in enumerate(row):
            row_cells[c_idx].text = str(val)
            p = row_cells[c_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.RIGHT
            for run in p.runs:
                run.font.size = Pt(9)
                if es_ultima_fila:
                    run.font.bold = True

            if es_ultima_fila:
                shading = OxmlElement('w:shd')
                shading.set(qn('w:val'), 'clear')
                shading.set(qn('w:color'), 'auto')
                shading.set(qn('w:fill'), 'D9E1F2')
                row_cells[c_idx]._tc.get_or_add_tcPr().append(shading)
            elif r_idx % 2 == 1:
                shading = OxmlElement('w:shd')
                shading.set(qn('w:val'), 'clear')
                shading.set(qn('w:color'), 'auto')
                shading.set(qn('w:fill'), 'F2F2F2')
                row_cells[c_idx]._tc.get_or_add_tcPr().append(shading)

def construir_documento_word(contexto, path_template=None):
    doc = Document(path_template) if path_template and os.path.exists(path_template) else Document()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(contexto["titulo_informe"])
    r_title.font.bold = True
    r_title.font.size = Pt(16)
    r_title.font.color.rgb = RGBColor(31, 78, 120)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"Responsable Comercial: {contexto['vendedor']} | Año: {contexto['anio']}")
    r_sub.font.italic = True
    r_sub.font.size = Pt(11)

    doc.add_paragraph()

    # TABLA 1: FABRICANTES
    h1 = doc.add_paragraph()
    r_h1 = h1.add_run("Tabla 1. Resumen de Ventas por Línea / Fabricante")
    r_h1.font.bold = True
    r_h1.font.size = Pt(13)
    r_h1.font.color.rgb = RGBColor(31, 78, 120)

    headers_prov = ["Fabricante / Proveedor", contexto["col_b_act"], contexto["col_b_ant"], contexto["col_b_prev"]]
    rows_prov = [[item["PROVEEDOR"], item["v_act"], item["v_p2"], item["v_p3"]] for item in contexto["tabla_proveedores"]]
    t_prov = doc.add_table(rows=1, cols=4)
    dar_formato_tabla(t_prov, headers_prov, rows_prov)

    doc.add_paragraph()

    # GRÁFICA
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(contexto["grafica_ventas"], width=Inches(6.0))

    doc.add_paragraph()

    # TABLA 2: CLIENTES
    h2 = doc.add_paragraph()
    r_h2 = h2.add_run("Tabla 2. Resumen de Ventas por Cliente")
    r_h2.font.bold = True
    r_h2.font.size = Pt(13)
    r_h2.font.color.rgb = RGBColor(31, 78, 120)

    headers_cli = ["Cliente", contexto["col_b_act"], contexto["col_b_ant"], contexto["col_b_prev"]]
    rows_cli = [[item["CLIENTE"], item["v_act"], item["v_p2"], item["v_p3"]] for item in contexto["tabla_clientes"]]
    t_cli = doc.add_table(rows=1, cols=4)
    dar_formato_tabla(t_cli, headers_cli, rows_cli)

    doc.add_paragraph()

    # ANÁLISIS COMERCIAL
    h3 = doc.add_paragraph()
    r_h3 = h3.add_run("Análisis Comercial del Período")
    r_h3.font.bold = True
    r_h3.font.size = Pt(13)
    r_h3.font.color.rgb = RGBColor(31, 78, 120)

    p_ana = doc.add_paragraph(contexto["analisis_texto"])
    p_ana.paragraph_format.line_spacing = 1.15

    doc.add_paragraph()

    # ASPECTOS Y COMPROMISOS
    h4 = doc.add_paragraph()
    r_h4 = h4.add_run("Aspectos a Mejorar y Compromisos Comerciales")
    r_h4.font.bold = True
    r_h4.font.size = Pt(13)
    r_h4.font.color.rgb = RGBColor(31, 78, 120)

    for asp in contexto["aspectos_mejorar"]:
        p_asp = doc.add_paragraph()
        p_asp.paragraph_format.left_indent = Inches(0.2)
        p_asp.paragraph_format.line_spacing = 1.15
        r_tit = p_asp.add_run(f"• {asp['titulo']}: ")
        r_tit.font.bold = True
        p_asp.add_run(asp["descripcion"])

    p_cierre = doc.add_paragraph(contexto["texto_cierre_aspectos"])
    p_cierre.paragraph_format.line_spacing = 1.15

    return doc

def render_modulo_informe(df_global):
    st.title("📄 Generador de Informe Bimensual")

    if df_global is None or df_global.empty:
        st.warning("⚠️ Primero debe cargar los datos de ventas en la aplicación.")
        return

    df_global = preparar_dataframe(df_global)

    tab_gen, tab_clientes, tab_fabricantes = st.tabs(["🚀 Generar Informe", "📋 Cargar Lista Clientes", "🏭 Cargar Lista Fabricantes"])

    with tab_clientes:
        st.subheader("Cargar Archivo de Clientes Principales")
        file_cli = st.file_uploader("Cargue la lista de clientes (Excel/CSV):", type=["xlsx", "xls", "csv"], key="file_cli_inf")
        if file_cli:
            try:
                df_c = pd.read_excel(file_cli) if file_cli.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file_cli)
                st.session_state['df_clientes_custom'] = df_c
                st.success("✅ Lista de clientes cargada correctamente.")
                st.dataframe(df_c.head())
            except Exception as e:
                st.error(f"Error al leer archivo de clientes: {e}")

    with tab_fabricantes:
        st.subheader("Cargar Archivo de Fabricantes Principales")
        file_fab = st.file_uploader("Cargue la lista de fabricantes (Excel/CSV):", type=["xlsx", "xls", "csv"], key="file_fab_inf")
        if file_fab:
            try:
                df_f = pd.read_excel(file_fab) if file_fab.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file_fab)
                st.session_state['df_fabricantes_custom'] = df_f
                st.success("✅ Lista de fabricantes cargada correctamente.")
                st.dataframe(df_f.head())
            except Exception as e:
                st.error(f"Error al leer archivo de fabricantes: {e}")

    with tab_gen:
        col_vendedor = next((c for c in df_global.columns if str(c).strip().upper() in ["VENDEDOR", "RESPONSABLE"]), "VENDEDOR")
        
        # BÚSQUEDA ESPECÍFICA DE LA COLUMNA NOMBRE SN
        col_cliente = next((c for c in df_global.columns if str(c).strip().upper() in ["NOMBRE SN", "NOMBRE_SN", "NOMBRE  SN"]), None)
        if not col_cliente:
            col_cliente = next((c for c in df_global.columns if str(c).strip().upper() in ["CLIENTE", "NOMBRE_CLIENTE", "TERCERO", "RAZON_SOCIAL"]), None)

        col_prov = next((c for c in df_global.columns if str(c).strip().upper() in ["FABRICANTE", "PROVEEDOR", "MARCA"]), None)

        st.subheader("Configuración del Informe")
        col1, col2, col3 = st.columns(3)

        with col1:
            lista_vendedores = sorted(df_global[col_vendedor].dropna().unique()) if col_vendedor in df_global.columns else ["Todos"]
            vendedor_sel = st.selectbox("Seleccione Vendedor / Responsable:", lista_vendedores)
        with col2:
            bim_sel = st.selectbox("Seleccione el Bimestre del Informe:", list(MAPA_BIMESTRES.keys()), format_func=lambda x: MAPA_BIMESTRES[x]["nombre"])
        with col3:
            anios_disponibles = sorted([int(x) for x in df_global["AÑO"].unique() if x > 0], reverse=True)
            anio_sel = st.selectbox("Seleccione el Año Actual:", anios_disponibles)

        if st.button("🚀 Generar Informe Bimensual"):
            cfg_b = MAPA_BIMESTRES[bim_sel]
            nom_b_act = cfg_b["nombre"]
            nom_b_prev = MAPA_BIMESTRES[cfg_b["previo"]]["nombre"]

            # 1. APLICAR NETEO EN BASE COMPLETA
            df_base_neteada, col_val = aplicar_neteo_total_linea(df_global)

            # 2. FILTRAR POR VENDEDOR
            if col_vendedor in df_base_neteada.columns and vendedor_sel != "Todos":
                df_base_neteada = df_base_neteada[df_base_neteada[col_vendedor] == vendedor_sel]

            # 3. EXTRAER LOS 3 DATAFRAMES EXACTOS
            anio_act = int(anio_sel)
            anio_prev_ano = int(anio_sel - 1)
            anio_p3 = int(anio_sel if bim_sel != "B1" else (anio_sel - 1))

            df_p1 = df_base_neteada[(df_base_neteada["AÑO"] == anio_act) & (df_base_neteada["MES"].isin(cfg_b["meses"]))]
            df_p2 = df_base_neteada[(df_base_neteada["AÑO"] == anio_prev_ano) & (df_base_neteada["MES"].isin(cfg_b["meses"]))]
            df_p3 = df_base_neteada[(df_base_neteada["AÑO"] == anio_p3) & (df_base_neteada["MES"].isin(cfg_b["meses_prev"]))]

            # --- CONSTRUCCIÓN DE TABLA CLIENTES ---
            tabla_clis = []
            if col_cliente and col_cliente in df_base_neteada.columns:
                lista_cli_custom = []
                if 'df_clientes_custom' in st.session_state and not st.session_state['df_clientes_custom'].empty:
                    df_cc = st.session_state['df_clientes_custom']
                    lista_cli_custom = [str(x).strip().upper() for x in df_cc.iloc[:, 0].dropna().unique()]

                if lista_cli_custom:
                    for cli in lista_cli_custom:
                        va = float(df_p1[df_p1[col_cliente].astype(str).str.strip().str.upper() == cli][col_val].sum()) if not df_p1.empty else 0.0
                        vp2 = float(df_p2[df_p2[col_cliente].astype(str).str.strip().str.upper() == cli][col_val].sum()) if not df_p2.empty else 0.0
                        vp3 = float(df_p3[df_p3[col_cliente].astype(str).str.strip().str.upper() == cli][col_val].sum()) if not df_p3.empty else 0.0
                        tabla_clis.append({"CLIENTE": cli, "v_act_num": va, "v_p2_num": vp2, "v_p3_num": vp3})
                    
                    tabla_clis = sorted(tabla_clis, key=lambda x: x["v_act_num"], reverse=True)
                else:
                    clis = set(df_p1[col_cliente].dropna()).union(df_p2[col_cliente].dropna()).union(df_p3[col_cliente].dropna())
                    for c in clis:
                        va = float(df_p1[df_p1[col_cliente] == c][col_val].sum()) if not df_p1.empty else 0.0
                        vp2 = float(df_p2[df_p2[col_cliente] == c][col_val].sum()) if not df_p2.empty else 0.0
                        vp3 = float(df_p3[df_p3[col_cliente] == c][col_val].sum()) if not df_p3.empty else 0.0
                        tabla_clis.append({"CLIENTE": str(c), "v_act_num": va, "v_p2_num": vp2, "v_p3_num": vp3})
                    
                    tabla_clis = sorted(tabla_clis, key=lambda x: x["v_act_num"], reverse=True)

                # TOTALES REALES DE LA BASE COMPLETA
                tot_cli_act = float(df_p1[col_val].sum()) if not df_p1.empty else 0.0
                tot_cli_p2 = float(df_p2[col_val].sum()) if not df_p2.empty else 0.0
                tot_cli_p3 = float(df_p3[col_val].sum()) if not df_p3.empty else 0.0

                for item in tabla_clis:
                    item["v_act"] = f"${item['v_act_num']:,.2f}"
                    item["v_p2"] = f"${item['v_p2_num']:,.2f}"
                    item["v_p3"] = f"${item['v_p3_num']:,.2f}"

                tabla_clis.append({
                    "CLIENTE": "TOTAL",
                    "v_act": f"${tot_cli_act:,.2f}",
                    "v_p2": f"${tot_cli_p2:,.2f}",
                    "v_p3": f"${tot_cli_p3:,.2f}"
                })

            # --- CONSTRUCCIÓN DE TABLA FABRICANTES ---
            tabla_provs = []
            if col_prov and col_prov in df_base_neteada.columns:
                lista_fab_custom = []
                if 'df_fabricantes_custom' in st.session_state and not st.session_state['df_fabricantes_custom'].empty:
                    df_fc = st.session_state['df_fabricantes_custom']
                    lista_fab_custom = [str(x).strip().upper() for x in df_fc.iloc[:, 0].dropna().unique()]

                if lista_fab_custom:
                    for fab in lista_fab_custom:
                        va = float(df_p1[df_p1[col_prov].astype(str).str.strip().str.upper() == fab][col_val].sum()) if not df_p1.empty else 0.0
                        vp2 = float(df_p2[df_p2[col_prov].astype(str).str.strip().str.upper() == fab][col_val].sum()) if not df_p2.empty else 0.0
                        vp3 = float(df_p3[df_p3[col_prov].astype(str).str.strip().str.upper() == fab][col_val].sum()) if not df_p3.empty else 0.0
                        tabla_provs.append({"PROVEEDOR": fab, "v_act_num": va, "v_p2_num": vp2, "v_p3_num": vp3})

                    sum_fab_act = sum(x["v_act_num"] for x in tabla_provs)
                    sum_fab_p2 = sum(x["v_p2_num"] for x in tabla_provs)
                    sum_fab_p3 = sum(x["v_p3_num"] for x in tabla_provs)

                    tot_fab_act = float(df_p1[col_val].sum()) if not df_p1.empty else 0.0
                    tot_fab_p2 = float(df_p2[col_val].sum()) if not df_p2.empty else 0.0
                    tot_fab_p3 = float(df_p3[col_val].sum()) if not df_p3.empty else 0.0

                    tabla_provs = sorted(tabla_provs, key=lambda x: x["v_act_num"], reverse=True)
                    tabla_provs.append({
                        "PROVEEDOR": "Otros",
                        "v_act_num": tot_fab_act - sum_fab_act,
                        "v_p2_num": tot_fab_p2 - sum_fab_p2,
                        "v_p3_num": tot_fab_p3 - sum_fab_p3
                    })
                else:
                    provs = set(df_p1[col_prov].dropna()).union(df_p2[col_prov].dropna()).union(df_p3[col_prov].dropna())
                    for p in provs:
                        va = float(df_p1[df_p1[col_prov] == p][col_val].sum()) if not df_p1.empty else 0.0
                        vp2 = float(df_p2[df_p2[col_prov] == p][col_val].sum()) if not df_p2.empty else 0.0
                        vp3 = float(df_p3[df_p3[col_prov] == p][col_val].sum()) if not df_p3.empty else 0.0
                        tabla_provs.append({"PROVEEDOR": str(p), "v_act_num": va, "v_p2_num": vp2, "v_p3_num": vp3})
                    
                    tabla_provs = sorted(tabla_provs, key=lambda x: x["v_act_num"], reverse=True)
                    tot_fab_act = float(df_p1[col_val].sum()) if not df_p1.empty else 0.0
                    tot_fab_p2 = float(df_p2[col_val].sum()) if not df_p2.empty else 0.0
                    tot_fab_p3 = float(df_p3[col_val].sum()) if not df_p3.empty else 0.0

                for item in tabla_provs:
                    item["v_act"] = f"${item['v_act_num']:,.2f}"
                    item["v_p2"] = f"${item['v_p2_num']:,.2f}"
                    item["v_p3"] = f"${item['v_p3_num']:,.2f}"

                tabla_provs.append({
                    "PROVEEDOR": "TOTAL",
                    "v_act": f"${tot_fab_act:,.2f}",
                    "v_p2": f"${tot_fab_p2:,.2f}",
                    "v_p3": f"${tot_fab_p3:,.2f}"
                })
            else:
                tot_fab_act = float(df_p1[col_val].sum()) if not df_p1.empty else 0.0
                tot_fab_p2 = float(df_p2[col_val].sum()) if not df_p2.empty else 0.0
                tot_fab_p3 = float(df_p3[col_val].sum()) if not df_p3.empty else 0.0

            head_b_act = f"VENTA {bim_sel} {anio_sel}"
            head_b_ant_anio = f"VENTA {bim_sel} {anio_sel-1}"
            head_b_prev = f"VENTA {cfg_b['previo']} {anio_p3}"

            buf_grafica = generar_grafica_comparativa_fabricantes([head_b_act, head_b_ant_anio, head_b_prev], [tot_fab_act, tot_fab_p2, tot_fab_p3])
            
            txt_analisis, aspectos_lista, txt_cierre = generar_analisis_y_aspectos(
                tot_fab_act, tot_fab_p2, tot_fab_p3, df_p1, df_p3, nom_b_act, nom_b_prev, anio_sel, col_val, col_cliente
            )

            path_template = "templates/BIMENSUAL MAY-JUN.docx"

            contexto = {
                "vendedor": str(vendedor_sel),
                "periodo_titulo": str(cfg_b["texto_titulo"]),
                "titulo_informe": f"INFORME BIMENSUAL {cfg_b['texto_titulo'].upper()}",
                "anio": str(anio_sel),
                "col_b_act": head_b_act,
                "col_b_ant": head_b_ant_anio,
                "col_b_prev": head_b_prev,
                "tabla_proveedores": tabla_provs,
                "tabla_clientes": tabla_clis,
                "analisis_texto": txt_analisis,
                "aspectos_mejorar": aspectos_lista,
                "texto_cierre_aspectos": txt_cierre,
                "grafica_ventas": buf_grafica
            }

            try:
                doc = construir_documento_word(contexto, path_template=path_template)
                
                output_buf = io.BytesIO()
                doc.save(output_buf)
                output_buf.seek(0)

                st.success("✅ Informe generado exitosamente con todas las reglas aplicadas.")
                st.download_button(
                    label="📥 Descargar Informe Word (.docx)",
                    data=output_buf,
                    file_name=f"INFORME_BIMENSUAL_{bim_sel}_{vendedor_sel}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as err:
                st.error(f"⚠️ Error al construir el documento Word: {err}")
