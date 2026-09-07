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

def identificar_columnas_base(df):
    """Identifica automáticamente las columnas clave en el DataFrame principal."""
    cols = list(df.columns)
    
    col_vendedor = next((c for c in cols if any(k in c.upper() for k in ["VENDEDOR", "ASESOR", "COMERCIAL"])), None)
    col_prov = next((c for c in cols if any(k in c.upper() for k in ["FABRICANTE", "PROVEEDOR", "MARCA"])), None)
    col_cliente = next((c for c in cols if "NOMBRE" in c.upper() and "SN" in c.upper()), None)
    if not col_cliente:
        col_cliente = next((c for c in cols if any(k in c.upper() for k in ["CLIENTE", "RAZON_SOCIAL", "SN"])), None)
        
    col_anio = next((c for c in cols if "AÑO" in c.upper() or "ANIO" in c.upper() or "YEAR" in c.upper()), None)
    col_mes = next((c for c in cols if "MES" in c.upper() or "MONTH" in c.upper()), None)
    col_val = next((c for c in cols if "TOTAL" in c.upper() and "LINEA" in c.upper()), None)
    if not col_val:
        col_val = next((c for c in cols if any(k in c.upper() for k in ["VALOR", "VENTA", "SUBTOTAL", "TOTAL"])), None)
        
    col_tipo_doc = next((c for c in cols if "TIPO" in c.upper() and "DOC" in c.upper()), None)
    if not col_tipo_doc:
        col_tipo_doc = next((c for c in cols if "TIPO" in c.upper() or "DOCUMENTO" in c.upper()), None)

    return col_vendedor, col_prov, col_cliente, col_anio, col_mes, col_val, col_tipo_doc

def aplicar_neteo_documentos(df, col_val, col_tipo_doc):
    """Aplica la regla de neteo restando Notas Crédito/Devoluciones."""
    if not col_tipo_doc or col_tipo_doc not in df.columns:
        return df

    df = df.copy()
    
    def obtener_signo(val):
        s = str(val).strip().upper()
        if any(kw in s for kw in ["NOTA CREDITO", "NOTA CRÉDITO", "NOTA", "NC", "CREDITO", "CRÉDITO", "DEVOLUCION", "DEVOLUCIÓN"]):
            return -1.0
        return 1.0

    df["__signo_neteo"] = df[col_tipo_doc].apply(obtener_signo)
    df[col_val] = pd.to_numeric(df[col_val], errors="coerce").fillna(0.0)
    df[col_val] = df[col_val] * df["__signo_neteo"]
    df.drop(columns=["__signo_neteo"], inplace=True, errors="ignore")
    return df

def formatear_moneda(val):
    if pd.isna(val) or val is None:
        return "$0"
    val = float(val)
    if val < 0:
        return f"-${abs(val):,.0f}"
    return f"${val:,.0f}"

def calcular_variacion_pct(v_actual, v_previo):
    if v_previo == 0:
        return 100.0 if v_actual > 0 else 0.0
    return ((v_actual - v_previo) / abs(v_previo)) * 100.0

def generar_grafica_barras_tres_periodos(v_p1, v_p2, v_p3, lbl_p1, lbl_p2, lbl_p3, titulo):
    fig, ax = plt.subplots(figsize=(7, 3.5), dpi=300)
    
    periodos = [lbl_p3, lbl_p2, lbl_p1]
    valores = [v_p3, v_p2, v_p1]
    colores = ["#B0C4DE", "#4682B4", "#1F4E78"]

    bars = ax.bar(periodos, valores, color=colores, width=0.55)
    ax.set_title(titulo, fontsize=11, fontweight="bold", pad=12, color="#1F4E78")
    
    max_v = max(valores) if max(valores) > 0 else 1
    ax.set_ylim(0, max_v * 1.18)
    
    for bar in bars:
        height = bar.get_height()
        texto = formatear_moneda(height)
        ax.annotate(
            texto,
            xy=(bar.get_x() + bar.get_width() / 2, height),
            xytext=(0, 4),
            textcoords="offset points",
            ha="center", va="bottom",
            fontsize=9, fontweight="bold"
        )

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_color("#CCCCCC")
    ax.yaxis.set_visible(False)
    ax.tick_params(axis="x", labelsize=9)
    plt.tight_layout()

    buf = io.BytesIO()
    plt.savefig(buf, format="png", bbox_inches="tight", dpi=300)
    plt.close(fig)
    buf.seek(0)
    return buf

def generar_analisis_y_aspectos(tot_p1, tot_p2, tot_p3, lbl_p1, lbl_p2, lbl_p3, df_p1, col_vendedor, col_cliente, col_val, vendedor_sel):
    var_p2 = calcular_variacion_pct(tot_p1, tot_p2)
    var_p3 = calcular_variacion_pct(tot_p1, tot_p3)

    num_p1 = formatear_moneda(tot_p1)
    num_p2 = formatear_moneda(tot_p2)
    num_p3 = formatear_moneda(tot_p3)

    if var_p2 >= 0:
        txt_p2 = f"un incremento del {var_p2:.1f}% respecto al mismo período del año anterior ({lbl_p2}: {num_p2})"
    else:
        txt_p2 = f"una variación del {var_p2:.1f}% frente al mismo período del año anterior ({lbl_p2}: {num_p2})"

    if var_p3 >= 0:
        txt_p3 = f"un crecimiento del {var_p3:.1f}% comparado con el bimestre inmediatamente anterior ({lbl_p3}: {num_p3})"
    else:
        txt_p3 = f"un comportamiento de {var_p3:.1f}% en comparación con el bimestre inmediatamente anterior ({lbl_p3}: {num_p3})"

    top_clis = []
    if not df_p1.empty and col_cliente in df_p1.columns and col_val in df_p1.columns:
        grp = df_p1.groupby(col_cliente, sort=False)[col_val].sum().reset_index()
        top_clis = grp.head(2)[col_cliente].tolist()

    txt_clis = f", impulsado principalmente por cuentas clave como { ' y '.join(top_clis) }" if top_clis else ""

    analisis_texto = (
        f"Durante el período {lbl_p1}, la gestión comercial de {vendedor_sel} alcanzó una facturación neta total de "
        f"{num_p1}. Este resultado representa {txt_p2}, y {txt_p3}{txt_clis}. "
        f"A continuación se detallan las oportunidades estratégicas para sostener la dinámica de ventas."
    )

    aspectos_mejorar = [
        {
            "titulo": "Profundización de Cuentas Estratégicas",
            "descripcion": "Diseñar e implementar planes de desarrollo sobre la cartera actual para maximizar la venta cruzada y proteger la recurrencia."
        },
        {
            "titulo": "Recuperación e Incremento de Frecuencia",
            "descripcion": "Establecer contacto directo con clientes de baja rotación en el bimestre previo para reactivar pedidos y regularizar el volumen de compra."
        },
        {
            "titulo": "Diversificación de Portafolio",
            "descripcion": "Impulsar las líneas y marcas de mayor margen negociado para optimizar la rentabilidad global del portafolio en el siguiente bimestre."
        }
    ]

    texto_cierre_aspectos = "El cumplimiento de estas metas permitirá asegurar un crecimiento sostenido durante el próximo bimestre."

    return analisis_texto, aspectos_mejorar, texto_cierre_aspectos

def aplicar_formato_celda_total(row_cells):
    """Aplica formato de fondo a la última fila de total."""
    for cell in row_cells:
        shading = OxmlElement('w:shd')
        shading.set(qn('w:val'), 'clear')
        shading.set(qn('w:color'), 'auto')
        shading.set(qn('w:fill'), 'D9E1F2')
        cell._tc.get_or_add_tcPr().append(shading)

def rellenar_o_crear_tabla(doc, index_tabla, headers, rows_data):
    """Inserta o reescribe los datos en la tabla existente de la plantilla respetando su estructura."""
    if len(doc.tables) > index_tabla:
        table = doc.tables[index_tabla]
        
        while len(table.rows) > 1:
            r = table.rows[-1]._tr
            r.getparent().remove(r)
            
        for r_idx, row_values in enumerate(rows_data):
            row_cells = table.add_row().cells
            es_ultima = (r_idx == len(rows_data) - 1)
            
            for c_idx, val in enumerate(row_values):
                if c_idx < len(row_cells):
                    row_cells[c_idx].text = str(val)
                    p = row_cells[c_idx].paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.RIGHT
                    
                    if es_ultima:
                        for run in p.runs:
                            run.font.bold = True
                            
            if es_ultima:
                aplicar_formato_celda_total(row_cells)
    else:
        table = doc.add_table(rows=1, cols=len(headers))
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
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), '1F4E78')
            hdr_cells[i]._tc.get_or_add_tcPr().append(shd)

        for r_idx, row_values in enumerate(rows_data):
            row_cells = table.add_row().cells
            es_ultima = (r_idx == len(rows_data) - 1)
            for c_idx, val in enumerate(row_values):
                row_cells[c_idx].text = str(val)
                p = row_cells[c_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.RIGHT
                for run in p.runs:
                    run.font.size = Pt(9)
                    if es_ultima:
                        run.font.bold = True
            if es_ultima:
                aplicar_formato_celda_total(row_cells)

def construir_documento_word_respetando_plantilla(contexto, path_template=None):
    if path_template and os.path.exists(path_template):
        doc = Document(path_template)
    else:
        doc = Document()

    for p in doc.paragraphs:
        if "{{TITULO}}" in p.text or "INFORME BIMENSUAL" in p.text:
            p.text = contexto["titulo_informe"]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(31, 78, 120)
        elif "{{RESPONSABLE}}" in p.text:
            p.text = f"Responsable Comercial: {contexto['vendedor']} | Año: {contexto['anio']}"
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    headers_prov = ["Fabricante / Proveedor", contexto["col_b_act"], contexto["col_b_ant"], contexto["col_b_prev"]]
    rows_prov = [[item["PROVEEDOR"], item["v_act"], item["v_p2"], item["v_p3"]] for item in contexto["tabla_proveedores"]]
    rellenar_o_crear_tabla(doc, 0, headers_prov, rows_prov)

    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(contexto["grafica_ventas"], width=Inches(6.0))

    headers_cli = ["Cliente", contexto["col_b_act"], contexto["col_b_ant"], contexto["col_b_prev"]]
    rows_cli = [[item["CLIENTE"], item["v_act"], item["v_p2"], item["v_p3"]] for item in contexto["tabla_clientes"]]
    rellenar_o_crear_tabla(doc, 1, headers_cli, rows_cli)

    doc.add_paragraph(contexto["analisis_texto"])

    for asp in contexto["aspectos_mejorar"]:
        p_asp = doc.add_paragraph()
        p_asp.paragraph_format.left_indent = Inches(0.2)
        r_tit = p_asp.add_run(f"• {asp['titulo']}: ")
        r_tit.font.bold = True
        p_asp.add_run(asp["descripcion"])

    doc.add_paragraph(contexto["texto_cierre_aspectos"])

    return doc

def render_modulo_informe(df_global=None, *args, **kwargs):
    """Función principal del módulo que acepta df_global recibido desde app.py."""
    st.title("📊 Generador de Informe Bimensual Comercial")

    df_base = None

    # Verifica si la base viene transmitida desde app.py
    if df_global is not None and isinstance(df_global, pd.DataFrame) and not df_global.empty:
        df_base = df_global
        st.sidebar.success("Base de ventas recibida del sistema principal.")
    else:
        st.sidebar.header("1. Carga de Archivos")
        file_base = st.sidebar.file_uploader("Base de Ventas Principal (Excel/CSV)", type=["xlsx", "xls", "csv"])
        if file_base:
            try:
                if file_base.name.endswith(".csv"):
                    df_base = pd.read_csv(file_base)
                else:
                    df_base = pd.read_excel(file_base)
            except Exception as e:
                st.error(f"Error al leer la base principal: {e}")
                return

    st.sidebar.header("Filtros / Listas Propias (Opciones)")
    file_provs = st.sidebar.file_uploader("Filtro Opcional: Lista Fabricantes (Excel/CSV)", type=["xlsx", "xls", "csv"])
    file_clis = st.sidebar.file_uploader("Filtro Opcional: Lista Clientes (Excel/CSV)", type=["xlsx", "xls", "csv"])

    path_template = "plantilla_informe.docx" if os.path.exists("plantilla_informe.docx") else None

    if df_base is None or df_base.empty:
        st.info("Por favor, suba o cargue el archivo de ventas principal para comenzar.")
        return

    col_vendedor, col_prov, col_cliente, col_anio, col_mes, col_val, col_tipo_doc = identificar_columnas_base(df_base)

    if not all([col_vendedor, col_prov, col_cliente, col_anio, col_mes, col_val]):
        st.error("No se pudieron identificar automáticamente todas las columnas requeridas en el archivo base.")
        return

    df_base_net = aplicar_neteo_documentos(df_base, col_val, col_tipo_doc)

    st.sidebar.header("2. Parámetros del Informe")
    
    # ORDEN NATURAL: Conserva el orden en el que viene el archivo original
    vendedores = df_base_net[col_vendedor].dropna().drop_duplicates().astype(str).tolist()
    vendedor_sel = st.sidebar.selectbox("Seleccionar Vendedor:", vendedores)

    anios = df_base_net[col_anio].dropna().drop_duplicates().astype(int).tolist()
    anio_sel = st.sidebar.selectbox("Seleccionar Año:", anios)

    bim_sel_k = st.sidebar.selectbox("Seleccionar Bimestre:", list(MAPA_BIMESTRES.keys()), format_func=lambda k: MAPA_BIMESTRES[k]["nombre"])
    info_bim = MAPA_BIMESTRES[bim_sel_k]

    lista_prov_custom = []
    if file_provs:
        df_p_c = pd.read_csv(file_provs) if file_provs.name.endswith(".csv") else pd.read_excel(file_provs)
        lista_prov_custom = df_p_c.iloc[:, 0].dropna().tolist()

    lista_cli_custom = []
    if file_clis:
        df_c_c = pd.read_csv(file_clis) if file_clis.name.endswith(".csv") else pd.read_excel(file_clis)
        lista_cli_custom = df_c_c.iloc[:, 0].dropna().tolist()

    anio_act = anio_sel
    meses_act = info_bim["meses"]

    anio_prev_ano = anio_sel - 1
    meses_prev_ano = info_bim["meses"]

    if info_bim["previo"] == "B6":
        anio_p3 = anio_sel - 1
    else:
        anio_p3 = anio_sel
    meses_p3 = info_bim["meses_prev"]

    df_vend = df_base_net[df_base_net[col_vendedor].astype(str) == str(vendedor_sel)]

    df_p1 = df_vend[(df_vend[col_anio] == anio_act) & (df_vend[col_mes].isin(meses_act))]
    df_p2 = df_vend[(df_vend[col_anio] == anio_prev_ano) & (df_vend[col_mes].isin(meses_prev_ano))]
    df_p3 = df_vend[(df_vend[col_anio] == anio_p3) & (df_vend[col_mes].isin(meses_p3))]

    tot_p1 = float(df_p1[col_val].sum())
    tot_p2 = float(df_p2[col_val].sum())
    tot_p3 = float(df_p3[col_val].sum())

    lbl_p1 = f"{info_bim['texto_titulo']} {anio_act}"
    lbl_p2 = f"{info_bim['texto_titulo']} {anio_prev_ano}"
    lbl_p3 = f"{MAPA_BIMESTRES[info_bim['previo']]['texto_titulo']} {anio_p3}"

    col_b_act = f"Ventas {info_bim['nombre'].split()[0]} {anio_act}"
    col_b_ant = f"Ventas {info_bim['nombre'].split()[0]} {anio_prev_ano}"
    col_b_prev = f"Ventas {info_bim['previo']} {anio_p3}"

    # ORDEN NATURAL FABRICANTES: Mantiene exactamente la secuencia en que están registrados
    provs = lista_prov_custom if lista_prov_custom else df_vend[col_prov].dropna().drop_duplicates().tolist()
    tabla_proveedores = []
    for pr in provs:
        v_a = float(df_p1[df_p1[col_prov] == pr][col_val].sum()) if not df_p1.empty else 0.0
        v_2 = float(df_p2[df_p2[col_prov] == pr][col_val].sum()) if not df_p2.empty else 0.0
        v_3 = float(df_p3[df_p3[col_prov] == pr][col_val].sum()) if not df_p3.empty else 0.0
        tabla_proveedores.append({
            "PROVEEDOR": pr,
            "v_act": formatear_moneda(v_a),
            "v_p2": formatear_moneda(v_2),
            "v_p3": formatear_moneda(v_3)
        })

    tabla_proveedores.append({
        "PROVEEDOR": "TOTAL",
        "v_act": formatear_moneda(tot_p1),
        "v_p2": formatear_moneda(tot_p2),
        "v_p3": formatear_moneda(tot_p3)
    })

    # ORDEN NATURAL CLIENTES: Mantiene exactamente la secuencia en que están registrados
    clis = lista_cli_custom if lista_cli_custom else df_vend[col_cliente].dropna().drop_duplicates().tolist()
    tabla_clientes = []
    for cl in clis:
        v_a = float(df_p1[df_p1[col_cliente] == cl][col_val].sum()) if not df_p1.empty else 0.0
        v_2 = float(df_p2[df_p2[col_cliente] == cl][col_val].sum()) if not df_p2.empty else 0.0
        v_3 = float(df_p3[df_p3[col_cliente] == cl][col_val].sum()) if not df_p3.empty else 0.0
        tabla_clientes.append({
            "CLIENTE": cl,
            "v_act": formatear_moneda(v_a),
            "v_p2": formatear_moneda(v_2),
            "v_p3": formatear_moneda(v_3)
        })

    tabla_clientes.append({
        "CLIENTE": "TOTAL",
        "v_act": formatear_moneda(tot_p1),
        "v_p2": formatear_moneda(tot_p2),
        "v_p3": formatear_moneda(tot_p3)
    })

    titulo_grafica = f"VENTAS {info_bim['texto_titulo'].upper()} {anio_act}"
    grafica_buf = generar_grafica_barras_tres_periodos(tot_p1, tot_p2, tot_p3, lbl_p1, lbl_p2, lbl_p3, titulo_grafica)

    analisis_txt, aspectos, cierre_txt = generar_analisis_y_aspectos(
        tot_p1, tot_p2, tot_p3, lbl_p1, lbl_p2, lbl_p3, df_p1, col_vendedor, col_cliente, col_val, vendedor_sel
    )

    contexto = {
        "titulo_informe": f"INFORME BIMENSUAL DE VENTAS - {info_bim['texto_titulo'].upper()} {anio_act}",
        "vendedor": vendedor_sel,
        "anio": anio_act,
        "col_b_act": col_b_act,
        "col_b_ant": col_b_ant,
        "col_b_prev": col_b_prev,
        "tabla_proveedores": tabla_proveedores,
        "tabla_clientes": tabla_clientes,
        "grafica_ventas": grafica_buf,
        "analisis_texto": analisis_txt,
        "aspectos_mejorar": aspectos,
        "texto_cierre_aspectos": cierre_txt
    }

    st.subheader("Vista Previa del Informe")
    st.write(f"**Vendedor:** {vendedor_sel} | **Período:** {lbl_p1}")
    st.image(grafica_buf, width=600)

    st.write("**Resumen de Fabricantes / Proveedores:**")
    st.dataframe(pd.DataFrame(tabla_proveedores))

    st.write("**Resumen de Clientes:**")
    st.dataframe(pd.DataFrame(tabla_clientes))

    doc = construir_documento_word_respetando_plantilla(contexto, path_template)
    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)

    st.download_button(
        label="📥 Descargar Informe en Word (.docx)",
        data=out_buf,
        file_name=f"Informe_Bimensual_{vendedor_sel}_{info_bim['nombre'].split()[0]}_{anio_act}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

if __name__ == "__main__":
    render_modulo_informe()
