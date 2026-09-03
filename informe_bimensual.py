import os
import io
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
import jinja2

MAPA_BIMESTRES = {
    "B1": {"nombre": "B1 (Ene - Feb)", "texto_titulo": "Enero - Febrero", "meses": [1, 2], "previo": "B6", "meses_prev": [11, 12]},
    "B2": {"nombre": "B2 (Mar - Abr)", "texto_titulo": "Marzo - Abril", "meses": [3, 4], "previo": "B1", "meses_prev": [1, 2]},
    "B3": {"nombre": "B3 (May - Jun)", "texto_titulo": "Mayo - Junio", "meses": [5, 6], "previo": "B2", "meses_prev": [3, 4]},
    "B4": {"nombre": "B4 (Jul - Ago)", "texto_titulo": "Julio - Agosto", "meses": [7, 8], "previo": "B3", "meses_prev": [5, 6]},
    "B5": {"nombre": "B5 (Sep - Oct)", "texto_titulo": "Septiembre - Octubre", "meses": [9, 10], "previo": "B4", "meses_prev": [7, 8]},
    "B6": {"nombre": "B6 (Nov - Dic)", "texto_titulo": "Noviembre - Diciembre", "meses": [11, 12], "previo": "B5", "meses_prev": [9, 10]},
}

def preparar_dataframe(df):
    """Asegura que existan las columnas de AÑO y MES extraídas desde la fecha si no existen explícitamente."""
    if df is None or df.empty:
        return df

    df = df.copy()
    
    # 1. Detectar o calcular MES y AÑO a partir de columnas de fecha si no están
    if "MES" not in df.columns or "AÑO" not in df.columns:
        col_fecha = None
        for col in df.columns:
            if "FECHA" in str(col).upper():
                col_fecha = col
                break
        
        if col_fecha:
            df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
            if "MES" not in df.columns:
                df["MES"] = df[col_fecha].dt.month
            if "AÑO" not in df.columns:
                df["AÑO"] = df[col_fecha].dt.year

    return df


def generar_grafica_comparativa(labels, valores_totales):
    fig, ax = plt.subplots(figsize=(6.5, 3.5))
    colores = ['#1f77b4', '#ff7f0e', '#2ca02c']
    
    bars = ax.bar(labels, valores_totales, color=colores, width=0.45)
    ax.set_ylabel('Ventas ($)', fontsize=10, fontfamily='sans-serif')
    ax.set_title('Comparativo de Ventas Totales por Período', fontsize=11, fontweight='bold', fontfamily='sans-serif')
    
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


def generar_analisis_y_aspectos(v_act, v_p2, v_p3, df_p1, col_prov, nom_b_act, nom_b_prev, anio_actual):
    var_vs_anio = ((v_act - v_p2) / v_p2 * 100) if v_p2 > 0 else 0
    var_vs_bim = ((v_act - v_p3) / v_p3 * 100) if v_p3 > 0 else 0

    tend_anio = "un crecimiento" if var_vs_anio >= 0 else "un decrecimiento"
    tend_bim = "un incremento" if var_vs_bim >= 0 else "una disminución"

    texto_analisis = (
        f"Durante el período {nom_b_act} de {anio_actual}, se alcanzaron ventas totales de ${v_act:,.2f}. "
        f"Al comparar este resultado con el mismo período del año anterior ({nom_b_act} {anio_actual-1}), "
        f"se evidencia {tend_anio} del {abs(var_vs_anio):.2f}% (frente a ${v_p2:,.2f}). "
        f"Por otra parte, en relación con el comportamiento del bimestre inmediatamente anterior ({nom_b_prev}), se registró "
        f"{tend_bim} del {abs(var_vs_bim):.2f}% respecto a los ${v_p3:,.2f} facturados previamente."
    )

    col_val = "VALOR_VENTA" if "VALOR_VENTA" in df_p1.columns else ("VALOR" if "VALOR" in df_p1.columns else "TOTAL")
    col_cli = "CLIENTE" if "CLIENTE" in df_p1.columns else "NOMBRE_CLIENTE"
    
    top_clis = df_p1.groupby(col_cli)[col_val].sum().sort_values(ascending=False).index.tolist() if col_cli in df_p1.columns and not df_p1.empty else []
    cli_1 = str(top_clis[0]) if len(top_clis) > 0 else "cuentas principales"
    cli_2 = str(top_clis[1]) if len(top_clis) > 1 else "cuentas secundarias"
    cli_3 = str(top_clis[2]) if len(top_clis) > 2 else "otros clientes estratégicos"

    aspectos = [
        {"titulo": "Recuperación de negocios pendientes y mayor seguimiento comercial", "descripcion": f"Se requiere fortalecer el seguimiento a clientes con procesos de compra pendientes como {cli_1} y {cli_2}, estableciendo fechas de compromiso y realizando un acompañamiento más cercano con las áreas técnicas y de compras."},
        {"titulo": "Mejorar la disponibilidad de inventario en productos estratégicos", "descripcion": "Algunos negocios se vieron afectados por retrasos derivados del desabastecimiento de productos de alta rotación, principalmente sabores, estándares de crioscopio y otros insumos especializados."},
        {"titulo": "Incrementar la participación de las líneas CHARM y productos de mayor rentabilidad", "descripcion": f"Existe una oportunidad importante para fortalecer la comercialización de productos como lactasa, hisopos de luminometria, GMP y demás soluciones CHARM, aprovechando la cartera activa de {cli_3}."},
        {"titulo": "Recuperación de clientes y productos con disminución en ventas", "descripcion": "Se evidencian reducciones en la compra de algunos productos representativos dentro del portafolio. Es necesario identificar las causas de la disminución y presentar alternativas comerciales."},
        {"titulo": "Aumentar la venta cruzada en clientes activos", "descripcion": "Clientes con compras recurrentes representan una oportunidad constante para incrementar la participación mediante la incorporación de nuevas líneas de negocio."},
        {"titulo": "Fortalecer la planeación comercial con clientes estratégicos", "descripcion": "Promover que los clientes compartan sus proyecciones mensuales o trimestrales de consumo permitirá mejorar la planeación de inventarios y anticipar necesidades."},
        {"titulo": "Continuar la recuperación de cuentas con alto potencial", "descripcion": "Se continuará con las gestiones comerciales para recuperar clientes y líneas de negocio que presentan oportunidades de crecimiento."},
        {"titulo": "Optimización de la gestión de cartera y recaudo oportuno", "descripcion": "Monitorear semanalmente la conversión de facturación a cartera efectiva para asegurar el flujo de caja sin frenar la colocación de nuevos pedidos."}
    ]

    texto_cierre = "De acuerdo con la revisión, es importante potenciar la venta de productos como lactasa, hisopos, GMP y en general las líneas de negocio de CHARM."

    return texto_analisis, aspectos, texto_cierre


def render_modulo_informe(df_global):
    st.title("📄 Generador de Informe Bimensual")

    if df_global is None or df_global.empty:
        st.warning("⚠️ Primero debe cargar los datos de ventas en la aplicación.")
        return

    # Preparar columnas de fechas/meses sin alterar el DF original
    df_global = preparar_dataframe(df_global)

    # LAS TRES PESTAÑAS ORIGINALES DEL MÓDULO
    tab_gen, tab_clientes, tab_fabricantes = st.tabs(["🚀 Generar Informe", "📋 Cargar Lista Clientes", "🏭 Cargar Lista Fabricantes"])

    with tab_clientes:
        st.subheader("Cargar Archivo de Clientes")
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
        st.subheader("Cargar Archivo de Fabricantes / Proveedores")
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
        # FILTRAR ÚNICAMENTE FACTURAS
        col_tipo_doc = None
        for col in ["TIPO_DOCUMENTO", "TIPO_DOC", "DOCUMENTO", "TIPO"]:
            if col in df_global.columns:
                col_tipo_doc = col
                break

        if col_tipo_doc:
            df_global = df_global[df_global[col_tipo_doc].astype(str).str.upper().str.contains("FACTURA", na=False)]

        # Detectar columnas de Vendedor, Cliente, Proveedor y Valor
        col_vendedor = "VENDEDOR" if "VENDEDOR" in df_global.columns else "RESPONSABLE"
        col_cliente = "CLIENTE" if "CLIENTE" in df_global.columns else "NOMBRE_CLIENTE"
        col_prov = "PROVEEDOR" if "PROVEEDOR" in df_global.columns else ("FABRICANTE" if "FABRICANTE" in df_global.columns else "MARCA")
        col_val = "VALOR_VENTA" if "VALOR_VENTA" in df_global.columns else ("VALOR" if "VALOR" in df_global.columns else "TOTAL")

        st.subheader("Configuración del Informe")
        col1, col2, col3 = st.columns(3)

        with col1:
            vendedor_sel = st.selectbox("Seleccione el Vendedor/Responsable:", sorted(df_global[col_vendedor].dropna().unique()))
        with col2:
            bim_sel = st.selectbox("Seleccione el Bimestre del Informe:", list(MAPA_BIMESTRES.keys()), format_func=lambda x: MAPA_BIMESTRES[x]["nombre"])
        with col3:
            anios_disponibles = sorted(df_global["AÑO"].dropna().astype(int).unique(), reverse=True)
            anio_sel = st.selectbox("Seleccione el Año Actual:", anios_disponibles)

        if st.button("🚀 Generar Informe Bimensual"):
            cfg_b = MAPA_BIMESTRES[bim_sel]
            nom_b_act = cfg_b["nombre"]
            nom_b_prev = MAPA_BIMESTRES[cfg_b["previo"]]["nombre"]

            df_vend = df_global[df_global[col_vendedor] == vendedor_sel]
            
            df_p1 = df_vend[(df_vend["AÑO"] == anio_sel) & (df_vend["MES"].isin(cfg_b["meses"]))]
            df_p2 = df_vend[(df_vend["AÑO"] == (anio_sel - 1)) & (df_vend["MES"].isin(cfg_b["meses"]))]
            
            anio_p3 = anio_sel if bim_sel != "B1" else (anio_sel - 1)
            df_p3 = df_vend[(df_vend["AÑO"] == anio_p3) & (df_vend["MES"].isin(cfg_b["meses_prev"]))]

            v_act = float(df_p1[col_val].sum()) if not df_p1.empty else 0.0
            v_p2 = float(df_p2[col_val].sum()) if not df_p2.empty else 0.0
            v_p3 = float(df_p3[col_val].sum()) if not df_p3.empty else 0.0

            # TABLA FABRICANTES
            tabla_provs = []
            if col_prov in df_vend.columns:
                provs = set(df_p1[col_prov].dropna()).union(df_p2[col_prov].dropna()).union(df_p3[col_prov].dropna())
                for p in provs:
                    va = float(df_p1[df_p1[col_prov] == p][col_val].sum()) if not df_p1.empty else 0.0
                    vp2 = float(df_p2[df_p2[col_prov] == p][col_val].sum()) if not df_p2.empty else 0.0
                    vp3 = float(df_p3[df_p3[col_prov] == p][col_val].sum()) if not df_p3.empty else 0.0
                    tabla_provs.append({"PROVEEDOR": str(p), "v_act": f"${va:,.2f}", "v_p2": f"${vp2:,.2f}", "v_p3": f"${vp3:,.2f}", "raw_val": va})
                tabla_provs = sorted(tabla_provs, key=lambda x: x["raw_val"], reverse=True)

            # TABLA CLIENTES
            tabla_clis = []
            if col_cliente in df_vend.columns:
                clis = set(df_p1[col_cliente].dropna()).union(df_p2[col_cliente].dropna()).union(df_p3[col_cliente].dropna())
                for c in clis:
                    va = float(df_p1[df_p1[col_cliente] == c][col_val].sum()) if not df_p1.empty else 0.0
                    vp2 = float(df_p2[df_p2[col_cliente] == c][col_val].sum()) if not df_p2.empty else 0.0
                    vp3 = float(df_p3[df_p3[col_cliente] == c][col_val].sum()) if not df_p3.empty else 0.0
                    tabla_clis.append({"CLIENTE": str(c), "v_act": f"${va:,.2f}", "v_p2": f"${vp2:,.2f}", "v_p3": f"${vp3:,.2f}", "raw_val": va})
                tabla_clis = sorted(tabla_clis, key=lambda x: x["raw_val"], reverse=True)

            head_b_act = f"VENTA {bim_sel} {anio_sel}"
            head_b_ant_anio = f"VENTA {bim_sel} {anio_sel-1}"
            head_b_prev = f"VENTA {cfg_b['previo']} {anio_p3}"

            buf_grafica = generar_grafica_comparativa([head_b_act, head_b_ant_anio, head_b_prev], [v_act, v_p2, v_p3])
            txt_analisis, aspectos_lista, txt_cierre = generar_analisis_y_aspectos(v_act, v_p2, v_p3, df_p1, col_prov, nom_b_act, nom_b_prev, anio_sel)

            path_template = "templates/BIMENSUAL MAY-JUN.docx"
            if not os.path.exists(path_template):
                st.error(f"❌ No se encontró la plantilla Word en {path_template}")
                return

            # Carga segura mediante stream en memoria para el archivo Word
            with open(path_template, "rb") as f:
                template_bytes = io.BytesIO(f.read())

            doc = DocxTemplate(template_bytes)

            contexto = {
                "vendedor": str(vendedor_sel),
                "nombre_responsable": str(vendedor_sel),
                "periodo_titulo": str(cfg_b["texto_titulo"]),
                "titulo_informe": f"INFORME BIMENSUAL {cfg_b['texto_titulo'].upper()} - {vendedor_sel.upper()}",
                "anio": str(anio_sel),
                "col_b_act": head_b_act,
                "col_b_ant": head_b_ant_anio,
                "col_b_prev": head_b_prev,
                "tabla_proveedores": tabla_provs,
                "tabla_clientes": tabla_clis,
                "total_b_act": f"${v_act:,.2f}",
                "total_b_ant": f"${v_p2:,.2f}",
                "total_b_prev": f"${v_p3:,.2f}",
                "analisis_texto": txt_analisis,
                "aspectos_mejorar": aspectos_lista,
                "texto_cierre_aspectos": txt_cierre,
                "grafica_ventas": InlineImage(doc, buf_grafica, width=Inches(6.0))
            }

            try:
                jinja_env = jinja2.Environment(autoescape=False)
                doc.render(contexto, jinja_env)
                
                output_buf = io.BytesIO()
                doc.save(output_buf)
                output_buf.seek(0)

                st.success("✅ Informe generado correctamente.")
                st.download_button(
                    label="📥 Descargar Informe Word (.docx)",
                    data=output_buf,
                    file_name=f"INFORME_BIMENSUAL_{bim_sel}_{vendedor_sel}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as err:
                st.error(f"⚠️ Error al renderizar plantilla Word: {err}")
