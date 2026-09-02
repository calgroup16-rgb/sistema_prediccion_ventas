import streamlit as st
import pandas as pd
import numpy as np
from informe_bimensual import render_modulo_informe

# Configuración de página
st.set_page_config(
    page_title="Sistema de Predicción Comercial",
    page_icon="📊",
    layout="wide"
)

# Palabras clave para detectar sector de alimentación animal
PALABRAS_CLAVE_ANIMAL = [
    'ALIMENTO', 'BALANCEADO', 'FORRAJE', 'GANADO', 'AVICOLA', 'AVÍCOLA',
    'PORCINO', 'PREMEZCLA', 'HARINA', 'MASCOTAS', 'ACUÍCOLA', 'ACUICOLA',
    'NUTRICION', 'NUTRICIONAL', 'AQUA', 'AGRO', 'ANIMAL', 'VETERINARIA',
    'BOVINO', 'AVICULTURA', 'AVICOLAS'
]

def cargar_y_preparar_datos(file):
    """Carga el Excel de ventas y estandariza las columnas requeridas."""
    df = pd.read_excel(file)
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    # Identificar columna de Cliente (priorizando NOMBRE SN)
    col_cliente = None
    if 'NOMBRE SN' in df.columns:
        col_cliente = 'NOMBRE SN'
    else:
        col_cliente = next((c for c in df.columns if 'CLIENTE' in c or 'NOMBRE' in c or 'SN' in c), None)

    col_fecha = next((c for c in df.columns if 'FECHA' in c or 'ANIO' in c or 'AÑO' in c), None)
    col_valor = next((c for c in df.columns if 'VALOR' in c or 'VENTA' in c or 'MONTO' in c or 'TOTAL' in c), None)
    col_prod = next((c for c in df.columns if 'PRODUCTO' in c or 'ITEM' in c or 'DESCRIPCION' in c), None)

    if not all([col_cliente, col_fecha, col_valor, col_prod]):
        st.error("No se pudieron identificar las columnas requeridas en la base principal.")
        return None

    df = df.rename(columns={
        col_cliente: 'CLIENTE',
        col_fecha: 'FECHA',
        col_valor: 'VALOR_VENTA',
        col_prod: 'PRODUCTO'
    })

    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df['ANIO'] = df['FECHA'].dt.year
    df['VALOR_VENTA'] = pd.to_numeric(df['VALOR_VENTA'], errors='coerce').fillna(0)
    df['CLIENTE'] = df['CLIENTE'].astype(str).str.strip()
    df['PRODUCTO'] = df['PRODUCTO'].astype(str).str.strip()

    return df

st.title("📊 Sistema de Predicción Comercial y Análisis de Ventas")

# --- BARRA LATERAL ---
st.sidebar.header("📂 1. Base de Datos Principal")
archivo_ventas = st.sidebar.file_uploader("Subir Excel de Ventas Global", type=["xlsx", "xls"])

if archivo_ventas is not None:
    st.session_state['df_ventas'] = cargar_y_preparar_datos(archivo_ventas)

if 'df_ventas' in st.session_state and st.session_state['df_ventas'] is not None:
    df_global = st.session_state['df_ventas']

    st.sidebar.markdown("---")
    st.sidebar.header("🎯 2. Filtro / Lista Propia (Opcional)")

    archivo_lista_clientes = st.sidebar.file_uploader(
        "Subir Listado Especifico (Excel)", 
        type=["xlsx", "xls"]
    )

    todos_clientes = sorted(df_global['CLIENTE'].dropna().unique().tolist())
    clientes_manuales = st.sidebar.multiselect(
        "O seleccione cliente(s) por 'NOMBRE SN':",
        options=todos_clientes,
        placeholder="Seleccione 1 o varios clientes..."
    )

    # DEFINICIÓN DEL GRUPO DE TRABAJO (GLOBAL O FILTRADO)
    if archivo_lista_clientes is not None:
        df_lista = pd.read_excel(archivo_lista_clientes)
        df_lista.columns = [str(col).strip().upper() for col in df_lista.columns]
        col_target = 'NOMBRE SN' if 'NOMBRE SN' in df_lista.columns else df_lista.columns[0]
        lista_nombres = df_lista[col_target].dropna().astype(str).str.strip().unique().tolist()
        df_trabajo = df_global[df_global['CLIENTE'].isin(lista_nombres)]
        st.sidebar.success(f"Análisis enfocado en {len(lista_nombres)} cliente(s) de la lista.")
    elif clientes_manuales:
        df_trabajo = df_global[df_global['CLIENTE'].isin(clientes_manuales)]
        st.sidebar.success(f"Análisis enfocado en {len(clientes_manuales)} cliente(s) seleccionado(s).")
    else:
        df_trabajo = df_global.copy()
        st.sidebar.info("Modo Global: Analizando TODOS los clientes del Excel principal.")

    # --- PESTAÑAS PRINCIPALES CON PREDICCIONES ---
    tab1, tab2, tab3 = st.tabs([
        "💡 Sugerido de Venta & Productos Dejados de Comprar", 
        "⚠️ Predicción Top 20 Inactivos 2026", 
        "🌾 Predicción Sector Alimentación Animal & Cross-Selling"
    ])

    # -------------------------------------------------------------
    # PESTAÑA 1: PREDICCIÓN Y SUGERIDO DE VENTA
    # -------------------------------------------------------------
    with tab1:
        if not df_trabajo.empty:
            st.header("Predicción Comercial de Reenganche y Sugeridos")
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Clientes Evaluados", df_trabajo['CLIENTE'].nunique())
            m2.metric("Ventas Acumuladas ($)", f"${df_trabajo['VALOR_VENTA'].sum():,.2f}")
            m3.metric("Transacciones Analizadas", f"{len(df_trabajo):,}")

            st.markdown("---")
            st.subheader("⚠️ Productos que DEJARON DE COMPRAR en el Último Periodo")
            
            max_anio = df_global['ANIO'].max()
            
            hist_compra = df_trabajo.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()
            reciente_compra = df_trabajo[df_trabajo['ANIO'] == max_anio].groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()

            merged = pd.merge(hist_compra, reciente_compra, on=['CLIENTE', 'PRODUCTO'], how='left', suffixes=('_HISTORICO', f'_{max_anio}'))
            merged[f'VALOR_VENTA_{max_anio}'] = merged[f'VALOR_VENTA_{max_anio}'].fillna(0)

            dejaron = merged[merged[f'VALOR_VENTA_{max_anio}'] == 0].copy()
            dejaron = dejaron.sort_values(by=['CLIENTE', 'VALOR_VENTA_HISTORICO'], ascending=[True, False])
            dejaron.columns = ['NOMBRE SN', 'PRODUCTO DEJADO DE COMPRAR', 'VENTA HISTÓRICA ($)', f'VENTA {max_anio} ($)']

            if not dejaron.empty:
                st.dataframe(dejaron[['NOMBRE SN', 'PRODUCTO DEJADO DE COMPRAR', 'VENTA HISTÓRICA ($)']], use_container_width=True)
                
                st.markdown("---")
                st.subheader("🎯 Predicción: Productos Sugeridos a Vender/Ofrecer")
                sugerido = dejaron.groupby('NOMBRE SN').first().reset_index()
                sugerido['ESTRATEGIA PREDICHA'] = "Prioridad Alta: Reofrecer producto top histórico que dejó de consumir"
                st.dataframe(sugerido[['NOMBRE SN', 'PRODUCTO DEJADO DE COMPRAR', 'VENTA HISTÓRICA ($)', 'ESTRATEGIA PREDICHA']], use_container_width=True)
            else:
                st.info("No se detectó abandono de productos en el período reciente para los clientes analizados.")
        else:
            st.warning("No se encontraron registros para los clientes solicitados.")

    # -------------------------------------------------------------
    # PESTAÑA 2: PREDICCIÓN TOP 20 INACTIVOS 2026
    # -------------------------------------------------------------
    with tab2:
        st.header("Top 20 Clientes con Mayor Riesgo de Fuga (Inactivos en 2026)")
        st.caption("Filtro de predicción: Clientes con volumen importante (2019-2025) que registraron $0 ventas en 2026.")

        clientes_2026 = set(df_trabajo[df_trabajo['ANIO'] == 2026]['CLIENTE'].unique())
        df_hist = df_trabajo[(df_trabajo['ANIO'] >= 2019) & (df_trabajo['ANIO'] <= 2025)]
        inactivos = df_hist[~df_hist['CLIENTE'].isin(clientes_2026)]

        if not inactivos.empty:
            top_inactivos = inactivos.groupby('CLIENTE')['VALOR_VENTA'].sum().reset_index()
            top_inactivos = top_inactivos.sort_values(by='VALOR_VENTA', ascending=False).head(20)
            top_inactivos.columns = ['NOMBRE SN', 'VENTA ACUMULADA HISTÓRICA 2019-2025 ($)']

            st.dataframe(top_inactivos, use_container_width=True)

            st.markdown("---")
            st.subheader("📦 Plan de Reenganche: Productos de Mayor Impacto por Cliente")
            
            df_inactivos_det = df_hist[df_hist['CLIENTE'].isin(top_inactivos['NOMBRE SN'])]
            rec_inactivos = df_inactivos_det.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()
            rec_top3 = rec_inactivos.sort_values(by=['CLIENTE', 'VALOR_VENTA'], ascending=[True, False]).groupby('CLIENTE').head(3)
            rec_top3.columns = ['NOMBRE SN', 'PRODUCTO CLAVE A REOFRECER', 'VALOR HISTÓRICO ($)']
            st.dataframe(rec_top3, use_container_width=True)
        else:
            st.info("No se hallaron clientes inactivos bajo estas condiciones en la selección actual.")

    # -------------------------------------------------------------
    # PESTAÑA 3: ALIMENTACIÓN ANIMAL & CROSS-SELLING
    # -------------------------------------------------------------
    with tab3:
        st.header("🌾 Predicción de Oportunidades: Sector Alimentación Animal")
        st.caption("Identificación automática de clientes del rubro animal y predicción de nuevos insumos compatibles a vender.")

        patron = '|'.join(PALABRAS_CLAVE_ANIMAL)
        es_cliente = df_trabajo['CLIENTE'].str.contains(patron, case=False, na=False)
        es_prod = df_trabajo['PRODUCTO'].str.contains(patron, case=False, na=False)

        df_animal = df_trabajo[es_cliente | es_prod]

        if not df_animal.empty:
            top_animal = df_animal.groupby('CLIENTE')['VALOR_VENTA'].sum().nlargest(10).reset_index()
            top_animal.columns = ['NOMBRE SN', 'VENTAS TOTALES ACUMULADAS ($)']

            col1, col2 = st.columns([1, 1])
            with col1:
                st.subheader("Top Clientes Identificados")
                st.dataframe(top_animal, use_container_width=True)

            with col2:
                st.subheader("💡 Predicción de Insumos Novedosos a Vender")
                
                # Identificar productos del rubro animal en el catálogo general que este cliente aún NO ha comprado
                catalog_animal = df_global[df_global['PRODUCTO'].str.contains(patron, case=False, na=False)]['PRODUCTO'].unique()
                
                predicciones = []
                for cli in top_animal['NOMBRE SN']:
                    comprados = set(df_trabajo[df_trabajo['CLIENTE'] == cli]['PRODUCTO'].unique())
                    oportunidades = [p for p in catalog_animal if p not in comprados]
                    sugerido_nuevo = oportunidades[0] if oportunidades else "Portafolio de nutrición animal cubierto"
                    
                    predicciones.append({
                        'NOMBRE SN': cli,
                        'POSIBLE INSUMO A OFRECER': sugerido_nuevo
                    })
                
                df_pred = pd.DataFrame(predicciones)
                st.dataframe(df_pred, use_container_width=True)
        else:
            st.info("No se detectaron empresas o productos del sector de alimentación animal en la selección actual.")

else:
    st.info("👋 Por favor, suba el archivo Excel general de ventas en el menú de la izquierda para desplegar la información.")
from informe_bimensual import render_modulo_informe

# ... (Todo tu código previo de app.py permanece igual) ...

# Agregar el módulo de generación de informe al final de tu app
st.divider()
if 'df' in locals():
    render_modulo_informe(df)
