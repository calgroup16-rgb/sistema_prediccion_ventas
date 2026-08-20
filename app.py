import streamlit as st
import pandas as pd
import numpy as np

# Configuración de página
st.set_page_config(
    page_title="Sistema de Predicción Comercial",
    page_icon="📊",
    layout="wide"
)

# Terminology y palabras clave para detectar sector de alimentación animal
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
        st.error(f"No se pudieron identificar las columnas requeridas.")
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
        "Subir Listado de Clientes (Excel)", 
        type=["xlsx", "xls"]
    )

    todos_clientes = sorted(df_global['CLIENTE'].dropna().unique().tolist())
    clientes_manuales = st.sidebar.multiselect(
        "O seleccione cliente(s) individualmente:",
        options=todos_clientes,
        placeholder="Seleccione 1 o varios clientes..."
    )

    # LÓGICA DE FILTRADO
    # Si hay archivo o selección manual -> usar solo esos clientes.
    # Si no hay nada -> usar TODOS los clientes del Excel principal.
    if archivo_lista_clientes is not None:
        df_lista = pd.read_excel(archivo_lista_clientes)
        df_lista.columns = [str(col).strip().upper() for col in df_lista.columns]
        col_target = 'NOMBRE SN' if 'NOMBRE SN' in df_lista.columns else df_lista.columns[0]
        lista_nombres = df_lista[col_target].dropna().astype(str).str.strip().unique().tolist()
        df_trabajo = df_global[df_global['CLIENTE'].isin(lista_nombres)]
        st.sidebar.success(f"Filtrado activo: {len(lista_nombres)} cliente(s) solicitados.")
    elif clientes_manuales:
        df_trabajo = df_global[df_global['CLIENTE'].isin(clientes_manuales)]
        st.sidebar.success(f"Filtrado activo: {len(clientes_manuales)} cliente(s) seleccionado(s).")
    else:
        df_trabajo = df_global.copy()
        st.sidebar.info("Modo Global: Analizando TODOS los clientes de la base de datos.")

    # --- PESTAÑAS PRINCIPALES ---
    tab1, tab2, tab3 = st.tabs([
        "📋 Análisis & Sugerido Comercial", 
        "⚠️ Top Inactivos 2026", 
        "🌾 Detección: Sector Alimentación Animal"
    ])

    # -------------------------------------------------------------
    # PESTAÑA 1: ANÁLISIS DE CLIENTES Y SUGERIDO
    # -------------------------------------------------------------
    with tab1:
        if not df_trabajo.empty:
            st.header("Análisis de Ventas y Productos Sugeridos")
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Clientes en Análisis", df_trabajo['CLIENTE'].nunique())
            m2.metric("Ventas Totales ($)", f"${df_trabajo['VALOR_VENTA'].sum():,.2f}")
            m3.metric("Transacciones Registradas", f"{len(df_trabajo):,}")

            st.markdown("---")
            st.subheader("⚠️ Productos que DEJARON DE COMPRAR")
            
            max_anio = df_global['ANIO'].max()
            
            # Histórico acumulado vs Compras del último año
            hist_compra = df_trabajo.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()
            reciente_compra = df_trabajo[df_trabajo['ANIO'] == max_anio].groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()

            merged = pd.merge(hist_compra, reciente_compra, on=['CLIENTE', 'PRODUCTO'], how='left', suffixes=('_HISTORICO', f'_{max_anio}'))
            merged[f'VALOR_VENTA_{max_anio}'] = merged[f'VALOR_VENTA_{max_anio}'].fillna(0)

            # Dejaron de comprar
            dejaron = merged[merged[f'VALOR_VENTA_{max_anio}'] == 0].copy()
            dejaron = dejaron.sort_values(by=['CLIENTE', 'VALOR_VENTA_HISTORICO'], ascending=[True, False])
            dejaron.columns = ['NOMBRE SN', 'PRODUCTO DEJADO DE COMPRAR', 'VENTA HISTÓRICA ACUMULADA ($)', f'VENTA {max_anio} ($)']

            if not dejaron.empty:
                st.dataframe(dejaron[['NOMBRE SN', 'PRODUCTO DEJADO DE COMPRAR', 'VENTA HISTÓRICA ACUMULADA ($)']], use_container_width=True)
                
                st.markdown("---")
                st.subheader("💡 Sugerido de Venta Recomendado")
                sugerido = dejaron.groupby('NOMBRE SN').first().reset_index()
                sugerido['ACCION COMERCIAL'] = "Reofrecer producto principal histórico que dejó de consumir"
                st.dataframe(sugerido[['NOMBRE SN', 'PRODUCTO DEJADO DE COMPRAR', 'VENTA HISTÓRICA ACUMULADA ($)', 'ACCION COMERCIAL']], use_container_width=True)
            else:
                st.info("Los clientes analizados no registran abandono de productos en el último año.")
        else:
            st.warning("No se encontraron datos para los clientes indicados en la base principal.")

    # -------------------------------------------------------------
    # PESTAÑA 2: INACTIVOS EN 2026
    # -------------------------------------------------------------
    with tab2:
        st.header("Clientes Inactivos en 2026 (Compraron entre 2019-2025 pero 0 en 2026)")

        clientes_2026 = set(df_trabajo[df_trabajo['ANIO'] == 2026]['CLIENTE'].unique())
        df_hist = df_trabajo[(df_trabajo['ANIO'] >= 2019) & (df_trabajo['ANIO'] <= 2025)]
        inactivos = df_hist[~df_hist['CLIENTE'].isin(clientes_2026)]

        if not inactivos.empty:
            top_inactivos = inactivos.groupby('CLIENTE')['VALOR_VENTA'].sum().reset_index()
            top_inactivos = top_inactivos.sort_values(by='VALOR_VENTA', ascending=False).head(20)
            top_inactivos.columns = ['NOMBRE SN', 'VENTA ACUMULADA 2019-2025 ($)']

            st.dataframe(top_inactivos, use_container_width=True)
        else:
            st.info("No hay clientes inactivos en 2026 dentro del grupo seleccionado.")

    # -------------------------------------------------------------
    # PESTAÑA 3: SECTOR ALIMENTACIÓN ANIMAL
    # -------------------------------------------------------------
    with tab3:
        st.header("🌾 Sector Alimentación Animal (Sobre selección actual)")

        patron = '|'.join(PALABRAS_CLAVE_ANIMAL)
        es_cliente = df_trabajo['CLIENTE'].str.contains(patron, case=False, na=False)
        es_prod = df_trabajo['PRODUCTO'].str.contains(patron, case=False, na=False)

        df_animal = df_trabajo[es_cliente | es_prod]

        if not df_animal.empty:
            empresas_animal = df_animal.groupby('CLIENTE')['VALOR_VENTA'].sum().reset_index()
            empresas_animal = empresas_animal.sort_values(by='VALOR_VENTA', ascending=False).head(10)
            empresas_animal.columns = ['NOMBRE SN', 'VENTAS TOTALES ($)']

            st.dataframe(empresas_animal, use_container_width=True)
        else:
            st.info("No se detectaron empresas o productos de alimentación animal en la selección actual.")

else:
    st.info("👋 Por favor, suba el archivo Excel general de ventas para comenzar.")
