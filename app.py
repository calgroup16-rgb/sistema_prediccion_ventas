import streamlit as st
import pandas as pd
import numpy as np

# Configuración de página
st.set_page_config(
    page_title="Sistema de Predicción Comercial",
    page_icon="📊",
    layout="wide"
)

# Palabras clave para detectar sector de alimentación animal
PALABRAS_CLAVE_ANIMAL = [
    'ALIMENTO', 'BALANCEADO', 'FORRAJE', 'GANADO', 'AVICOLA', 
    'PORCINO', 'PREMEZCLA', 'HARINA', 'MASCOTAS', 'ACUÍCOLA', 'ACUICOLA'
]

def cargar_y_preparar_datos(file):
    """Carga el Excel de ventas y estandariza las columnas requeridas."""
    df = pd.read_excel(file)
    
    # Estandarización de nombres de columnas (limpieza de espacios y mayúsculas)
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    # Detección específica de la columna NOMBRE SN o variantes de cliente
    col_cliente = None
    if 'NOMBRE SN' in df.columns:
        col_cliente = 'NOMBRE SN'
    else:
        col_cliente = next((c for c in df.columns if 'CLIENTE' in c or 'NOMBRE' in c or 'SN' in c), None)

    col_fecha = next((c for c in df.columns if 'FECHA' in c or 'ANIO' in c or 'AÑO' in c), None)
    col_valor = next((c for c in df.columns if 'VALOR' in c or 'VENTA' in c or 'MONTO' in c or 'TOTAL' in c), None)
    col_prod = next((c for c in df.columns if 'PRODUCTO' in c or 'ITEM' in c or 'DESCRIPCION' in c), None)

    if not all([col_cliente, col_fecha, col_valor, col_prod]):
        st.error(f"No se identificaron correctamente todas las columnas. Columna de cliente detectada: {col_cliente}")
        return None

    # Renombrar columnas
    df = df.rename(columns={
        col_cliente: 'CLIENTE',
        col_fecha: 'FECHA',
        col_valor: 'VALOR_VENTA',
        col_prod: 'PRODUCTO'
    })

    # Limpieza de fechas y números
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df['ANIO'] = df['FECHA'].dt.year
    df['VALOR_VENTA'] = pd.to_numeric(df['VALOR_VENTA'], errors='coerce').fillna(0)
    df['CLIENTE'] = df['CLIENTE'].astype(str).str.strip()
    df['PRODUCTO'] = df['PRODUCTO'].astype(str).str.strip()

    return df

st.title("📊 Sistema de Predicción Comercial y Análisis de Ventas")

# --- BARRA LATERAL: CARGA DE ARCHIVOS Y FILTROS ---
st.sidebar.header("📂 Carga de Archivos")

# 1. ARCHIVO PRINCIPAL (Persiste en Session State)
archivo_ventas = st.sidebar.file_uploader("1. Archivo General de Ventas (Excel)", type=["xlsx", "xls"])

if archivo_ventas is not None:
    # Cargar y guardar en la sesión de Streamlit para que no se pierda al interactuar
    st.session_state['df_ventas'] = cargar_y_preparar_datos(archivo_ventas)

if 'df_ventas' in st.session_state and st.session_state['df_ventas'] is not None:
    df = st.session_state['df_ventas']

    st.sidebar.markdown("---")
    st.sidebar.header("🎯 Filtro Personalizado de Clientes")

    # 2. LISTADO SECUNDARIO DE CLIENTES ESPECÍFICOS
    archivo_lista_clientes = st.sidebar.file_uploader(
        "2. Cargar Listado Específico (Búsqueda Propia)", 
        type=["xlsx", "xls"]
    )

    todos_clientes = sorted(df['CLIENTE'].unique().tolist())
    clientes_manuales = st.sidebar.multiselect(
        "O seleccione cliente(s) del archivo global:",
        options=todos_clientes,
        placeholder="Buscar cliente por NOMBRE SN..."
    )

    # Determinar qué clientes se van a consultar sobre la base de datos principal
    clientes_objetivo = []
    if archivo_lista_clientes is not None:
        df_lista = pd.read_excel(archivo_lista_clientes)
        df_lista.columns = [str(col).strip().upper() for col in df_lista.columns]
        
        # Buscar 'NOMBRE SN' o usar la primera columna del archivo cargado
        col_target = 'NOMBRE SN' if 'NOMBRE SN' in df_lista.columns else df_lista.columns[0]
        clientes_objetivo = df_lista[col_target].dropna().astype(str).str.strip().unique().tolist()
        st.sidebar.success(f"Buscando {len(clientes_objetivo)} clientes específicos sobre la base global.")
    elif clientes_manuales:
        clientes_objetivo = clientes_manuales

    # --- PESTAÑAS PRINCIPALES ---
    tab1, tab2, tab3 = st.tabs([
        "👥 Análisis de Clientes (Global / Específico)", 
        "⚠️ Top 20 Inactivos 2026 & Recomendación", 
        "🌾 Top 10 Alimentación Animal"
    ])

    # -------------------------------------------------------------
    # PESTAÑA 1: ANÁLISIS GLOBAL O POR LISTA ESPECÍFICA
    # -------------------------------------------------------------
    with tab1:
        if clientes_objetivo:
            st.header("Análisis de Clientes Filtrados (Lista Propia / Selección)")
            df_filtrado = df[df['CLIENTE'].isin(clientes_objetivo)]

            if not df_filtrado.empty:
                m1, m2, m3 = st.columns(3)
                m1.metric("Clientes Encontrados", df_filtrado['CLIENTE'].nunique())
                m2.metric("Ventas Totales ($)", f"${df_filtrado['VALOR_VENTA'].sum():,.2f}")
                m3.metric("Total Transacciones", f"{len(df_filtrado):,}")

                st.subheader("Histórico de Compras por Cliente y Producto")
                resumen_cliente = df_filtrado.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].agg(['sum', 'count']).reset_index()
                resumen_cliente.columns = ['NOMBRE SN', 'PRODUCTO', 'TOTAL VENTAS ($)', 'CANTIDAD COMPRAS']
                st.dataframe(resumen_cliente.sort_values(by=['NOMBRE SN', 'TOTAL VENTAS ($)'], ascending=[True, False]), use_container_width=True)
            else:
                st.warning("No se encontraron coincidencias de estos clientes en la base de datos principal.")
        else:
            st.header("Análisis Global de Toda la Base de Datos")
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Clientes", df['CLIENTE'].nunique())
            m2.metric("Ventas Totales Globales", f"${df['VALOR_VENTA'].sum():,.2f}")
            m3.metric("Registros de Ventas", f"{len(df):,}")

            st.subheader("Top Clientes Globales por Volumen de Compra")
            resumen_global = df.groupby('CLIENTE')['VALOR_VENTA'].sum().nlargest(15).reset_index()
            resumen_global.columns = ['NOMBRE SN', 'TOTAL VENTAS ACUMULADAS ($)']
            st.dataframe(resumen_global, use_container_width=True)

    # -------------------------------------------------------------
    # PESTAÑA 2: TOP 20 INACTIVOS 2026
    # -------------------------------------------------------------
    with tab2:
        st.header("Top 20 Clientes Inactivos en 2026 (Mayor Venta 2019-2025)")
        st.caption("Filtro global automático: Clientes que compraron entre 2019 y 2025 pero NO tienen compras registradas en 2026.")

        clientes_2026 = set(df[df['ANIO'] == 2026]['CLIENTE'].unique())
        df_historico = df[(df['ANIO'] >= 2019) & (df['ANIO'] <= 2025)]
        df_inactivos = df_historico[~df_historico['CLIENTE'].isin(clientes_2026)]

        if not df_inactivos.empty:
            top_20 = df_inactivos.groupby('CLIENTE')['VALOR_VENTA'].sum().nlargest(20).reset_index()
            top_20.columns = ['NOMBRE SN', 'VENTA ACUMULADA 2019-2025 ($)']

            st.subheader("Listado Top 20 Clientes a Recuperar")
            st.dataframe(top_20, use_container_width=True)

            st.subheader("📦 Productos Sugeridos para Reofrecer (Según Historial)")
            df_top20_detalle = df_historico[df_historico['CLIENTE'].isin(top_20['NOMBRE SN'])]
            
            recom_prod = df_top20_detalle.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()
            top_3_prod = recom_prod.sort_values(by=['CLIENTE', 'VALOR_VENTA'], ascending=[True, False]).groupby('CLIENTE').head(3)
            top_3_prod.columns = ['NOMBRE SN', 'PRODUCTO HISTÓRICO', 'VALOR HISTÓRICO ($)']
            st.dataframe(top_3_prod, use_container_width=True)
        else:
            st.warning("No se encontraron registros de clientes inactivos para el periodo seleccionado.")

    # -------------------------------------------------------------
    # PESTAÑA 3: SECTOR ALIMENTACIÓN ANIMAL
    # -------------------------------------------------------------
    with tab3:
        st.header("Top 10 Clientes Activos - Sector Alimentación Animal")
        st.caption("Filtro de clientes con compras activas en 2026 orientados a productos del sector nutrición/alimentación animal.")

        df_2026 = df[df['ANIO'] == 2026]
        patron_regex = '|'.join(PALABRAS_CLAVE_ANIMAL)
        df_animal = df_2026[df_2026['PRODUCTO'].str.contains(patron_regex, case=False, na=False)]

        if not df_animal.empty:
            top_10_animal = df_animal.groupby('CLIENTE')['VALOR_VENTA'].sum().nlargest(10).reset_index()
            top_10_animal.columns = ['NOMBRE SN', 'VENTAS 2026 ($)']

            col_a, col_b = st.columns([1, 1])
            with col_a:
                st.subheader("Top 10 Clientes Activos (2026)")
                st.dataframe(top_10_animal, use_container_width=True)

            with col_b:
                st.subheader("💡 Oportunidad Comercial Sugerida")
                df_animal_top = df_animal[df_animal['CLIENTE'].isin(top_10_animal['NOMBRE SN'])]
                prod_potenciales = df_animal_top.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()
                prod_potenciales = prod_potenciales.sort_values(by=['CLIENTE', 'VALOR_VENTA'], ascending=[True, False]).groupby('CLIENTE').head(1)
                prod_potenciales.columns = ['NOMBRE SN', 'PRODUCTO PRINCIPAL ACTUAL', 'VENTAS ($)']
                st.dataframe(prod_potenciales, use_container_width=True)
        else:
            st.info("No se registraron transacciones del sector alimentación animal en el año 2026.")

else:
    st.info("👋 Por favor, suba el archivo Excel general de ventas en el menú de la izquierda para comenzar.")
