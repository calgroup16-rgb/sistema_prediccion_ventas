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
    
    # Estandarización de nombres de columnas a mayúsculas
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    # Intento de identificación de columnas principales
    col_cliente = next((c for c in df.columns if 'CLIENTE' in c or 'NOMBRE' in c), None)
    col_fecha = next((c for c in df.columns if 'FECHA' in c or 'ANIO' in c or 'AÑO' in c), None)
    col_valor = next((c for c in df.columns if 'VALOR' in c or 'VENTA' in c or 'MONTO' in c or 'TOTAL' in c), None)
    col_prod = next((c for c in df.columns if 'PRODUCTO' in c or 'ITEM' in c or 'DESCRIPCION' in c), None)

    if not all([col_cliente, col_fecha, col_valor, col_prod]):
        st.error("El archivo debe contener columnas de Cliente, Fecha, Valor Venta y Producto.")
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
st.markdown("Analice clientes, prediga reenganches, identifique oportunidades y revise clientes por sector.")

# --- BARRA LATERAL: CARGA DE ARCHIVOS Y FILTROS ---
st.sidebar.header("📂 Carga de Archivos")

archivo_ventas = st.sidebar.file_uploader("1. Archivo General de Ventas (Excel)", type=["xlsx", "xls"])

if archivo_ventas is not None:
    df = cargar_y_preparar_datos(archivo_ventas)

    if df is not None:
        st.sidebar.markdown("---")
        st.sidebar.header("🎯 Selección / Listado de Clientes")

        archivo_lista_clientes = st.sidebar.file_uploader(
            "2. Listado Específico de Clientes (Excel - Opcional)", 
            type=["xlsx", "xls"]
        )

        todos_clientes = sorted(df['CLIENTE'].unique().tolist())
        clientes_manuales = st.sidebar.multiselect(
            "O seleccione cliente(s) manualmente:",
            options=todos_clientes,
            placeholder="Buscar uno o varios clientes..."
        )

        # Definir la lista activa de clientes seleccionados
        clientes_objetivo = []
        if archivo_lista_clientes is not None:
            df_lista = pd.read_excel(archivo_lista_clientes)
            col_target = df_lista.columns[0]
            clientes_objetivo = df_lista[col_target].dropna().astype(str).str.strip().unique().tolist()
            st.sidebar.success(f"Cargados {len(clientes_objetivo)} clientes desde la lista.")
        elif clientes_manuales:
            clientes_objetivo = clientes_manuales

        # --- PESTAÑAS PRINCIPALES ---
        tab1, tab2, tab3 = st.tabs([
            "👥 Análisis por Cliente(s)", 
            "⚠️ Top 20 Inactivos 2026 & Recomendación", 
            "🌾 Top 10 Alimentación Animal"
        ])

        # -------------------------------------------------------------
        # PESTAÑA 1: ANÁLISIS INDIVIDUAL / COLECTIVO
        # -------------------------------------------------------------
        with tab1:
            st.header("Análisis Personalizado de Clientes")
            if clientes_objetivo:
                df_filtrado = df[df['CLIENTE'].isin(clientes_objetivo)]

                if not df_filtrado.empty:
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Clientes Analizados", len(clientes_objetivo))
                    m2.metric("Ventas Totales", f"${df_filtrado['VALOR_VENTA'].sum():,.2f}")
                    m3.metric("Total Registro Compras", f"{len(df_filtrado):,}")

                    st.subheader("Histórico de Compras por Cliente y Producto")
                    resumen_cliente = df_filtrado.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].agg(['sum', 'count']).reset_index()
                    resumen_cliente.columns = ['CLIENTE', 'PRODUCTO', 'TOTAL VENTAS ($)', 'CANTIDAD COMPRAS']
                    st.dataframe(resumen_cliente.sort_values(by=['CLIENTE', 'TOTAL VENTAS ($)'], ascending=[True, False]), use_container_width=True)
                else:
                    st.warning("No se encontraron registros de ventas para la lista de clientes seleccionada.")
            else:
                st.info("👈 Cargue un listado de clientes en Excel o seleccione uno o varios en el menú lateral para ver su desglose.")

        # -------------------------------------------------------------
        # PESTAÑA 2: TOP 20 INACTIVOS 2026 (REPUBLICACIÓN Y PREDICCIÓN)
        # -------------------------------------------------------------
        with tab2:
            st.header("Top 20 Clientes Inactivos en 2026 (Mayor Venta 2019-2025)")
            st.caption("Aplica únicamente para clientes que registraron ventas entre 2019 y 2025 pero NO registraron compras en 2026.")

            # Identificar clientes con compra en 2026
            clientes_2026 = set(df[df['ANIO'] == 2026]['CLIENTE'].unique())
            
            # Filtro 2019 - 2025
            df_historico = df[(df['ANIO'] >= 2019) & (df['ANIO'] <= 2025)]
            
            # Excluir los activos en 2026
            df_inactivos = df_historico[~df_historico['CLIENTE'].isin(clientes_2026)]

            if not df_inactivos.empty:
                # Obtener Top 20 por monto total
                top_20 = df_inactivos.groupby('CLIENTE')['VALOR_VENTA'].sum().nlargest(20).reset_index()
                top_20.columns = ['CLIENTE', 'VENTA ACUMULADA 2019-2025 ($)']

                st.subheader("Listado Top 20 Clientes a Recuperar")
                st.dataframe(top_20, use_container_width=True)

                st.subheader("📦 Productos Sugeridos para Reofrecer")
                df_top20_detalle = df_historico[df_historico['CLIENTE'].isin(top_20['CLIENTE'])]
                
                # Productos más comprados por estos clientes
                recom_prod = df_top20_detalle.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()
                recom_prod = recom_prod.sort_values(by=['CLIENTE', 'VALOR_VENTA'], ascending=[True, False])
                
                # Top 3 productos históricos por cliente
                top_3_prod = recom_prod.groupby('CLIENTE').head(3)
                top_3_prod.columns = ['CLIENTE', 'PRODUCTO HISTÓRICO', 'VALOR HISTÓRICO ($)']
                st.dataframe(top_3_prod, use_container_width=True)
            else:
                st.warning("No se encontraron registros de clientes inactivos para el periodo 2019-2025 en relación a 2026.")

        # -------------------------------------------------------------
        # PESTAÑA 3: SECTOR ALIMENTACIÓN ANIMAL (TOP 10 ACTIVOS)
        # -------------------------------------------------------------
        with tab3:
            st.header("Top 10 Clientes Activos - Sector Alimentación Animal")
            st.caption("Identificación de clientes más fuertes que siguen comprando en 2026 dentro del rubro de alimentación animal.")

            # Filtrar activos en 2026
            df_2026 = df[df['ANIO'] == 2026]

            # Filtrar por palabras clave del sector
            patron_regex = '|'.join(PALABRAS_CLAVE_ANIMAL)
            df_animal = df_2026[df_2026['PRODUCTO'].str.contains(patron_regex, case=False, na=False)]

            if not df_animal.empty:
                top_10_animal = df_animal.groupby('CLIENTE')['VALOR_VENTA'].sum().nlargest(10).reset_index()
                top_10_animal.columns = ['CLIENTE', 'VENTAS 2026 ($)']

                col_a, col_b = st.columns([1, 1])
                with col_a:
                    st.subheader("Top 10 Clientes Más Fuertes (2026)")
                    st.dataframe(top_10_animal, use_container_width=True)

                with col_b:
                    st.subheader("💡 Oportunidad Comercial Sugerida")
                    # Detección de productos de alto consumo para cross-selling
                    df_animal_top = df_animal[df_animal['CLIENTE'].isin(top_10_animal['CLIENTE'])]
                    prod_potenciales = df_animal_top.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()
                    prod_potenciales = prod_potenciales.sort_values(by=['CLIENTE', 'VALOR_VENTA'], ascending=[True, False]).groupby('CLIENTE').head(1)
                    prod_potenciales.columns = ['CLIENTE', 'PRODUCTO PRINCIPAL ACTUAL', 'VENTAS ($)']
                    st.dataframe(prod_potenciales, use_container_width=True)
            else:
                st.info("No se registraron productos del sector alimentación animal para el año 2026 con los términos de búsqueda actuales.")

else:
    st.info("👋 Por favor, suba el archivo Excel de ventas en el menú de la izquierda para comenzar el análisis.")
