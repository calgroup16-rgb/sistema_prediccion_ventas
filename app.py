import streamlit as st
import pandas as pd
import numpy as np

# Configuración de la página
st.set_page_config(
    page_title="Sistema de Predicción Comercial",
    page_icon="📊",
    layout="wide"
)

# Terminology y palabras clave para detectar empresas e insumos de alimentación animal
PALABRAS_CLAVE_ANIMAL = [
    'ALIMENTO', 'BALANCEADO', 'FORRAJE', 'GANADO', 'AVICOLA', 'AVÍCOLA',
    'PORCINO', 'PREMEZCLA', 'HARINA', 'MASCOTAS', 'ACUÍCOLA', 'ACUICOLA',
    'NUTRICION', 'NUTRICIONAL', 'AQUA', 'AGRO', 'ANIMAL', 'VETERINARIA',
    'BOVINO', 'AVICULTURA', 'AVICOLAS'
]

def cargar_y_preparar_datos(file):
    """Carga el Excel de ventas y estandariza las columnas requeridas."""
    df = pd.read_excel(file)
    
    # Estandarización de nombres de columnas
    df.columns = [str(col).strip().upper() for col in df.columns]
    
    # Detección de la columna de Cliente ('NOMBRE SN')
    col_cliente = None
    if 'NOMBRE SN' in df.columns:
        col_cliente = 'NOMBRE SN'
    else:
        col_cliente = next((c for c in df.columns if 'CLIENTE' in c or 'NOMBRE' in c or 'SN' in c), None)

    col_fecha = next((c for c in df.columns if 'FECHA' in c or 'ANIO' in c or 'AÑO' in c), None)
    col_valor = next((c for c in df.columns if 'VALOR' in c or 'VENTA' in c or 'MONTO' in c or 'TOTAL' in c), None)
    col_prod = next((c for c in df.columns if 'PRODUCTO' in c or 'ITEM' in c or 'DESCRIPCION' in c), None)

    if not all([col_cliente, col_fecha, col_valor, col_prod]):
        st.error(f"No se identificaron correctamente las columnas requeridas. Cliente: {col_cliente}, Fecha: {col_fecha}, Valor: {col_valor}, Producto: {col_prod}")
        return None

    # Renombrar columnas a formato interno estándar
    df = df.rename(columns={
        col_cliente: 'CLIENTE',
        col_fecha: 'FECHA',
        col_valor: 'VALOR_VENTA',
        col_prod: 'PRODUCTO'
    })

    # Limpieza de tipos de datos
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
    df['ANIO'] = df['FECHA'].dt.year
    df['VALOR_VENTA'] = pd.to_numeric(df['VALOR_VENTA'], errors='coerce').fillna(0)
    df['CLIENTE'] = df['CLIENTE'].astype(str).str.strip()
    df['PRODUCTO'] = df['PRODUCTO'].astype(str).str.strip()

    return df

st.title("📊 Sistema de Predicción Comercial y Análisis de Ventas")

# --- BARRA LATERAL: BASE DE DATOS Y FILTROS ---
st.sidebar.header("📂 Base de Datos Principal")

archivo_ventas = st.sidebar.file_uploader("1. Subir Excel de Ventas Global", type=["xlsx", "xls"])

if archivo_ventas is not None:
    st.session_state['df_ventas'] = cargar_y_preparar_datos(archivo_ventas)

if 'df_ventas' in st.session_state and st.session_state['df_ventas'] is not None:
    df = st.session_state['df_ventas']

    st.sidebar.markdown("---")
    st.sidebar.header("🎯 Carga de Listado Específico")

    archivo_lista_clientes = st.sidebar.file_uploader(
        "2. Subir Listado de Clientes a Buscar (Excel)", 
        type=["xlsx", "xls"]
    )

    todos_clientes = sorted(df['CLIENTE'].dropna().unique().tolist())
    clientes_manuales = st.sidebar.multiselect(
        "O busque cliente(s) individualmente por 'NOMBRE SN':",
        options=todos_clientes,
        placeholder="Escriba o seleccione uno o varios..."
    )

    # Definir lista activa de clientes objetivo
    clientes_objetivo = []
    if archivo_lista_clientes is not None:
        df_lista = pd.read_excel(archivo_lista_clientes)
        df_lista.columns = [str(col).strip().upper() for col in df_lista.columns]
        col_target = 'NOMBRE SN' if 'NOMBRE SN' in df_lista.columns else df_lista.columns[0]
        clientes_objetivo = df_lista[col_target].dropna().astype(str).str.strip().unique().tolist()
        st.sidebar.success(f"Cargados {len(clientes_objetivo)} clientes de la lista propia.")
    elif clientes_manuales:
        clientes_objetivo = clientes_manuales

    # PESTAÑAS PRINCIPALES
    tab1, tab2, tab3 = st.tabs([
        "📋 Sugeridos & Productos Dejados de Comprar", 
        "⚠️ Top 20 Inactivos 2026", 
        "🌾 Detección Automatizada: Sector Alimentación Animal"
    ])

    # -------------------------------------------------------------
    # PESTAÑA 1: RECOMENDADOR Y DEPADOS DE COMPRAR (LISTADO PROPIO)
    # -------------------------------------------------------------
    with tab1:
        if clientes_objetivo:
            st.header("Análisis Comercial y Sugerido de Venta para Listado Cargar")
            df_filtrado = df[df['CLIENTE'].isin(clientes_objetivo)]

            if not df_filtrado.empty:
                m1, m2, m3 = st.columns(3)
                m1.metric("Clientes Encontrados", df_filtrado['CLIENTE'].nunique())
                m2.metric("Ventas Totales ($)", f"${df_filtrado['VALOR_VENTA'].sum():,.2f}")
                m3.metric("Registros de Venta", f"{len(df_filtrado):,}")

                st.markdown("---")
                st.subheader("⚠️ Productos que DEJARON DE COMPRAR (Oportunidad de Reenganche)")
                
                # Identificar productos comprados históricamente vs el último año disponible (2026)
                max_anio = df['ANIO'].max()
                
                compra_historica = df_filtrado.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()
                compra_reciente = df_filtrado[df_filtrado['ANIO'] == max_anio].groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()

                # Merge para identificar abandono de producto
                df_comparativo = pd.merge(compra_historica, compra_reciente, on=['CLIENTE', 'PRODUCTO'], how='left', suffixes=('_HISTORICO', f'_{max_anio}'))
                df_comparativo[f'VALOR_VENTA_{max_anio}'] = df_comparativo[f'VALOR_VENTA_{max_anio}'].fillna(0)

                # Dejaron de comprar: Compraron antes pero tienen 0 en el último año
                dejaron_de_comprar = df_comparativo[df_comparativo[f'VALOR_VENTA_{max_anio}'] == 0].copy()
                dejaron_de_comprar = dejaron_de_comprar.sort_values(by=['CLIENTE', 'VALOR_VENTA_HISTORICO'], ascending=[True, False])
                dejaron_de_comprar.columns = ['NOMBRE SN', 'PRODUCTO DEJADO DE COMPRAR', 'VENTA HISTÓRICA ACUMULADA ($)', f'VENTA {max_anio} ($)']

                if not dejaron_de_comprar.empty:
                    st.dataframe(dejaron_de_comprar, use_container_width=True)
                else:
                    st.info("Todos los productos comprados históricamente por estos clientes registran actividad en el período reciente.")

                st.markdown("---")
                st.subheader("💡 Sugerido de Venta / Productos a Ofrecer")
                
                # Sugerido basado en el top de productos históricos con mayor volumen que dejaron de comprar
                sugeridos = dejaron_de_comprar.groupby('NOMBRE SN').first().reset_index()
                sugeridos['ACCION SUGERIDA'] = "Reofrecer e incentivar venta del producto top histórico"
                st.dataframe(sugeridos[['NOMBRE SN', 'PRODUCTO DEJADO DE COMPRAR', 'VENTA HISTÓRICA ACUMULADA ($)', 'ACCION SUGERIDA']], use_container_width=True)

            else:
                st.warning("Ninguno de los clientes del listado cargado coincide con la base de datos global.")
        else:
            st.info("👈 Por favor, suba una lista de clientes o seleccione nombres en el menú de la izquierda para ver qué dejaron de comprar y qué sugerirles.")

    # -------------------------------------------------------------
    # PESTAÑA 2: TOP 20 INACTIVOS 2026
    # -------------------------------------------------------------
    with tab2:
        st.header("Top 20 Clientes Inactivos en 2026 (Mayor Venta 2019-2025)")
        st.caption("Filtro global automático para clientes que no registraron compras en el año 2026.")

        clientes_2026 = set(df[df['ANIO'] == 2026]['CLIENTE'].unique())
        df_historico = df[(df['ANIO'] >= 2019) & (df['ANIO'] <= 2025)]
        df_inactivos = df_historico[~df_historico['CLIENTE'].isin(clientes_2026)]

        if not df_inactivos.empty:
            top_20 = df_inactivos.groupby('CLIENTE')['VALOR_VENTA'].sum().nlargest(20).reset_index()
            top_20.columns = ['NOMBRE SN', 'VENTA ACUMULADA 2019-2025 ($)']

            st.dataframe(top_20, use_container_width=True)

            st.subheader("📦 Productos Clave a Reofrecer")
            df_top20_detalle = df_historico[df_historico['CLIENTE'].isin(top_20['NOMBRE SN'])]
            
            recom_prod = df_top20_detalle.groupby(['CLIENTE', 'PRODUCTO'])['VALOR_VENTA'].sum().reset_index()
            top_3_prod = recom_prod.sort_values(by=['CLIENTE', 'VALOR_VENTA'], ascending=[True, False]).groupby('CLIENTE').head(3)
            top_3_prod.columns = ['NOMBRE SN', 'PRODUCTO HISTÓRICO', 'VALOR HISTÓRICO ($)']
            st.dataframe(top_3_prod, use_container_width=True)
        else:
            st.warning("No se encontraron registros bajo las condiciones de inactividad indicadas.")

    # -------------------------------------------------------------
    # PESTAÑA 3: AUTOMATIZACIÓN - SECTOR ALIMENTACIÓN ANIMAL
    # -------------------------------------------------------------
    with tab3:
        st.header("🌾 Identificación Automática: Clientes Sector Alimentación Animal")
        st.caption("El sistema escanea automáticamente los nombres de empresa ('NOMBRE SN') y productos comprados para determinar si pertenecen a la industria de nutrición/alimentación animal y cruza cuáles insumos portafolio podemos ofrecerles.")

        patron_regex = '|'.join(PALABRAS_CLAVE_ANIMAL)
        
        # Filtro por nombre de empresa o por descripción de productos
        es_cliente_animal = df['CLIENTE'].str.contains(patron_regex, case=False, na=False)
        es_producto_animal = df['PRODUCTO'].str.contains(patron_regex, case=False, na=False)
        
        df_sector_animal = df[es_cliente_animal | es_producto_animal].copy()

        if not df_sector_animal.empty:
            top10_empresas_animal = df_sector_animal.groupby('CLIENTE')['VALOR_VENTA'].sum().nlargest(10).reset_index()
            top10_empresas_animal.columns = ['NOMBRE SN', 'VENTAS TOTALES ACUMULADAS ($)']

            st.subheader("Top 10 Empresas Detectadas del Sector Alimentación Animal")
            st.dataframe(top10_empresas_animal, use_container_width=True)

            st.markdown("---")
            st.subheader("💡 Insumos Compatibles / Oportunidad de Cross-Selling")
            
            # Identificar qué insumos de nutrición animal del portafolio NO han comprado todavía
            insumos_globales_animal = df[es_producto_animal]['PRODUCTO'].unique()
            
            sugerencias_insumos = []
            for cliente in top10_empresas_animal['NOMBRE SN']:
                prods_comprados = set(df[df['CLIENTE'] == cliente]['PRODUCTO'].unique())
                prods_no_comprados = [p for p in insumos_globales_animal if p not in prods_comprados]
                
                insumo_sugerido = prods_no_comprados[0] if prods_no_comprados else "Portafolio del sector cubierto"
                sugerencias_insumos.append({
                    'NOMBRE SN': cliente,
                    'SECTOR DETECTADO': 'Alimentación / Nutrición Animal',
                    'INSUMO / PRODUCTO COMPATIBLE SUGERIDO': insumo_sugerido
                })

            df_sugerencias = pd.DataFrame(sugerencias_insumos)
            st.dataframe(df_sugerencias, use_container_width=True)

        else:
            st.info("No se detectaron empresas ni productos relacionados con el sector de alimentación animal en el archivo cargado.")

else:
    st.info("👋 Por favor, suba el archivo Excel general de ventas en el menú de la izquierda para desplegar la información.")
