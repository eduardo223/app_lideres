import streamlit as st
import pandas as pd
import io
import os
import importlib
import procesador

# Recargamos dinámicamente procesador para garantizar que cualquier cambio en procesador.py se aplique sin reiniciar el servidor
importlib.reload(procesador)
from procesador import (
    calcular_metas_ciclo,
    generar_analisis_como_vamos,
    procesar_base_tableau_manager,
    guardar_comentario_lider,
    guardar_todos_comentarios,
    color_nivel,
    color_situacion,
    color_deuda_mora,
    actualizar_situacion_comercial_desde_mi_grupo,
    autenticar_usuario,
    cargar_usuarios,
    registrar_o_actualizar_usuario,
    MATRIZ_GANANCIA,
    ETIQUETAS_ACTIVAS,
    ETIQUETAS_FACTURACION,
    calcular_matriz_ganancia,
    obtener_potencializador_saldo,
    obtener_indice_activas,
    obtener_indice_facturacion,
    rotar_y_guardar_nuevo_ciclo,
    color_cumplimiento,
    color_avance,
    color_saldo,
    exportar_excel_con_colores,
    limpiar_numero
)

# 1. Configuración de la página
st.set_page_config(
    page_title="Panel de Control - Estado de Ciclo Líderes",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para mejorar el diseño estético
st.markdown("""
<style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .stMetric {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px 15px;
    }
    .metric-box {
        background-color: #F1F5F9;
        border-left: 4px solid #3B82F6;
        padding: 12px;
        border-radius: 6px;
        margin-bottom: 12px;
    }
</style>
""", unsafe_allow_html=True)

# Funciones de formato
def formato_cop(val):
    num = limpiar_numero(val, 0.0)
    return f"${num:,.0f}".replace(",", ".")

def formato_cop_signo(val):
    num = limpiar_numero(val, 0.0)
    if num == 0:
        return "$0"
    signo = "-" if num < 0 else ""
    return f"{signo}${abs(num):,.0f}".replace(",", ".")

# 2. Función para cargar y procesar los datos con cache de Streamlit
@st.cache_data
def load_and_process_data(ruta_o_buffer='Base para el como vamos.xlsx'):
    if isinstance(ruta_o_buffer, str):
        if not os.path.exists(ruta_o_buffer):
            return None
        return calcular_metas_ciclo(ruta_o_buffer)
    else:
        df_uploaded = pd.read_excel(ruta_o_buffer, sheet_name="Base para el como vamos")
        return calcular_metas_ciclo(df_uploaded)

# --- CONTROL DE SESIÓN Y LOGIN ---
if 'user' not in st.session_state:
    st.session_state['user'] = None

if st.session_state['user'] is None:
    st.markdown("<h2 style='text-align: center; color: #1E293B; margin-top: 1rem;'>🔑 Iniciar Sesión - Dashboard Liderazgo</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748B;'>Ingresa tus credenciales para acceder a la información de tu perfil y equipo.</p>", unsafe_allow_html=True)
    
    col_l1, col_l2, col_l3 = st.columns([1, 1.2, 1])
    with col_l2:
        with st.form("form_login"):
            input_user = st.text_input("👤 Usuario", value="", placeholder="ej. gerente, lider8425, lider7841, asesor")
            input_pass = st.text_input("🔒 Contraseña", type="password", value="")
            btn_login = st.form_submit_button("🚀 Ingresar al Dashboard", type="primary", use_container_width=True)
            
            if btn_login:
                user_auth = autenticar_usuario(input_user, input_pass)
                if user_auth:
                    st.session_state['user'] = user_auth
                    st.success(f"¡Bienvenido(a) {user_auth['nombre']}!")
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos.")
                    
        st.markdown("---")
        with st.expander("💡 Credenciales de Prueba (Demostración por Perfil)", expanded=True):
            st.markdown("""
            - 👑 **Gerente / Admin**: usuario `gerente` / clave `admin123` *(Acceso Total + Admin Usuarios)*
            - 👩‍💼 **Líder (Luz Dary Chacon - Grupo 8425)**: usuario `lider8425` / clave `lider123` *(Exclusivo Grupo 8425)*
            - 👩‍💼 **Líder (Carmenza Roncancio - Grupo 7841)**: usuario `lider7841` / clave `lider123` *(Exclusivo Grupo 7841)*
            - 👤 **Asesora / Consulta**: usuario `asesor` / clave `asesor123` *(Solo Lectura Facturación)*
            """)
    st.stop()

# Usuario logueado activo
current_user = st.session_state.get('user') or {}
user_nombre = current_user.get('nombre', 'Usuario')
user_rol = current_user.get('rol', 'asesor')
user_grupo = str(current_user.get('codigo_grupo', '')).strip() if current_user.get('codigo_grupo') else ""

# 3. BARRA LATERAL (Perfil de Usuario, Logout y Opciones según Rol)
st.sidebar.markdown(f"### 👤 {user_nombre}")
if user_rol == 'gerente':
    st.sidebar.caption("👑 **Rol**: Gerente / Administrador (Acceso Total)")
elif user_rol == 'lider':
    st.sidebar.caption(f"👩‍💼 **Rol**: Líder de Negocio (Grupo `{user_grupo}`)")
else:
    st.sidebar.caption("👤 **Rol**: Asesora / Consulta de Facturación")

if st.sidebar.button("🚪 Cerrar Sesión", type="secondary"):
    st.session_state['user'] = None
    st.rerun()

st.sidebar.markdown("---")

# Opciones de administración solo visibles para Gerente
if user_rol == 'gerente':
    st.sidebar.header("🔄 Rotación de Ciclo (Nuevo)")
    st.sidebar.caption("Sube el Excel del nuevo ciclo para convertir el ciclo actual en el 'Como vamos anterior' automáticamente.")
    
    nuevo_ciclo_file = st.sidebar.file_uploader("Cargar Nuevo Ciclo ('Cómo Vamos')", type=["xlsx"], key="uploader_nuevo_ciclo")
    if nuevo_ciclo_file is not None:
        if st.sidebar.button("🚀 Rotar Ciclo y Actualizar Histórico"):
            try:
                with st.spinner("Rotando hojas y guardando nuevo ciclo..."):
                    rotar_y_guardar_nuevo_ciclo(nuevo_ciclo_file)
                    st.cache_data.clear()
                    st.sidebar.success("✅ ¡Ciclo rotado con éxito! El nuevo ciclo ya es el activo.")
                    st.rerun()
            except PermissionError as pe:
                st.error("⚠️ **Archivo en uso**: El archivo `Base para el como vamos.xlsx` está actualmente abierto en Excel.")
                st.info("💡 **Solución**: Por favor, **cierra el archivo en Microsoft Excel** y vuelve a presionar el botón '🚀 Rotar Ciclo y Actualizar Histórico'.")
            except Exception as ex:
                st.error(f"❌ Ocurrió un error al rotar el ciclo: {ex}")

st.sidebar.markdown("---")
st.sidebar.header("📊 Actualizar Base Tableau Cam")
st.sidebar.caption("Sube la sábana de datos exportada de Tableau (`Base de Datos.xlsx`) para actualizar la base maestra conservando las notas de la Líder.")

base_tableau_file = st.sidebar.file_uploader("Cargar Base Tableau (.xlsx)", type=["xlsx"], key="uploader_base_tableau_sidebar")
if base_tableau_file is not None:
    if st.sidebar.button("💾 Guardar y Actualizar Base Tableau"):
        try:
            with open("Base de Datos.xlsx", "wb") as f:
                f.write(base_tableau_file.getbuffer())
            st.cache_data.clear()
            st.sidebar.success("✅ ¡Base de Tableau actualizada exitosamente!")
            st.rerun()
        except PermissionError:
            st.sidebar.error("⚠️ El archivo `Base de Datos.xlsx` está abierto en Excel. Ciérralo y vuelve a presionar el botón.")
        except Exception as e_tb:
            st.sidebar.error(f"❌ Error al guardar la base: {e_tb}")

st.sidebar.markdown("---")
st.sidebar.header("🔄 Actualizar Sit. Comercial (`mi_grupo.xls`)")
st.sidebar.caption("Sube `mi_grupo.xls` para actualizar la Situación Comercial de cada consultora en la Base de Datos vinculando por Código CB.")

mi_grupo_file = st.sidebar.file_uploader("Cargar mi_grupo (.xls / .xlsx)", type=["xls", "xlsx"], key="uploader_mi_grupo_sidebar")
if mi_grupo_file is not None:
    if st.sidebar.button("⚡ Actualizar Sit. Comercial desde mi_grupo"):
        with st.spinner("Actualizando Situación Comercial..."):
            res_up = actualizar_situacion_comercial_desde_mi_grupo(mi_grupo_file)
            if res_up.get('exito'):
                st.cache_data.clear()
                st.sidebar.success(f"✅ ¡Actualizado! {res_up['coincidencias']} coincidencia(s), {res_up['cambios']} cambio(s) de estado.")
                st.rerun()
            else:
                st.sidebar.error(f"❌ Error: {res_up.get('error')}")

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Recargar Datos Actuales"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")

# Carga de datos
with st.spinner("Cargando y procesando la base de datos..."):
    df_raw = load_and_process_data('Base para el como vamos.xlsx')

if df_raw is None:
    st.error("⚠️ No se encontró el archivo 'Base para el como vamos.xlsx' en la carpeta actual.")
    st.info("Por favor, asegúrate de colocar el archivo Excel 'Base para el como vamos.xlsx' en la carpeta o sube uno nuevo desde la barra lateral.")
    st.stop()

# Copia de trabajo
df = df_raw.copy()

# Header Principal
st.markdown("<div class='main-header'>📈 Panel de Control - Estado de Ciclo Matices</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Gestión de Líderes, Seguimiento de Metas e Indicadores de Crecimiento</div>", unsafe_allow_html=True)

# 3. BARRA LATERAL (Filtros dinámicos)
st.sidebar.header("🔐 Filtros de Control")

# Filtro por Gerencia
col_gerencia = 'Nombre Gerencia' if 'Nombre Gerencia' in df.columns else df.columns[0]
gerencias_disponibles = sorted([str(g) for g in df[col_gerencia].dropna().unique()])
gerencia_seleccionada = st.sidebar.selectbox(
    "🏢 Selecciona la Gerencia",
    options=["Todas"] + gerencias_disponibles,
    index=0
)

df_filtrado = df.copy()
if gerencia_seleccionada != "Todas":
    df_filtrado = df_filtrado[df_filtrado[col_gerencia] == gerencia_seleccionada]

# Filtro dinámico por Sector (según la Gerencia seleccionada)
col_sector = 'Nombre Setor' if 'Nombre Setor' in df_filtrado.columns else ''
if col_sector and col_sector in df_filtrado.columns:
    sectores_disponibles = sorted([str(s) for s in df_filtrado[col_sector].dropna().unique()])
    sector_seleccionado = st.sidebar.selectbox(
        "📍 Selecciona el Sector",
        options=["Todos"] + sectores_disponibles,
        index=0
    )
    if sector_seleccionado != "Todos":
        df_filtrado = df_filtrado[df_filtrado[col_sector] == sector_seleccionado]

# Filtro por Color / Clasificación
if 'Color' in df_filtrado.columns:
    colores_disponibles = sorted([str(c) for c in df_filtrado['Color'].dropna().unique()])
    colores_seleccionados = st.sidebar.multiselect(
        "🎨 Clasificación / Color",
        options=colores_disponibles,
        default=colores_disponibles
    )
    if colores_seleccionados:
        df_filtrado = df_filtrado[df_filtrado['Color'].astype(str).isin(colores_seleccionados)]

# Buscador de Consultora / Líder
busqueda = st.sidebar.text_input("🔍 Buscar por Nombre o Código", "")
if busqueda.strip():
    col_nombre = 'Nombre de consultora' if 'Nombre de consultora' in df_filtrado.columns else ''
    col_codigo = 'Código de consultora' if 'Código de consultora' in df_filtrado.columns else ''
    
    mask = pd.Series(False, index=df_filtrado.index)
    if col_nombre and col_nombre in df_filtrado.columns:
        mask = mask | df_filtrado[col_nombre].astype(str).str.contains(busqueda, case=False, na=False)
    if col_codigo and col_codigo in df_filtrado.columns:
        mask = mask | df_filtrado[col_codigo].astype(str).str.contains(busqueda, case=False, na=False)
    df_filtrado = df_filtrado[mask]

st.sidebar.markdown("---")
st.sidebar.caption(f"📊 Mostrando **{len(df_filtrado)}** de **{len(df)}** registros")

# Asegurar conversión numérica limpia en df_filtrado para evitar sumar strings
columnas_numericas_clave = [
    'Objetivo Facturación', 'Real Facturación', 'Objetivo Activas', 'Real Activas',
    'Ganancia estimada', 'Ganancia_Matriz_COP', 'Potencializador_COP', 'Saldo',
    'Inicios', 'Reinicios', 'Recuperos', 'Disponibles', 'Falta para el 100%', 'Falta para el 110%',
    'Productividad', 'Cumplimiento Facturación', 'Avance % Facturación', 'Cumplimiento Activas'
]
for col in columnas_numericas_clave:
    if col in df_filtrado.columns:
        df_filtrado[col] = df_filtrado[col].apply(lambda v: limpiar_numero(v, 0.0))

# 4. TARJETAS DE KPIS SUPERIORES
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

total_consultoras = len(df_filtrado)

# Activas Reales vs Objetivo
obj_activas = float(df_filtrado['Objetivo Activas'].sum()) if 'Objetivo Activas' in df_filtrado.columns else 0.0
real_activas = float(df_filtrado['Real Activas'].sum()) if 'Real Activas' in df_filtrado.columns else 0.0
cump_activas = (real_activas / obj_activas * 100.0) if obj_activas > 0 else 0.0

# Facturación Real vs Objetivo
obj_fact = float(df_filtrado['Objetivo Facturación'].sum()) if 'Objetivo Facturación' in df_filtrado.columns else 0.0
real_fact = float(df_filtrado['Real Facturación'].sum()) if 'Real Facturación' in df_filtrado.columns else 0.0
cump_fact = (real_fact / obj_fact * 100.0) if obj_fact > 0 else 0.0

# Ganancia Estimada Total LN
ganancia_total = float(df_filtrado['Ganancia estimada'].sum()) if 'Ganancia estimada' in df_filtrado.columns else 0.0

# Inicios y Reinicios
inicios_totales = float(df_filtrado['Inicios'].sum()) if 'Inicios' in df_filtrado.columns else 0.0
reinicios_totales = float(df_filtrado['Reinicios'].sum()) if 'Reinicios' in df_filtrado.columns else 0.0

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)

with kpi1:
    st.metric("👥 Consultoras / Líderes", f"{total_consultoras}")

with kpi2:
    st.metric(
        "👥 Activas (Real / Obj)",
        f"{int(real_activas)}",
        f"↑ {cump_activas:.1f}% Cumplimiento" if obj_activas > 0 else "0.0%"
    )

with kpi3:
    st.metric(
        "💰 Facturación Total",
        f"${real_fact/1e6:.1f}M COP",
        f"↑ {cump_fact:.1f}% Cumplimiento" if obj_fact > 0 else "0.0%"
    )

with kpi4:
    st.metric(
        "💵 Ganancia Estimada Total LN",
        f"${ganancia_total:,.0f}".replace(",", ".")
    )

with kpi5:
    st.metric(
        "🚀 Inicios / Reinicios",
        f"{int(inicios_totales + reinicios_totales)}",
        f"↑ {inicios_totales:.1f} Inicios | {reinicios_totales:.1f} Reinicios"
    )

st.markdown("---")

# 5. CONTENIDO CON PESTAÑAS (TABS)
if user_rol == 'gerente':
    tab_tableau, tab_resumen, tab_ganancia, tab_diagnostico, tab_metas, tab_detalle, tab_exportar, tab_usuarios = st.tabs([
        "📊 Informe Tableau Cam",
        "📊 Resumen & KPIs",
        "💵 Simulador de Ganancia",
        "🔎 Diagnóstico 'Cómo Vamos'",
        "🎯 Metas de Crecimiento (Procesador)",
        "👥 Detalle Completo",
        "📤 Exportar Datos",
        "🔑 Gestión de Usuarios & Roles"
    ])
else:
    tab_tableau, tab_resumen, tab_ganancia, tab_diagnostico, tab_metas, tab_detalle, tab_exportar = st.tabs([
        "📊 Informe Tableau Cam",
        "📊 Resumen & KPIs",
        "💵 Simulador de Ganancia",
        "🔎 Diagnóstico 'Cómo Vamos'",
        "🎯 Metas de Crecimiento (Procesador)",
        "👥 Detalle Completo",
        "📤 Exportar Datos"
    ])

# --- TAB 0: INFORME TABLEAU MANAGER ("INFORME TABLEAU CAM") ---
with tab_tableau:
    st.subheader("📊 Informe Tableau Manager ('Informe Tableau Cam')")
    st.markdown("Automatización de la Base Maestra de Tableau: Carga única, segmentación privada por Líder/Gerencia, seguimiento de Puntos/Deuda/Crédito y notas persistentes.")

    # 1. Cargar la base de Tableau
    df_tableau = procesar_base_tableau_manager('Base de Datos.xlsx')
    
    # Filtrar por rol de Líder en la base de Tableau
    if df_tableau is not None and not df_tableau.empty and user_rol == 'lider' and user_grupo:
        mask_tg = pd.Series(False, index=df_tableau.index)
        if 'Grupo' in df_tableau.columns:
            mask_tg = mask_tg | df_tableau['Grupo'].astype(str).str.contains(user_grupo, case=False, na=False)
        if 'Sector' in df_tableau.columns:
            mask_tg = mask_tg | df_tableau['Sector'].astype(str).str.contains(user_grupo, case=False, na=False)
        if 'Nombre' in df_tableau.columns:
            mask_tg = mask_tg | (df_tableau['Nombre'].astype(str).str.strip().str.lower() == current_user['nombre'].strip().lower())
        df_tableau = df_tableau[mask_tg].copy()
    
    # Subidor de administración solo visible para Gerente
    if user_rol == 'gerente':
        with st.expander("⚙️ Opciones de Administración (Actualizar Base Tableau & Sit. Comercial desde mi_grupo)", expanded=False):
            col_adm1, col_adm2 = st.columns(2)
            with col_adm1:
                st.markdown("###### 📁 1. Actualizar Base Completa Tableau (`Base de Datos.xlsx`)")
                archivo_tableau = st.file_uploader("Selecciona `Base de Datos.xlsx`", type=["xlsx"], key="tableau_uploader")
                if archivo_tableau is not None:
                    try:
                        with open("Base de Datos.xlsx", "wb") as f:
                            f.write(archivo_tableau.getbuffer())
                        st.success("✅ ¡Base de Datos.xlsx actualizada exitosamente!")
                        st.rerun()
                    except Exception as e_up:
                        st.error(f"Error al actualizar la base: {e_up}")

            with col_adm2:
                st.markdown("###### 🔄 2. Actualizar Sit. Comercial desde `mi_grupo.xls`")
                st.caption("Vincula por Código CB y actualiza la Situación Comercial de cada consultora.")
                file_mg = st.file_uploader("Selecciona `mi_grupo.xls`", type=["xls", "xlsx"], key="mi_grupo_uploader_tab")
                
                # Botón directo si ya existe mi_grupo.xls en la carpeta local
                if os.path.exists("mi_grupo.xls"):
                    if st.button("⚡ Actualizar Sit. Comercial desde 'mi_grupo.xls' local", type="secondary"):
                        res_mg = actualizar_situacion_comercial_desde_mi_grupo("mi_grupo.xls")
                        if res_mg.get('exito'):
                            st.success(f"✅ ¡Actualización exitosa! {res_mg['coincidencias']} coincidencia(s), {res_mg['cambios']} cambio(s) de estado.")
                            if res_mg['cambios'] > 0 and res_mg.get('detalles'):
                                st.markdown("##### 📋 Resumen de Asesoras que cambiaron de Estado Comercial:")
                                st.dataframe(pd.DataFrame(res_mg['detalles']), use_container_width=True)
                            st.rerun()
                        else:
                            st.error(f"Error: {res_mg.get('error')}")

                if file_mg is not None:
                    if st.button("🚀 Actualizar Sit. Comercial desde archivo subido", type="primary"):
                        res_mg = actualizar_situacion_comercial_desde_mi_grupo(file_mg)
                        if res_mg.get('exito'):
                            st.success(f"✅ ¡Actualización exitosa! {res_mg['coincidencias']} coincidencia(s), {res_mg['cambios']} cambio(s) de estado.")
                            if res_mg['cambios'] > 0 and res_mg.get('detalles'):
                                st.markdown("##### 📋 Resumen de Asesoras que cambiaron de Estado Comercial:")
                                st.dataframe(pd.DataFrame(res_mg['detalles']), use_container_width=True)
                            st.rerun()
                        else:
                            st.error(f"Error: {res_mg.get('error')}")

    if df_tableau is None or df_tableau.empty:
        st.warning("⚠️ No se encontró la base de datos `Base de Datos.xlsx`. Por favor, sube el archivo desde la barra lateral o el panel de administración superior.")
    else:
        # Filtros de navegación rápida para Tableau Manager
        col_t1, col_t2, col_t3 = st.columns([1, 1, 1])
        
        # Filtro de Gerencia
        gerencias_t = sorted([str(g) for g in df_tableau['Gerencia'].dropna().unique()]) if 'Gerencia' in df_tableau.columns else []
        with col_t1:
            ger_sel_t = st.selectbox("🏢 Gerencia (Tableau)", options=["Todas"] + gerencias_t, key="tab_ger_sel")
        
        df_tab_filt = df_tableau.copy()
        if ger_sel_t != "Todas" and 'Gerencia' in df_tab_filt.columns:
            df_tab_filt = df_tab_filt[df_tab_filt['Gerencia'] == ger_sel_t]
            
        # Filtro de Sector / Líder
        sectores_t = sorted([str(s) for s in df_tab_filt['Sector'].dropna().unique()]) if 'Sector' in df_tab_filt.columns else []
        with col_t2:
            sec_sel_t = st.selectbox("📍 Sector / Líder (Tableau)", options=["Todos"] + sectores_t, key="tab_sec_sel")
            
        if sec_sel_t != "Todos" and 'Sector' in df_tab_filt.columns:
            df_tab_filt = df_tab_filt[df_tab_filt['Sector'] == sec_sel_t]

        # Buscador de Asesora / Consultora
        with col_t3:
            busq_t = st.text_input("🔍 Buscar Asesora (Nombre / Código)", "", key="tab_busq")
            
        if busq_t.strip():
            mask_t = pd.Series(False, index=df_tab_filt.index)
            if 'Nombre' in df_tab_filt.columns:
                mask_t = mask_t | df_tab_filt['Nombre'].astype(str).str.contains(busq_t, case=False, na=False)
            if 'Codigo CB' in df_tab_filt.columns:
                mask_t = mask_t | df_tab_filt['Codigo CB'].astype(str).str.contains(busq_t, case=False, na=False)
            df_tab_filt = df_tab_filt[mask_t]

        # Filtros adicionales por columna en un expansor dedicado
        with st.expander("🔍 Filtros Avanzados por Columna (Sit. Comercial, Nivel, Mora, Pedidos y Notas)", expanded=True):
            fc1, fc2, fc3, fc4, fc5 = st.columns(5)

            # 1. Filtro por Sit. Comercial
            sits_disponibles = sorted([str(s) for s in df_tab_filt['Sit. Comercial'].dropna().unique()]) if 'Sit. Comercial' in df_tab_filt.columns else []
            with fc1:
                sits_sel = st.multiselect("🚦 Sit. Comercial", options=sits_disponibles, default=[], key="filt_sit_com")
            if sits_sel and 'Sit. Comercial' in df_tab_filt.columns:
                df_tab_filt = df_tab_filt[df_tab_filt['Sit. Comercial'].astype(str).isin(sits_sel)]

            # 2. Filtro por Nivel / Color
            colores_tab_disp = sorted([str(c) for c in df_tab_filt['Color'].dropna().unique()]) if 'Color' in df_tab_filt.columns else []
            with fc2:
                colores_tab_sel = st.multiselect("🏆 Nivel / Color", options=colores_tab_disp, default=[], key="filt_color_tab")
            if colores_tab_sel and 'Color' in df_tab_filt.columns:
                df_tab_filt = df_tab_filt[df_tab_filt['Color'].astype(str).isin(colores_tab_sel)]

            # 3. Filtro por Deuda Mora
            with fc3:
                f_mora_opt = st.selectbox(
                    "⚠️ Deuda Mora",
                    options=["Todas", "Con Deuda Mora (> $0)", "Sin Deuda Mora ($0)", "Mora Crítica (> $500k)"],
                    key="filt_mora_opt"
                )
            if 'Deuda Mora' in df_tab_filt.columns:
                if f_mora_opt == "Con Deuda Mora (> $0)":
                    df_tab_filt = df_tab_filt[df_tab_filt['Deuda Mora'] > 0]
                elif f_mora_opt == "Sin Deuda Mora ($0)":
                    df_tab_filt = df_tab_filt[df_tab_filt['Deuda Mora'] <= 0]
                elif f_mora_opt == "Mora Crítica (> $500k)":
                    df_tab_filt = df_tab_filt[df_tab_filt['Deuda Mora'] > 500000]

            # 4. Filtro por Pedidos Pendientes
            with fc4:
                f_ped_opt = st.selectbox(
                    "⌛ Ped. Pendientes",
                    options=["Todos", "Con Pedidos Pendientes (> 0)", "Sin Pedidos Pendientes (0)"],
                    key="filt_ped_opt"
                )
            if 'Ped. Pendientes' in df_tab_filt.columns:
                if f_ped_opt == "Con Pedidos Pendientes (> 0)":
                    df_tab_filt = df_tab_filt[df_tab_filt['Ped. Pendientes'] > 0]
                elif f_ped_opt == "Sin Pedidos Pendientes (0)":
                    df_tab_filt = df_tab_filt[df_tab_filt['Ped. Pendientes'] <= 0]

            # 5. Filtro por Notas / Comentarios
            with fc5:
                f_notas_opt = st.selectbox(
                    "📝 Notas / Comentarios",
                    options=["Todos", "Con Notas / Comentarios", "Sin Notas"],
                    key="filt_notas_opt"
                )
            if 'Comentarios_Lider' in df_tab_filt.columns:
                if f_notas_opt == "Con Notas / Comentarios":
                    df_tab_filt = df_tab_filt[df_tab_filt['Comentarios_Lider'].astype(str).str.strip() != ""]
                elif f_notas_opt == "Sin Notas":
                    df_tab_filt = df_tab_filt[df_tab_filt['Comentarios_Lider'].astype(str).str.strip() == ""]

        # Subpestañas internas dentro de Informe Tableau Cam
        tab_tab_main, tab_tab_pago, tab_tab_niveles = st.tabs([
            "📋 Base Maestra Gestionable",
            "⌛ Aguardando Pago / Pendientes",
            "🎨 Análisis por Nivel & Estado"
        ])

        # --- SUBPESTAÑA 1: BASE MAESTRA GESTIONABLE ---
        with tab_tab_main:
            # Tarjetas de resumen rápido
            tc1, tc2, tc3, tc4, tc5 = st.columns(5)
            with tc1:
                st.metric("👥 Asesoras Activas", f"{len(df_tab_filt)}")
            with tc2:
                tot_pts = df_tab_filt['Pts Acum'].sum() if 'Pts Acum' in df_tab_filt.columns else 0
                st.metric("🏆 Pts Acumulados", f"{int(tot_pts):,}".replace(",", "."))
            with tc3:
                tot_mora = df_tab_filt['Deuda Mora'].sum() if 'Deuda Mora' in df_tab_filt.columns else 0
                st.metric("⚠️ Deuda Mora Total", f"${tot_mora/1e6:.2f}M COP")
            with tc4:
                tot_cred = df_tab_filt['Credito Disponible'].sum() if 'Credito Disponible' in df_tab_filt.columns else 0
                st.metric("💳 Crédito Dispon.", f"${tot_cred/1e6:.2f}M COP")
            with tc5:
                tot_pago = len(df_tab_filt[df_tab_filt['Ped. Pendientes'] > 0]) if 'Ped. Pendientes' in df_tab_filt.columns else 0
                st.metric("⌛ Aguardando Pago", f"{tot_pago} pers.")

            st.markdown("---")

            # Editor de Comentarios en Masa / Guardar Comentarios
            st.markdown("##### 📝 Comentarios y Notas Persistentes de la Líder")
            st.caption("Escribe las notas de gestión por cada asesora. Se guardarán automáticamente por `Codigo CB` y se mantendrán aunque actualices la base con nuevos cortes de Tableau.")

            # Selección de columnas principales para la vista inicial
            cols_deseadas = [
                'Codigo CB', 'Nombre', 'Color', 'Sit. Comercial',
                'Pts Acum', 'Pts Mant', 'Pts Asc',
                'Deuda Total', 'Deuda Mora', 'Credito Total', 'Credito Disponible',
                'Pts Natura', 'Pts AVON', 'Ped. Pendientes', 'Comentarios_Lider'
            ]
            cols_disp_tab = [c for c in cols_deseadas if c in df_tab_filt.columns]
            
            # Excluir columnas duplicadas o redundantes (ej: 'Situación' vs 'Sit. Comercial', 'Pts Acumulados' vs 'Pts Acum')
            cols_desactivar_duplicadas = {
                'Situación', 'Situacion', 'Pts Acumulados', 'Pts Para Mantener',
                'Pts para Ascender ', 'Deuda Mora ', 'Código CB', 'Codigo_CB_key', 'Unnamed: 55'
            }
            cols_adicionales = [
                c for c in df_tab_filt.columns
                if c not in cols_disp_tab and c not in cols_desactivar_duplicadas and not str(c).startswith('Unnamed:')
            ]
            todas_las_columnas = cols_disp_tab + cols_adicionales

            df_edit_view = df_tab_filt[todas_las_columnas].copy()

            # Renombrar columnas clave para la visualización inicial requerida
            nombres_header_exactos = {
                'Codigo CB': 'Código CB',
                'Nombre': 'Asesora / Consultora',
                'Color': 'Nivel / Color',
                'Sit. Comercial': 'Sit. Comercial',
                'Pts Acum': 'Pts Acum',
                'Pts Mant': 'Pts Mant',
                'Pts Asc': 'Pts Asc',
                'Deuda Total': 'Deuda Total',
                'Deuda Mora': 'Deuda Mora',
                'Credito Total': 'Credito Total',
                'Credito Disponible': 'Credito Disponible',
                'Pts Natura': 'Pts Natura',
                'Pts AVON': 'Pts AVON',
                'Ped. Pendientes': 'Ped. Pendientes',
                'Comentarios_Lider': 'Notas / Comentarios Líder'
            }
            df_edit_view = df_edit_view.rename(columns=nombres_header_exactos)

            # Limpiar cualquier flotante residual en todo el DataFrame para eliminar decimales (.000000)
            for c in df_edit_view.columns:
                if pd.api.types.is_float_dtype(df_edit_view[c]):
                    df_edit_view[c] = df_edit_view[c].fillna(0).round().astype('int64')

            # Usar st.data_editor para permitir editar notas directamente en la tabla
            col_config = {}
            for col_name in df_edit_view.columns:
                # Si es una columna de dinero o puntos, formatear como entero sin decimales
                if 'Deuda' in col_name or 'Credito' in col_name or 'Fact.' in col_name:
                    col_config[col_name] = st.column_config.NumberColumn(col_name, format="$%d", disabled=True)
                elif 'Pts' in col_name or 'Ped.' in col_name or 'Ciclos' in col_name:
                    col_config[col_name] = st.column_config.NumberColumn(col_name, format="%d", disabled=True)
                else:
                    col_config[col_name] = st.column_config.TextColumn(str(col_name), disabled=True)

            if "Notas / Comentarios Líder" in df_edit_view.columns:
                col_config["Notas / Comentarios Líder"] = st.column_config.TextColumn("Notas / Comentarios Líder", disabled=False)

            df_edit_styled = df_edit_view.style.map(
                color_nivel, subset=['Nivel / Color'] if 'Nivel / Color' in df_edit_view.columns else []
            ).map(
                color_situacion, subset=['Sit. Comercial'] if 'Sit. Comercial' in df_edit_view.columns else []
            ).map(
                color_deuda_mora, subset=['Deuda Mora'] if 'Deuda Mora' in df_edit_view.columns else []
            )

            # Columnas seleccionadas / visibles por defecto en la vista principal (Imagen 1)
            cols_visibles_iniciales = [
                'Código CB', 'Asesora / Consultora', 'Nivel / Color', 'Sit. Comercial',
                'Pts Acum', 'Pts Mant', 'Pts Asc',
                'Deuda Total', 'Deuda Mora', 'Credito Total', 'Credito Disponible',
                'Pts Natura', 'Pts AVON', 'Ped. Pendientes', 'Notas / Comentarios Líder'
            ]
            cols_visibles_iniciales_disp = [c for c in cols_visibles_iniciales if c in df_edit_view.columns]

            edited_df = st.data_editor(
                df_edit_styled,
                column_config=col_config,
                column_order=cols_visibles_iniciales_disp,
                use_container_width=True,
                hide_index=True,
                key="editor_tabla_tableau"
            )

            # Botón para guardar los cambios de comentarios
            if st.button("💾 Guardar Cambios en Notas de la Líder", type="primary"):
                dict_guardar = {}
                for idx, row in edited_df.iterrows():
                    codigo_key = str(row.get('Código CB', '')).strip()
                    nota_val = str(row.get('Notas / Comentarios Líder', '')).strip()
                    if codigo_key:
                        dict_guardar[codigo_key] = nota_val
                
                if guardar_todos_comentarios(dict_guardar):
                    st.success("✅ ¡Comentarios guardados exitosamente!")
                    st.rerun()

        # --- SUBPESTAÑA 2: AGUARDANDO PAGO / PEDIDOS PENDIENTES ---
        with tab_tab_pago:
            st.markdown("##### ⌛ Consultoras Aguardando Pago / Pedidos Pendientes")
            st.caption("Listado filtrado de asesoras que tienen pedidos retenidos o saldo en mora pendiente por pago.")

            df_pago = df_tab_filt[(df_tab_filt['Ped. Pendientes'] > 0) | (df_tab_filt['Ped. Mora'] > 0) | (df_tab_filt['Deuda Mora'] > 0)].copy()

            if df_pago.empty:
                st.info("🎉 ¡Excelente! No hay asesoras aguardando pago en la selección actual.")
            else:
                cols_pago_show = [c for c in ['Codigo CB', 'Nombre', 'Color', 'Sit. Comercial', 'Deuda Total', 'Deuda Mora', 'Credito Disponible', 'Ped. Pendientes', 'Ped. Mora', 'Comentarios_Lider'] if c in df_pago.columns]
                df_pago_formatted = df_pago[cols_pago_show].copy()

                if 'Deuda Total' in df_pago_formatted.columns:
                    df_pago_formatted['Deuda Total'] = df_pago_formatted['Deuda Total'].apply(formato_cop)
                if 'Deuda Mora' in df_pago_formatted.columns:
                    df_pago_formatted['Deuda Mora'] = df_pago_formatted['Deuda Mora'].apply(formato_cop_signo)
                if 'Credito Disponible' in df_pago_formatted.columns:
                    df_pago_formatted['Credito Disponible'] = df_pago_formatted['Credito Disponible'].apply(formato_cop)

                st.dataframe(
                    df_pago_formatted.style
                    .map(color_nivel, subset=['Color'] if 'Color' in df_pago_formatted.columns else [])
                    .map(color_situacion, subset=['Sit. Comercial'] if 'Sit. Comercial' in df_pago_formatted.columns else [])
                    .map(color_deuda_mora, subset=['Deuda Mora'] if 'Deuda Mora' in df_pago_formatted.columns else []),
                    use_container_width=True
                )

                # Selector para enviar WhatsApp rápido a la consultora que está aguardando pago
                st.markdown("###### 📲 Enviar Recordatorio de Pago por WhatsApp")
                lideres_pago = sorted(df_pago['Nombre'].dropna().astype(str).unique())
                asesora_sel = st.selectbox("👤 Selecciona la Asesora:", options=lideres_pago, key="sel_asesora_pago")
                if asesora_sel:
                    row_p = df_pago[df_pago['Nombre'].astype(str) == asesora_sel].iloc[0]
                    nombre_p = row_p.get('Nombre', '')
                    deuda_tot_p = formato_cop(row_p.get('Deuda Total', 0))
                    deuda_mora_p = formato_cop(row_p.get('Deuda Mora', 0))
                    ped_pend_p = int(limpiar_numero(row_p.get('Ped. Pendientes', 0)))
                    
                    msg_pago = (
                        f"Hola *{nombre_p}*, 👋\n\n"
                        f"Te recordamos que tienes *{ped_pend_p} pedido(s) pendiente(s)* por liberación.\n"
                        f"💰 *Deuda Total:* {deuda_tot_p}\n"
                        f"⚠️ *Deuda en Mora:* {deuda_mora_p}\n\n"
                        f"Por favor realiza el pago lo antes posible para liberar tu pedido. ¡Gracias! ✨"
                    )
                    st.text_area("📋 Mensaje listo para copiar:", msg_pago, height=140)

        # --- SUBPESTAÑA 3: ANÁLISIS POR NIVEL Y ESTADO COMERCIAL ---
        with tab_tab_niveles:
            st.markdown("##### 🎨 Clasificación por Niveles y Estado Comercial")
            col_n1, col_n2 = st.columns(2)

            with col_n1:
                st.markdown("###### 🏆 Distribución por Nivel (`Color`)")
                if 'Color' in df_tab_filt.columns:
                    df_color_group = df_tab_filt.groupby('Color').agg(
                        Cantidad=('Color', 'count'),
                        Total_Pts_Acum=('Pts Acum', 'sum'),
                        Total_Deuda=('Deuda Total', 'sum')
                    ).reset_index()
                    df_color_group['Total_Deuda'] = df_color_group['Total_Deuda'].apply(formato_cop)
                    st.dataframe(df_color_group, use_container_width=True)

            with col_n2:
                st.markdown("###### 📊 Distribución por Situación Comercial")
                if 'Sit. Comercial' in df_tab_filt.columns:
                    df_sit_group = df_tab_filt.groupby('Sit. Comercial').agg(
                        Cantidad=('Sit. Comercial', 'count'),
                        Total_Deuda_Mora=('Deuda Mora', 'sum')
                    ).reset_index()
                    df_sit_group['Total_Deuda_Mora'] = df_sit_group['Total_Deuda_Mora'].apply(formato_cop)
                    st.dataframe(
                        df_sit_group.style
                        .map(color_situacion, subset=['Sit. Comercial'] if 'Sit. Comercial' in df_sit_group.columns else [])
                        .map(color_deuda_mora, subset=['Total_Deuda_Mora'] if 'Total_Deuda_Mora' in df_sit_group.columns else []),
                        use_container_width=True
                    )

st.markdown("---")

# --- TAB 1: RESUMEN Y KPIS ---
with tab_resumen:
    st.subheader("📊 Análisis General del Estado del Ciclo")
    
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        st.markdown("##### 📍 Desempeño por Sector")
        if 'Nombre Setor' in df_filtrado.columns:
            agg_dict = {}
            if 'Código de consultora' in df_filtrado.columns:
                agg_dict['Líderes'] = ('Código de consultora', 'count')
            if 'Real Activas' in df_filtrado.columns:
                agg_dict['Activas_Reales'] = ('Real Activas', 'sum')
            if 'Objetivo Activas' in df_filtrado.columns:
                agg_dict['Objetivo_Activas'] = ('Objetivo Activas', 'sum')
            if 'Real Facturación' in df_filtrado.columns:
                agg_dict['Facturación_Real'] = ('Real Facturación', 'sum')
                
            if agg_dict:
                resumen_sector = df_filtrado.groupby('Nombre Setor').agg(**agg_dict).reset_index()
                if 'Activas_Reales' in resumen_sector.columns and 'Objetivo_Activas' in resumen_sector.columns:
                    resumen_sector['% Cumpl. Activas'] = (resumen_sector['Activas_Reales'] / resumen_sector['Objetivo_Activas'] * 100).round(1).astype(str) + '%'
                if 'Facturación_Real' in resumen_sector.columns:
                    resumen_sector['Facturación_Real'] = resumen_sector['Facturación_Real'].apply(formato_cop)
                st.dataframe(resumen_sector, use_container_width=True)
            else:
                st.write(df_filtrado[['Nombre Setor']].drop_duplicates())
            
    with col_right:
        st.markdown("##### 🎨 Distribución por Clasificación / Color")
        if 'Color' in df_filtrado.columns:
            agg_color = {'Cantidad': ('Color', 'count')}
            if 'Real Activas' in df_filtrado.columns:
                agg_color['Activas'] = ('Real Activas', 'sum')
            if 'Ganancia estimada' in df_filtrado.columns:
                agg_color['Ganancia_Estimada'] = ('Ganancia estimada', 'sum')
                
            resumen_color = df_filtrado.groupby('Color').agg(**agg_color).reset_index()
            if 'Ganancia_Estimada' in resumen_color.columns:
                resumen_color['Ganancia_Estimada'] = resumen_color['Ganancia_Estimada'].apply(formato_cop)
                
            st.dataframe(resumen_color, use_container_width=True)

# --- TAB 2: SIMULADOR DE GANANCIA ---
with tab_ganancia:
    st.subheader("💵 Matriz y Simulador de Ganancia Estimada")
    st.markdown("Cálculo interactivo según las reglas oficiales de la hoja `#Ganancia#` del simulador LN.")
    
    col_mat, col_pot = st.columns([3, 2])
    
    with col_mat:
        st.markdown("##### 1. Matriz de Ganancia por Cumplimiento (% Activas vs % Facturación)")
        df_matriz_view = pd.DataFrame(
            [[f"{val*100:.2f}%" for val in row] for row in MATRIZ_GANANCIA],
            index=ETIQUETAS_ACTIVAS,
            columns=ETIQUETAS_FACTURACION
        )
        st.dataframe(df_matriz_view, use_container_width=True)
        st.caption("⚠️ **Regla de Inicios**: Si los Inicios de la líder son menores a 4, se le descuenta **-0.5%** a la matriz.")
        
    with col_pot:
        st.markdown("##### 2. Potencializador de Ganancia por Saldo")
        df_pot_view = pd.DataFrame({
            "Saldo (Pedidos/Cartera)": ["-4 o menos", "-3 a -2", "-1", "0", "1", "2", "3", "4", "5", "6 a 7", "8 a 9", "10 o más"],
            "% Potencializador": ["-30%", "-25%", "-20%", "-15%", "-5%", "0%", "+5%", "+10%", "+15%", "+20%", "+25%", "+30%"]
        })
        st.dataframe(df_pot_view, use_container_width=True)
        
    st.markdown("---")
    st.markdown("#### 🧮 Simulador Interactivo")
    
    # Calcular valores por defecto dinámicos a partir de la líder o del filtro activo
    if not df_filtrado.empty:
        if len(df_filtrado) == 1:
            row_sel = df_filtrado.iloc[0]
            val_f_real = limpiar_numero(row_sel.get('Real Facturación', 51229798.0))
            val_f_obj = limpiar_numero(row_sel.get('Objetivo Facturación', 48994379.0))
            val_a_real = int(limpiar_numero(row_sel.get('Real Activas', 150)))
            val_a_obj = int(limpiar_numero(row_sel.get('Objetivo Activas', 145)))
            val_inicios = int(limpiar_numero(row_sel.get('Inicios', 7)))
            val_saldo = int(limpiar_numero(row_sel.get('Saldo', 10)))
        else:
            val_f_real = float(df_filtrado['Real Facturación'].mean()) if 'Real Facturación' in df_filtrado.columns else 51229798.0
            val_f_obj = float(df_filtrado['Objetivo Facturación'].mean()) if 'Objetivo Facturación' in df_filtrado.columns else 48994379.0
            val_a_real = int(df_filtrado['Real Activas'].mean()) if 'Real Activas' in df_filtrado.columns else 150
            val_a_obj = int(df_filtrado['Objetivo Activas'].mean()) if 'Objetivo Activas' in df_filtrado.columns else 145
            val_inicios = int(df_filtrado['Inicios'].mean()) if 'Inicios' in df_filtrado.columns else 7
            val_saldo = int(df_filtrado['Saldo'].mean()) if 'Saldo' in df_filtrado.columns else 10
    else:
        val_f_real, val_f_obj, val_a_real, val_a_obj, val_inicios, val_saldo = 51229798.0, 48994379.0, 150, 145, 7, 10

    # Crear identificador único basado en el primer líder o cantidad de registros para forzar actualización al cambiar de líder
    leader_key = str(df_filtrado.iloc[0].get('Nombre de consultora', 'all')) if len(df_filtrado) == 1 else f"all_{len(df_filtrado)}"

    col_s1, col_s2, col_s3 = st.columns(3)
    
    with col_s1:
        sim_fact_real = st.number_input("Facturación Real ($ COP)", value=float(val_f_real), step=100000.0, key=f"sim_f_real_{leader_key}")
        sim_fact_obj = st.number_input("Facturación Objetivo ($ COP)", value=float(val_f_obj), step=100000.0, key=f"sim_f_obj_{leader_key}")
    
    with col_s2:
        sim_act_real = st.number_input("Activas Reales", value=int(val_a_real), step=1, key=f"sim_a_real_{leader_key}")
        sim_act_obj = st.number_input("Activas Objetivo", value=int(val_a_obj), step=1, key=f"sim_a_obj_{leader_key}")
        
    with col_s3:
        sim_inicios = st.number_input("Inicios Reales", value=int(val_inicios), step=1, key=f"sim_ini_{leader_key}")
        sim_saldo = st.number_input("Saldo Pendiente", value=int(val_saldo), step=1, key=f"sim_sal_{leader_key}")
        
    cump_f_sim = (sim_fact_real / sim_fact_obj) if sim_fact_obj > 0 else 0.0
    cump_a_sim = (sim_act_real / sim_act_obj) if sim_act_obj > 0 else 0.0
    
    pct_matriz_sim, idx_a, idx_f = calcular_matriz_ganancia(cump_a_sim, cump_f_sim, sim_inicios)
    ganancia_matriz_sim = sim_fact_real * pct_matriz_sim
    
    pct_pot_sim = obtener_potencializador_saldo(sim_saldo)
    potencializador_sim_cop = ganancia_matriz_sim * pct_pot_sim
    
    ganancia_estimada_sim_total = ganancia_matriz_sim + potencializador_sim_cop
    
    # Tarjetas de Resultados del Simulador
    res1, res2, res3, res4 = st.columns(4)
    with res1:
        st.metric(
            "% Cumpl. Fact. / Activas",
            f"{cump_f_sim*100:.1f}% / {cump_a_sim*100:.1f}%"
        )
    with res2:
        st.metric(
            "% Ganancia Matriz",
            f"{pct_matriz_sim*100:.2f}%",
            "⚠️ -0.5% por Inicios < 4" if sim_inicios < 4 else "↑ Sin penalización"
        )
    with res3:
        st.metric(
            "Potencializador Saldo",
            f"{pct_pot_sim*100:+.0f}%",
            f"↑ {formato_cop_signo(potencializador_sim_cop)}" if potencializador_sim_cop >= 0 else f"↓ {formato_cop_signo(potencializador_sim_cop)}"
        )
    with res4:
        st.metric(
            "Ganancia Estimada Total",
            f"${ganancia_estimada_sim_total:,.0f}".replace(",", ".")
        )

# --- TAB 3: DIAGNÓSTICO 'CÓMO VAMOS' ---
with tab_diagnostico:
    st.subheader("🔎 Tablas Dinámicas 'Cómo Vamos'")
    st.markdown("Generación de tablas dinámicas automatizadas según los requerimientos del negocio.")
    
    diag = generar_analisis_como_vamos(df_filtrado)
    col_lider = 'Nombre de consultora' if 'Nombre de consultora' in df_filtrado.columns else df_filtrado.columns[0]

    # --- 1. TABLA DE FACTURACIÓN (Formato exacto Clery Cuellar + Ganancia Estimada Total) ---
    st.markdown("#### 💰 1. Tabla de Facturación y Cumplimiento (Ordenadas de Mayor a Menor Cumplimiento)")
    
    cols_fact_exactas = [
        col_lider, 'Objetivo Facturación', 'Real Facturación', 'Cumplimiento Facturación',
        'Avance % Facturación', 'Productividad', 'Falta para el 100%', 'Falta para el 110%', 'Real Activas', 'Ganancia estimada'
    ]
    cols_presentes = [c for c in cols_fact_exactas if c in df_filtrado.columns]
    
    # Ordenar por Cumplimiento de Facturación (Top arriba, menos top abajo)
    if 'Cumplimiento Facturación' in df_filtrado.columns:
        df_fact_sorted = df_filtrado.sort_values(by='Cumplimiento Facturación', ascending=False)
    else:
        df_fact_sorted = df_filtrado
        
    df_fact_view = df_fact_sorted[cols_presentes].copy()
    
    # Renombrar exactamente como en la plantilla original de Clery (Screenshot 2)
    nombres_clery = {
        col_lider: 'LÍDER DE NEGOCIOS',
        'Objetivo Facturación': 'DESAFÍO FACTURACIÓN',
        'Real Facturación': 'FACTURACIÓN A HOY',
        'Cumplimiento Facturación': 'CUMPLIMIENTO DE FACTURACIÓN',
        'Avance % Facturación': 'AVANCE %',
        'Productividad': 'PRODUCTIVIDAD',
        'Falta para el 100%': 'FALTA PARA EL 100%',
        'Falta para el 110%': 'CUÁNTO FALTA PARA EL 110%',
        'Real Activas': 'PEDIDOS',
        'Ganancia estimada': 'GANANCIA ESTIMADA TOTAL'
    }
    df_fact_view = df_fact_view.rename(columns=nombres_clery)
    
    # Formatear números para visualización impecable
    df_fact_formatted = df_fact_view.copy()
    if 'DESAFÍO FACTURACIÓN' in df_fact_formatted.columns:
        df_fact_formatted['DESAFÍO FACTURACIÓN'] = df_fact_formatted['DESAFÍO FACTURACIÓN'].apply(formato_cop)
    if 'FACTURACIÓN A HOY' in df_fact_formatted.columns:
        df_fact_formatted['FACTURACIÓN A HOY'] = df_fact_formatted['FACTURACIÓN A HOY'].apply(formato_cop)
    if 'CUMPLIMIENTO DE FACTURACIÓN' in df_fact_formatted.columns:
        df_fact_formatted['CUMPLIMIENTO DE FACTURACIÓN'] = df_fact_formatted['CUMPLIMIENTO DE FACTURACIÓN'].apply(
            lambda v: f"➡️ {limpiar_numero(v):.2f}%" if pd.notna(v) and limpiar_numero(v) > 0 else "0.00%"
        )
    if 'AVANCE %' in df_fact_formatted.columns:
        df_fact_formatted['AVANCE %'] = df_fact_formatted['AVANCE %'].apply(
            lambda v: f"{limpiar_numero(v):+.1f}%" if pd.notna(v) else "0.0%"
        )
    if 'PRODUCTIVIDAD' in df_fact_formatted.columns:
        df_fact_formatted['PRODUCTIVIDAD'] = df_fact_formatted['PRODUCTIVIDAD'].apply(formato_cop)
    if 'FALTA PARA EL 100%' in df_fact_formatted.columns:
        df_fact_formatted['FALTA PARA EL 100%'] = df_fact_formatted['FALTA PARA EL 100%'].apply(formato_cop_signo)
    if 'CUÁNTO FALTA PARA EL 110%' in df_fact_formatted.columns:
        df_fact_formatted['CUÁNTO FALTA PARA EL 110%'] = df_fact_formatted['CUÁNTO FALTA PARA EL 110%'].apply(formato_cop_signo)
    if 'GANANCIA ESTIMADA TOTAL' in df_fact_formatted.columns:
        df_fact_formatted['GANANCIA ESTIMADA TOTAL'] = df_fact_formatted['GANANCIA ESTIMADA TOTAL'].apply(formato_cop)
    if 'PEDIDOS' in df_fact_formatted.columns:
        df_fact_formatted['PEDIDOS'] = df_fact_formatted['PEDIDOS'].apply(lambda v: f"{int(limpiar_numero(v))}")

    # Renderizar con semáforo de colores según comportamiento
    st.dataframe(
        df_fact_formatted.style
        .map(color_cumplimiento, subset=['CUMPLIMIENTO DE FACTURACIÓN'] if 'CUMPLIMIENTO DE FACTURACIÓN' in df_fact_formatted.columns else [])
        .map(color_avance, subset=['AVANCE %'] if 'AVANCE %' in df_fact_formatted.columns else []),
        use_container_width=True
    )

    st.markdown("---")

    # --- 2. TABLA DE ACTIVAS EN ORDEN DE MAYOR A MENOR ---
    st.markdown("#### 👥 2. Tabla de Activas (Ordenadas de Mayor a Menor Desempeño)")
    cols_act = [c for c in [col_lider, 'Nombre Setor', 'Color', 'Real Activas', 'Objetivo Activas', 'Cumplimiento Activas'] if c in df_filtrado.columns]
    
    if 'Cumplimiento Activas' in df_filtrado.columns:
        df_act_sorted = df_filtrado.sort_values(by='Cumplimiento Activas', ascending=False)
    else:
        df_act_sorted = df_filtrado.sort_values(by='Real Activas', ascending=False)
        
    df_activas_order = df_act_sorted[cols_act].copy()
    
    if 'Real Activas' in df_activas_order.columns:
        df_activas_order['Real Activas'] = df_activas_order['Real Activas'].apply(lambda v: f"{int(limpiar_numero(v))}")
    if 'Objetivo Activas' in df_activas_order.columns:
        df_activas_order['Objetivo Activas'] = df_activas_order['Objetivo Activas'].apply(lambda v: f"{int(limpiar_numero(v))}")
        
    # Calcular cumplimiento % real a partir de activas reales / objetivo
    df_activas_order['Cumplimiento Activas'] = df_act_sorted.apply(
        lambda r: f"{(limpiar_numero(r.get('Real Activas', 0)) / limpiar_numero(r.get('Objetivo Activas', 1)) * 100):.1f}%" if limpiar_numero(r.get('Objetivo Activas', 0)) > 0 else "0.0%",
        axis=1
    )
    
    st.dataframe(
        df_activas_order.style.map(color_cumplimiento, subset=['Cumplimiento Activas'] if 'Cumplimiento Activas' in df_activas_order.columns else []),
        use_container_width=True
    )

    st.markdown("---")

    # --- 3. TABLA DE SALDO Y POTENCIALIZADORES ---
    st.markdown("#### ⚠️ 3. Tabla de Saldos y Potencializadores (Ordenadas por Ganancia Estimada)")
    cols_saldo = [c for c in [col_lider, 'Nombre Setor', 'Saldo', 'Potencializador_Pct', 'Ganancia_Matriz_COP', 'Potencializador_COP', 'Ganancia estimada'] if c in df_filtrado.columns]
    
    # Ordenar por Ganancia Estimada de mayor a menor (Top ganadoras arriba)
    if 'Ganancia estimada' in df_filtrado.columns:
        df_saldo_sorted = df_filtrado.sort_values(by='Ganancia estimada', ascending=False)
    else:
        df_saldo_sorted = df_filtrado.sort_values(by='Saldo', ascending=True)
        
    df_saldo_view = df_saldo_sorted[cols_saldo].copy()
    
    if 'Saldo' in df_saldo_view.columns:
        df_saldo_view['Saldo'] = df_saldo_view['Saldo'].apply(lambda v: f"{int(limpiar_numero(v))}")
    if 'Potencializador_Pct' in df_saldo_view.columns:
        df_saldo_view['Potencializador_Pct'] = df_saldo_view['Potencializador_Pct'].apply(lambda v: f"{limpiar_numero(v)*100:+.0f}%" if pd.notna(v) else "0%")
    if 'Ganancia_Matriz_COP' in df_saldo_view.columns:
        df_saldo_view['Ganancia_Matriz_COP'] = df_saldo_view['Ganancia_Matriz_COP'].apply(formato_cop)
    if 'Potencializador_COP' in df_saldo_view.columns:
        df_saldo_view['Potencializador_COP'] = df_saldo_view['Potencializador_COP'].apply(formato_cop_signo)
    if 'Ganancia estimada' in df_saldo_view.columns:
        df_saldo_view['Ganancia estimada'] = df_saldo_view['Ganancia estimada'].apply(formato_cop)
        
    st.dataframe(
        df_saldo_view.style.map(color_saldo, subset=['Saldo'] if 'Saldo' in df_saldo_view.columns else []),
        use_container_width=True
    )

    st.markdown("---")

    # --- 4. TABLA DE DISPONIBLES Y ENTRADAS ---
    st.markdown("#### 🎯 4. Tabla de Disponibles y Entradas")
    cols_disp = [c for c in [col_lider, 'Nombre Setor', 'Disponibles', 'Real Activas', 'Inicios', 'Reinicios', 'Recuperos', 'Inactiva 1', 'Inactiva 2', 'Inactiva 3'] if c in df_filtrado.columns]
    df_disp_view = df_filtrado[cols_disp].sort_values(by='Disponibles', ascending=False).copy()
    
    for c_int in ['Disponibles', 'Real Activas', 'Inicios', 'Reinicios', 'Recuperos', 'Inactiva 1', 'Inactiva 2', 'Inactiva 3']:
        if c_int in df_disp_view.columns:
            df_disp_view[c_int] = df_disp_view[c_int].apply(lambda v: f"{int(limpiar_numero(v))}")
            
    st.dataframe(df_disp_view, use_container_width=True)

    st.markdown("---")

    # --- 5. TABLAS DE RETENCIÓN DE INACTIVAS ---
    st.markdown("#### 🛡️ 5. Tablas de Retención de Inactivas (Límites Máximos de Fuga)")
    st.info("💡 **Reglas de Retención**: Se calculan sobre las disponibles actuales. Muestran el límite máximo de consultoras que se pueden 'dejar ir' y la retención mínima requerida.")
    
    if diag and 'disponibles' in diag and 'retencion' in diag['disponibles']:
        ret_data = diag['disponibles']['retencion']
        df_ret_resumen = pd.DataFrame([
            {
                "Nivel Inactiva": "Inactiva 1",
                "Total Actual": int(ret_data['inactiva_1']['total']),
                "% Máx. Fuga Permitida": "12%",
                "Máx. Fuga (Cant. Pers.)": f"{ret_data['inactiva_1']['max_fuga_cant']:.0f}",
                "% Mín. Retención Requerida": "88%",
                "Mín. Retención (Cant. Pers.)": f"{ret_data['inactiva_1']['min_retencion_cant']:.0f}"
            },
            {
                "Nivel Inactiva": "Inactiva 2",
                "Total Actual": int(ret_data['inactiva_2']['total']),
                "% Máx. Fuga Permitida": "8%",
                "Máx. Fuga (Cant. Pers.)": f"{ret_data['inactiva_2']['max_fuga_cant']:.0f}",
                "% Mín. Retención Requerida": "92%",
                "Mín. Retención (Cant. Pers.)": f"{ret_data['inactiva_2']['min_retencion_cant']:.0f}"
            },
            {
                "Nivel Inactiva": "Inactiva 3",
                "Total Actual": int(ret_data['inactiva_3']['total']),
                "% Máx. Fuga Permitida": "6%",
                "Máx. Fuga (Cant. Pers.)": f"{ret_data['inactiva_3']['max_fuga_cant']:.0f}",
                "% Mín. Retención Requerida": "94%",
                "Mín. Retención (Cant. Pers.)": f"{ret_data['inactiva_3']['min_retencion_cant']:.0f}"
            }
        ])
        st.dataframe(df_ret_resumen, use_container_width=True)

    st.markdown("---")

    # --- 6. MÓDULO DE COMPARTIR POR WHATSAPP ---
    st.markdown("#### 📲 6. Módulo para Compartir Resumen por WhatsApp")
    st.caption("Selecciona una Líder para generar su reporte en formato texto listo para copiar o enviar directamente por WhatsApp Web / Móvil.")
    
    lista_lideres = sorted(df_filtrado[col_lider].dropna().astype(str).unique())
    lider_sel = st.selectbox("👤 Selecciona la Líder para enviar reporte:", options=lista_lideres)
    
    if lider_sel:
        row_l = df_filtrado[df_filtrado[col_lider].astype(str) == lider_sel].iloc[0]
        
        r_fact = formato_cop(row_l.get('Real Facturación', 0))
        o_fact = formato_cop(row_l.get('Objetivo Facturación', 0))
        c_fact = row_l.get('Cumplimiento Facturación', 0)
        c_fact_str = f"{limpiar_numero(c_fact):.2f}%" if pd.notna(c_fact) and limpiar_numero(c_fact) > 0 else "0.00%"
        fal_100 = formato_cop_signo(row_l.get('Falta para el 100%', 0))
        fal_110 = formato_cop_signo(row_l.get('Falta para el 110%', 0))
        act_r = row_l.get('Real Activas', 0)
        sal_l = row_l.get('Saldo', 0)
        gan_l = formato_cop(row_l.get('Ganancia estimada', 0))
        sector_l = row_l.get('Nombre Setor', 'General')
        
        msg_wa = (
            f"📊 *REPORTE CÓMO VAMOS*\n"
            f"👤 *Líder:* {lider_sel}\n"
            f"📍 *Sector:* {sector_l}\n\n"
            f"💰 *Facturación Real:* {r_fact}\n"
            f"🎯 *Objetivo Facturación:* {o_fact}\n"
            f"📈 *Cumplimiento Facturación:* {c_fact_str}\n"
            f"💵 *Falta para 100%:* {fal_100}\n"
            f"🚀 *Falta para 110%:* {fal_110}\n"
            f"👥 *Activas Reales:* {act_r}\n"
            f"⚠️ *Saldo Pendiente:* {sal_l}\n"
            f"💵 *Ganancia Estimada:* {gan_l}\n"
        )
        
        col_w1, col_w2 = st.columns([2, 1])
        with col_w1:
            st.text_area("📋 Mensaje listo para copiar:", msg_wa, height=220)
        with col_w2:
            import urllib.parse
            url_wa = f"https://api.whatsapp.com/send?text={urllib.parse.quote(msg_wa)}"
            st.markdown(f"<br><a href='{url_wa}' target='_blank' style='text-decoration:none;'><button style='background-color:#25D366; color:white; border:none; padding:14px 20px; font-size:16px; font-weight:bold; border-radius:8px; cursor:pointer; width:100%;'>📲 Enviar por WhatsApp</button></a>", unsafe_allow_html=True)


# --- TAB 3: METAS DE CRECIMIENTO ---
with tab_metas:
    st.subheader("🎯 Metas de Crecimiento Integradas (Procesador)")
    st.info("💡 **Reglas de Cálculo**: Las metas representan las activas necesarias para alcanzar cada tramo de incentivo (+1, +3, +5, +7, +9 sobre tus Activas Reales actuales).")
    
    cols_metas = [
        'Nombre de consultora', 'Nombre Setor', 'Real Activas', 'Avance % Facturación', 'Falta para el 100%',
        'Meta_Crecer_1plus_150k', 'Meta_Crecer_3plus_200k', 'Meta_Crecer_5plus_300k',
        'Meta_Crecer_7plus_500k', 'Meta_Crecer_9plus_750k'
    ]
    
    cols_existentes_metas = [c for c in cols_metas if c in df_filtrado.columns]
    
    df_metas_view = df_filtrado[cols_existentes_metas].copy()
    if 'Avance % Facturación' in df_metas_view.columns:
        df_metas_view['Avance % Facturación'] = df_metas_view['Avance % Facturación'].apply(
            lambda v: f"{v*100:+.2f}%" if pd.notna(v) else "N/A"
        )
    if 'Falta para el 100%' in df_metas_view.columns:
        df_metas_view['Falta para el 100%'] = df_metas_view['Falta para el 100%'].apply(formato_cop_signo)
        
    st.dataframe(
        df_metas_view.rename(columns={
            'Meta_Crecer_1plus_150k': 'Meta 1+ (+150k)',
            'Meta_Crecer_3plus_200k': 'Meta 3+ (+200k)',
            'Meta_Crecer_5plus_300k': 'Meta 5+ (+300k)',
            'Meta_Crecer_7plus_500k': 'Meta 7+ (+500k)',
            'Meta_Crecer_9plus_750k': 'Meta 9+ (+750k)',
            'Avance % Facturación': 'Avance % vs Ant.'
        }),
        use_container_width=True
    )

# --- TAB 4: DETALLE COMPLETO ---
with tab_detalle:
    st.subheader("👥 Tabla Completa de Líderes / Consultoras")
    
    # Selector de columnas para personalizar la vista
    columnas_disponibles = list(df_filtrado.columns)
    columnas_predeterminadas = [
        c for c in ['Nombre Gerencia', 'Nombre Setor', 'Código de consultora', 'Nombre de consultora', 'Color', 'Real Activas', 'Objetivo Facturación', 'Real Facturación', 'Cumplimiento Facturación', 'Falta para el 100%', 'Avance % Facturación', 'Ganancia estimada']
        if c in columnas_disponibles
    ]
    
    cols_seleccionadas = st.multiselect(
        "Selecciona las columnas que deseas visualizar:",
        options=columnas_disponibles,
        default=columnas_predeterminadas
    )
    
    if cols_seleccionadas:
        st.dataframe(df_filtrado[cols_seleccionadas], use_container_width=True)
    else:
        st.warning("Selecciona al menos una columna para mostrar.")

# --- TAB 5: EXPORTAR DATOS ---
with tab_exportar:
    st.subheader("📥 Exportar Resultados Procesados")
    st.write("Descarga la base de datos con los filtros actuales o exporta los reportes visuales a colores listos para compartir con las líderes.")
    
    col_exp1, col_exp2, col_exp3 = st.columns(3)
    
    with col_exp1:
        # Excel estándar completo
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name='Metas_Procesadas')
        excel_data = output_excel.getvalue()
        
        st.download_button(
            label="📄 Descargar Excel Completo (.xlsx)",
            data=excel_data,
            file_name="Resultado_Metas_Procesadas_Filtrado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    with col_exp2:
        # Reporte Excel a Colores para Líderes
        excel_colores_bytes = exportar_excel_con_colores({
            'Activas': df_filtrado[[c for c in ['Nombre de consultora', 'Nombre Setor', 'Color', 'Real Activas', 'Objetivo Activas', 'Cumplimiento Activas'] if c in df_filtrado.columns]],
            'Facturacion': df_filtrado[[c for c in ['Nombre de consultora', 'Nombre Setor', 'Real Facturación', 'Objetivo Facturación', 'Cumplimiento Facturación', 'Falta para el 100%'] if c in df_filtrado.columns]],
            'Saldos': df_filtrado[[c for c in ['Nombre de consultora', 'Nombre Setor', 'Saldo', 'Potencializador_Pct', 'Ganancia estimada'] if c in df_filtrado.columns]],
            'Disponibles': df_filtrado[[c for c in ['Nombre de consultora', 'Nombre Setor', 'Disponibles', 'Real Activas', 'Inicios', 'Reinicios', 'Recuperos'] if c in df_filtrado.columns]]
        })
        
        st.download_button(
            label="🎨 Descargar Reporte A COLORES (.xlsx)",
            data=excel_colores_bytes,
            file_name="Reporte_Lideres_Formato_Colores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    with col_exp3:
        # Generar CSV en memoria
        csv_data = df_filtrado.to_csv(index=False).encode('utf-8-sig')
        
        st.download_button(
            label="📊 Descargar CSV (.csv)",
            data=csv_data,
            file_name="Resultado_Metas_Procesadas_Filtrado.csv",
            mime="text/csv"
        )

# --- TAB 6: GESTIÓN DE USUARIOS (SOLO GERENTE) ---
if user_rol == 'gerente':
    with tab_usuarios:
        st.subheader("🔑 Gestión de Usuarios y Permisos (Solo Gerente)")
        st.markdown("Administra las credenciales de acceso, asigna roles y vincula el **Código de grupo** a cada Líder de Negocio.")

        col_u1, col_u2 = st.columns([1.2, 1])

        with col_u1:
            st.markdown("##### 👥 Usuarios Registrados Actuales")
            users_dict = cargar_usuarios()
            list_u = []
            for uname, udata in users_dict.items():
                list_u.append({
                    "Usuario": uname,
                    "Nombre": udata.get("nombre", ""),
                    "Rol": udata.get("rol", ""),
                    "Código de Grupo": udata.get("codigo_grupo") or "N/A"
                })
            st.dataframe(pd.DataFrame(list_u), use_container_width=True)

        with col_u2:
            st.markdown("##### ➕ Crear / Editar Usuario de Líder o Asesora")
            with st.form("form_nuevo_usuario"):
                nu_username = st.text_input("Usuario (Login)", placeholder="ej. lider9334")
                nu_nombre = st.text_input("Nombre Completo", placeholder="ej. Angela Mireya Montenegro")
                nu_pass = st.text_input("Contraseña", type="password", placeholder="Dejar vacío para mantener contraseña actual")
                nu_rol = st.selectbox("Rol de Acceso", options=["lider", "gerente", "asesor"])
                nu_grupo = st.text_input("Código de Grupo (Para Líderes)", placeholder="ej. 9334")
                
                btn_save_u = st.form_submit_button("💾 Guardar / Actualizar Usuario", type="primary", use_container_width=True)
                if btn_save_u:
                    ok_u, msg_u = registrar_o_actualizar_usuario(
                        nu_username, nu_nombre, nu_pass, nu_rol, nu_grupo
                    )
                    if ok_u:
                        st.success(f"✅ {msg_u}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg_u}")

# Footer
st.markdown("---")
st.caption("📈 Panel de Control Matices | Desarrollado con Streamlit & Pandas")