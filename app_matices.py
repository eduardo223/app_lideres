import streamlit as st
import pandas as pd
import urllib.parse
import os

import procesador
from procesador import (
    autenticar_usuario,
    cargar_usuarios,
    cargar_objetivos_arte,
    consultar_tableau_sql,
    consultar_geral_sql,
    calcular_metas_ciclo,
    limpiar_numero,
    cambiar_password_usuario,
    extraer_telefonos_colombia,
    limpiar_y_ordenar_columnas_tableau,
    color_nivel,
    color_situacion,
    color_deuda_mora,
    guardar_todos_comentarios
)

# 1. Configuración de página ultra-optimizada para Celulares, Tablets y Pantallas
st.set_page_config(
    page_title="App Matices - Gestión Líderes",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 2. CSS Ultra-Compacto y Responsivo para Smartphones y Tablets (Tema Natura & Avon)
st.markdown("""
<style>
    /* Reducción radical de márgenes de Streamlit para móviles */
    .block-container {
        padding-top: 0.5rem !important;
        padding-bottom: 2rem !important;
        padding-left: 0.5rem !important;
        padding-right: 0.5rem !important;
        max-width: 100% !important;
    }
    
    /* Pestañas compactas estilo App móvil con gradiente Natura-Avon */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px !important;
        background-color: rgba(227, 0, 123, 0.06);
        border-radius: 14px;
        padding: 4px;
        margin-bottom: 10px;
        border: 1px solid rgba(227, 0, 123, 0.15);
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 11.5px !important;
        font-weight: 700 !important;
        padding: 8px 10px !important;
        border-radius: 10px;
        color: #475569;
        transition: all 0.2s ease;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #FF6B00 0%, #E3007B 100%) !important;
        color: #ffffff !important;
        box-shadow: 0 4px 12px rgba(227, 0, 123, 0.35);
    }

    /* Cards KPI 2x2 y 3x2 Condensadas */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 6px;
        margin-bottom: 8px;
    }
    .kpi-grid-3 {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 6px;
        margin-bottom: 8px;
    }
    .kpi-card {
        background: #ffffff;
        border: 1px solid #fbcfe8;
        border-radius: 12px;
        padding: 8px 10px;
        text-align: center;
        box-shadow: 0 2px 6px rgba(227, 0, 123, 0.05);
    }
    .kpi-title {
        font-size: 10px;
        color: #64748b;
        text-transform: uppercase;
        font-weight: 700;
        margin-bottom: 2px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    .kpi-val {
        font-size: 15px;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.2;
    }
    .kpi-val-gradient {
        font-size: 15px;
        font-weight: 800;
        background: linear-gradient(135deg, #FF6B00 0%, #E3007B 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        line-height: 1.2;
    }
    .kpi-sub {
        font-size: 9.5px;
        font-weight: 700;
        margin-top: 3px;
        padding: 2px 4px;
        border-radius: 6px;
        display: inline-block;
    }
    .kpi-sub-green {
        background: #dcfce7;
        color: #15803d;
    }
    .kpi-sub-red {
        background: #fee2e2;
        color: #b91c1c;
    }
    .kpi-sub-blue {
        background: #e0f2fe;
        color: #0369a1;
    }
    .kpi-sub-orange {
        background: #ffedd5;
        color: #c2410c;
    }

    /* Botón WhatsApp Móvil */
    .btn-wa-link {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        background-color: #25D366;
        color: white !important;
        font-size: 11px;
        font-weight: 700;
        padding: 4px 9px;
        border-radius: 8px;
        text-decoration: none !important;
        box-shadow: 0 2px 6px rgba(37, 211, 102, 0.3);
    }

    /* Badges de estado */
    .badge-pill {
        display: inline-block;
        font-size: 10px;
        font-weight: 700;
        padding: 2px 6px;
        border-radius: 9999px;
    }
</style>
""", unsafe_allow_html=True)

# 3. Control de Sesión y Login Móvil
if 'user' not in st.session_state:
    st.session_state['user'] = None

# Pantalla de Login Móvil si no está autenticado
if st.session_state['user'] is None:
    st.markdown("<h3 style='text-align:center; margin-bottom:2px;'>📱 App Matices Móvil</h3>", unsafe_allow_html=True)
    st.caption("<p style='text-align:center; margin-bottom:12px;'>Portal exclusivo para Líderes de Negocio en Celulares y Tablets</p>", unsafe_allow_html=True)
    
    with st.form("form_login_mobile"):
        input_u = st.text_input("👤 Usuario o Correo", placeholder="Ej. lider7841 o correo@...").strip().lower()
        input_p = st.text_input("🔒 Contraseña", type="password", placeholder="••••••••")
        btn_log = st.form_submit_button("🚀 Entrar al Sistema", type="primary", use_container_width=True)
        
        if btn_log:
            u_auth = autenticar_usuario(input_u, input_p)
            if u_auth:
                st.session_state['user'] = u_auth
                st.query_params.clear()
                st.rerun()
            else:
                st.error("❌ Usuario o contraseña incorrectos.")
    st.stop()

current_user = st.session_state.get('user') or {}
user_nombre = current_user.get('nombre', 'Líder')
user_rol = current_user.get('rol', 'lider')
user_grupo = str(current_user.get('codigo_grupo', '')).strip().split('.')[0] if current_user.get('codigo_grupo') else ""
user_sector = str(current_user.get('codigo_sector', '')).strip() if current_user.get('codigo_sector') else ""

# Guard de Cambio Obligatorio de Contraseña con Botón de Cancelar/Salir
if current_user.get('debe_cambiar_password', False):
    st.warning("🔒 **Cambio Obligatorio de Contraseña Inicial**")
    st.info("Por tu seguridad, debes definir una contraseña personal para continuar.")
    with st.form("form_pwd_mobile"):
        pn = st.text_input("🔑 Nueva Contraseña Personal", type="password")
        pc = st.text_input("🔑 Confirmar Nueva Contraseña", type="password")
        if st.form_submit_button("💾 Guardar y Continuar", use_container_width=True):
            if len(pn.strip()) < 4 or pn != pc:
                st.error("❌ Las contraseñas deben coincidir y tener al menos 4 caracteres.")
            else:
                ok, msg = cambiar_password_usuario(current_user['username'], pn)
                if ok:
                    current_user['debe_cambiar_password'] = False
                    st.session_state['user'] = current_user
                    st.query_params.clear()
                    st.success("✅ Contraseña actualizada.")
                    st.rerun()
                else:
                    st.error(f"❌ {msg}")
    if st.button("🚪 Cancelar y Salir", key="btn_cancel_mob", use_container_width=True):
        st.session_state['user'] = None
        st.query_params.clear()
        st.rerun()
    st.stop()

# 4. Header Compacto Móvil
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown(f"**📱 App Matices** • `Grupo {user_grupo if user_grupo else 'General'}`")
    st.caption(f"👤 {user_nombre}")
with col_h2:
    if st.button("🚪 Salir", key="btn_logout_mob", use_container_width=True):
        st.session_state['user'] = None
        st.query_params.clear()
        st.rerun()

# Si es Gerente o Admin, permitir seleccionar qué grupo auditar
grupo_activo = user_grupo
if user_rol in ['gerente', 'superadmin']:
    usuarios_todos = cargar_usuarios()
    grupos_disponibles = sorted(list(set(
        str(u.get('codigo_grupo')).strip().split('.')[0]
        for u in usuarios_todos.values()
        if u.get('codigo_grupo')
    )))
    if grupos_disponibles:
        col_g_sel1, col_g_sel2 = st.columns([2, 1])
        with col_g_sel1:
            grupo_activo = st.selectbox("👩‍💼 Ver Grupo:", options=grupos_disponibles, index=0, key="sel_grp_gerente")
        with col_g_sel2:
            st.caption("*(Vista Gerencial)*")

# 5. Carga de Datos Relacionales (Tableau, Geral, Objetivos Arte y Cómo Vamos)
df_tab = consultar_tableau_sql(grupo=grupo_activo if grupo_activo else None)
df_geral = consultar_geral_sql(grupo=grupo_activo if grupo_activo else None)

# Carga de Objetivos Arte
mapa_arte = cargar_objetivos_arte()
arte_lider = mapa_arte.get('por_grupo', {}).get(str(grupo_activo).strip(), {}) if grupo_activo else {}

# Carga de Cómo Vamos
df_cv_all = None
if os.path.exists('Base para el como vamos.xlsx'):
    try:
        df_cv_all = calcular_metas_ciclo('Base para el como vamos.xlsx')
    except Exception:
        pass

df_cv = pd.DataFrame()
if df_cv_all is not None and not df_cv_all.empty and grupo_activo:
    col_g = next((c for c in df_cv_all.columns if 'grupo' in str(c).lower()), None)
    if col_g:
        df_cv = df_cv_all[df_cv_all[col_g].astype(str).str.split('.').str[0].str.strip() == str(grupo_activo).strip()]

# 6. Pestañas Principales Móviles (3 Pestañas Condensadas en MAYÚSCULAS)
tab_cv, tab_tab, tab_lideres = st.tabs([
    "🎯 MIS DESAFÍOS",
    "📋 MI LISTADO",
    "👑 MIS LÍDERES"
])

# Funciones de formato y styler para tablas dinámicas
def formato_cop(val):
    try:
        if pd.isna(val) or val == "" or val is None:
            return "$0"
        num = float(limpiar_numero(val, 0.0))
        return f"${num:,.0f}".replace(",", ".")
    except Exception:
        return "$0"

def formato_porcentaje(val):
    try:
        if pd.isna(val) or val == "" or val is None:
            return "0.0%"
        num = float(str(val).replace('%', '').strip())
        return f"{num:.1f}%"
    except Exception:
        return "0.0%"

def formato_saldo_entero(val):
    try:
        if pd.isna(val) or val == "" or val is None:
            return "+0"
        num = int(float(limpiar_numero(val, 0)))
        return f"{num:+d}"
    except Exception:
        return "+0"

def aplicar_mapa_styler(styler, func, subset=None):
    if hasattr(styler, 'map'):
        return styler.map(func, subset=subset)
    elif hasattr(styler, 'applymap'):
        return styler.applymap(func, subset=subset)
    return styler

# ==============================================================================
# TAB 1: MIS DESAFÍOS & CÓMO VAMOS (OBJETIVOS ARTE + CÓMO VAMOS)
# ==============================================================================
with tab_cv:
    st.markdown("##### 🎯 Cuadro de Mando de Desafíos Líder")
    
    if df_cv.empty:
        st.info(f"ℹ️ **Metas del ciclo para el Grupo {grupo_activo if grupo_activo else ''}:**\n\nEl archivo de metas se sincronizará automáticamente. Puedes gestionar tu red en **'📋 MI LISTADO'** y consultar el comparativo en **'👑 MIS LÍDERES'**.")
    else:
        row_cv = df_cv.iloc[0]
        
        # 1. Disponibles
        disp_r = int(limpiar_numero(row_cv.get('Disponibles', 0)))
        disp_m = int(arte_lider.get('disponibles_proyectadas', 0))
        disp_pct = (disp_r / disp_m * 100.0) if disp_m > 0 else 0.0

        # 2. Activas
        act_r = int(limpiar_numero(row_cv.get('Real Activas', 0)))
        act_m = int(arte_lider.get('desafio_activas', row_cv.get('Objetivo Activas', 0)))
        act_pct = (act_r / act_m * 100.0) if act_m > 0 else 0.0
        
        # 3. Facturación
        fact_r = float(limpiar_numero(row_cv.get('Real Facturación', 0.0)))
        fact_m = float(arte_lider.get('desafio_facturacion', row_cv.get('Objetivo Facturación', 0.0)))
        fact_pct = (fact_r / fact_m * 100.0) if fact_m > 0 else 0.0
        
        # 4. Ganancia Estimada
        ganancia_cop = float(limpiar_numero(row_cv.get('Ganancia estimada', 0.0)))

        # 5. Saldo Comercial
        saldo_r = int(limpiar_numero(row_cv.get('Saldo', 0)))
        saldo_m = int(arte_lider.get('saldo_meta', 2))
        brecha_s = saldo_m - saldo_r

        # 6. Inicios + Reinicios
        ini_r = int(limpiar_numero(row_cv.get('Inicios', 0)))
        rein_r = int(limpiar_numero(row_cv.get('Reinicios', 0)))
        tot_ini_rei = ini_r + rein_r
        ini_rei_m = int(arte_lider.get('meta_inicios_reinicios', 0))
        ini_rei_pct = (tot_ini_rei / ini_rei_m * 100.0) if ini_rei_m > 0 else 0.0

        # Inactivas y Recuperos
        i1_val = int(limpiar_numero(row_cv.get('Inactiva 1', 0)))
        i2_val = int(limpiar_numero(row_cv.get('Inactiva 2', 0)))
        i3_val = int(limpiar_numero(row_cv.get('Inactiva 3', 0)))
        recup_r = int(limpiar_numero(row_cv.get('Recuperos', 0)))
        recup_m = int(arte_lider.get('meta_recuperos', 0))
        recup_pct = (recup_r / recup_m * 100.0) if recup_m > 0 else 0.0

        # Formatos de moneda y porcentaje limpios
        ganancia_fmt = f"${ganancia_cop:,.0f}".replace(",", ".")
        
        # FILA 1: 6 Tarjetas Principales del Negocio (Grid 2x3 o 3x2)
        st.markdown(f"""<div class="kpi-grid">
<div class="kpi-card">
<div class="kpi-title">🎯 DISPONIBLES PROY.</div>
<div class="kpi-val">{disp_r}</div>
<div class="kpi-sub {'kpi-sub-green' if disp_pct>=100 else 'kpi-sub-orange'}">{disp_pct:.1f}% Desafío ({disp_m})</div>
</div>
<div class="kpi-card">
<div class="kpi-title">👥 ACTIVAS</div>
<div class="kpi-val">{act_r}</div>
<div class="kpi-sub {'kpi-sub-green' if act_pct>=100 else 'kpi-sub-orange'}">{act_pct:.1f}% Desafío ({act_m})</div>
</div>
<div class="kpi-card">
<div class="kpi-title">💰 FACTURACIÓN TOTAL</div>
<div class="kpi-val">${fact_r/1e6:.1f}M</div>
<div class="kpi-sub {'kpi-sub-green' if fact_pct>=100 else 'kpi-sub-orange'}">{fact_pct:.1f}% Desafío (${fact_m/1e6:.1f}M)</div>
</div>
<div class="kpi-card">
<div class="kpi-title">💵 GANANCIA ESTIMADA LN</div>
<div class="kpi-val-gradient">{ganancia_fmt}</div>
<div class="kpi-sub kpi-sub-blue">Comisión + Potencializador</div>
</div>
<div class="kpi-card">
<div class="kpi-title">⚖️ SALDO COMERCIAL</div>
<div class="kpi-val" style="color:{'#15803d' if saldo_r>=saldo_m else '#b91c1c'};">{saldo_r:+d}</div>
<div class="kpi-sub {'kpi-sub-green' if saldo_r>=saldo_m else 'kpi-sub-red'}">{'Meta lograda (+' + str(saldo_m) + ')' if saldo_r>=saldo_m else 'Meta: +' + str(saldo_m) + ' (Falta ' + f'{brecha_s:+d}' + ')'}</div>
</div>
<div class="kpi-card">
<div class="kpi-title">🚀 INICIOS + REINICIOS</div>
<div class="kpi-val">{tot_ini_rei}</div>
<div class="kpi-sub {'kpi-sub-green' if ini_rei_pct>=100 else 'kpi-sub-orange'}">{ini_rei_pct:.1f}% Desafío ({ini_rei_m})</div>
</div>
</div>""", unsafe_allow_html=True)

        # FILA 2: Bolsa de Recuperación de Red (4 Tarjetas Abiertas)
        st.markdown("<p style='font-size:11px; font-weight:800; color:#E3007B; margin:8px 0 4px 2px;'>🌸 BOLSA DE RECUPERACIÓN (INACTIVAS & RECUPEROS):</p>", unsafe_allow_html=True)
        st.markdown(f"""<div class="kpi-grid">
<div class="kpi-card">
<div class="kpi-title">🌸 INACTIVA 1 (I1)</div>
<div class="kpi-val">{i1_val}</div>
<div class="kpi-sub kpi-sub-green">1 ciclo sin pedido</div>
</div>
<div class="kpi-card">
<div class="kpi-title">🌸 INACTIVA 2 (I2)</div>
<div class="kpi-val">{i2_val}</div>
<div class="kpi-sub kpi-sub-orange">2 ciclos sin pedido</div>
</div>
<div class="kpi-card">
<div class="kpi-title">⚠️ INACTIVA 3 (I3)</div>
<div class="kpi-val" style="color:#b91c1c;">{i3_val}</div>
<div class="kpi-sub kpi-sub-red">¡Riesgo Fuga a I4!</div>
</div>
<div class="kpi-card">
<div class="kpi-title">🎯 RECUPEROS LOGRADOS</div>
<div class="kpi-val">{recup_r} / {recup_m}</div>
<div class="kpi-sub {'kpi-sub-green' if recup_pct>=100 else 'kpi-sub-orange'}">{recup_pct:.1f}% Meta Arte</div>
</div>
</div>""", unsafe_allow_html=True)

        # Barra de progreso visual de Facturación
        st.caption(f"🎯 **Avance de Facturación:** ({fact_pct:.1f}% del desafío oficial)")
        st.progress(min(1.0, fact_pct / 100.0))

# ==============================================================================
# TAB 2: MI LISTADO (TABLA MAESTRA EXACTA COMO EN PC)
# ==============================================================================
with tab_tab:
    st.markdown("##### 📋 Mi Listado - Base Maestra Gestionable")
    st.caption("Escribe las notas de gestión por cada asesora. Se guardarán de forma permanente por `Codigo CB`. Puedes usar el corrector del explorador (subrayado rojo y clic derecho) para sugerencias ortográficas directas.")

    # Filtros Rápidos
    f_c1, f_c2, f_c3 = st.columns(3)
    with f_c1:
        opciones_sit = ["Todas"]
        if not df_tab.empty and 'Sit. Comercial' in df_tab.columns:
            sits_unicas = [str(x) for x in df_tab['Sit. Comercial'].dropna().unique() if str(x).strip()]
            opciones_sit += sorted(sits_unicas)
        filtro_sit = st.selectbox("🏷️ Sit. Comercial", options=opciones_sit, key="mob_f_sit")
        
    with f_c2:
        filtro_mora = st.selectbox("💳 Deuda en Mora", options=["Todas", "🔴 Solo con Mora", "🟢 Al Día"], key="mob_f_mora")

    with f_c3:
        filtro_ped = st.selectbox("⌛ Pedidos Pendientes", options=["Todos", "Con Pedidos Pendientes (> 0)", "Sin Pedidos Pendientes (0)"], key="mob_f_ped")

    # Búsqueda rápida por nombre, documento o código
    busq_nom = st.text_input("🔍 Buscar asesora, documento o código CB...", placeholder="Escribe nombre, cédula o código...", key="mob_b_nom").strip()

    # Aplicar filtros
    df_tab_filtrado = df_tab.copy() if not df_tab.empty else pd.DataFrame()
    
    if not df_tab_filtrado.empty:
        if filtro_sit != "Todas" and 'Sit. Comercial' in df_tab_filtrado.columns:
            df_tab_filtrado = df_tab_filtrado[df_tab_filtrado['Sit. Comercial'].astype(str) == filtro_sit]
            
        if 'Deuda Mora' in df_tab_filtrado.columns:
            df_tab_filtrado['Deuda_Num'] = df_tab_filtrado['Deuda Mora'].apply(lambda x: limpiar_numero(x, 0.0))
            if filtro_mora == "🔴 Solo con Mora":
                df_tab_filtrado = df_tab_filtrado[df_tab_filtrado['Deuda_Num'] > 0]
            elif filtro_mora == "🟢 Al Día":
                df_tab_filtrado = df_tab_filtrado[df_tab_filtrado['Deuda_Num'] <= 0]

        if 'Ped. Pendientes' in df_tab_filtrado.columns:
            if filtro_ped == "Con Pedidos Pendientes (> 0)":
                df_tab_filtrado = df_tab_filtrado[df_tab_filtrado['Ped. Pendientes'] > 0]
            elif filtro_ped == "Sin Pedidos Pendientes (0)":
                df_tab_filtrado = df_tab_filtrado[df_tab_filtrado['Ped. Pendientes'] <= 0]
                
        if busq_nom:
            mask_busq = pd.Series(False, index=df_tab_filtrado.index)
            if 'Asesora / Consultora' in df_tab_filtrado.columns:
                mask_busq = mask_busq | df_tab_filtrado['Asesora / Consultora'].astype(str).str.contains(busq_nom, case=False, na=False)
            if 'DocumentoGPP' in df_tab_filtrado.columns:
                mask_busq = mask_busq | df_tab_filtrado['DocumentoGPP'].astype(str).str.contains(busq_nom, case=False, na=False)
            if 'Código CB' in df_tab_filtrado.columns:
                mask_busq = mask_busq | df_tab_filtrado['Código CB'].astype(str).str.contains(busq_nom, case=False, na=False)
            df_tab_filtrado = df_tab_filtrado[mask_busq]

    if df_tab_filtrado.empty:
        st.info("ℹ️ No hay consultoras con los filtros seleccionados.")
    else:
        # Limpiar, ordenar y estandarizar columnas para que coincidan exactamente con la base canónica (16 columnas)
        df_edit_view = limpiar_y_ordenar_columnas_tableau(df_tab_filtrado, {}, es_lider=True)

        # Limpiar cualquier flotante residual en todo el DataFrame para eliminar decimales (.000000)
        for c in df_edit_view.columns:
            if c not in ['DocumentoGPP', 'Celular', 'Código CB', 'Codigo CB'] and pd.api.types.is_float_dtype(df_edit_view[c]):
                df_edit_view[c] = df_edit_view[c].fillna(0).round().astype('int64')

        # Usar st.data_editor para permitir editar notas directamente en la tabla
        col_config = {}
        for col_name in df_edit_view.columns:
            # Si es una columna de dinero (Deuda o Facturación), formatear con $
            if 'Deuda' in col_name or 'Fact.' in col_name:
                col_config[col_name] = st.column_config.NumberColumn(col_name, format="$%d", disabled=True)
            # Si es DocumentoGPP, Celular o Código CB, formatear como texto limpio sin comas
            elif col_name in ['DocumentoGPP', 'Celular', 'Código CB', 'Codigo CB']:
                col_config[col_name] = st.column_config.TextColumn(str(col_name), disabled=True)
            # Si es una columna numérica (Pts, Crédito, Pedidos, Ciclos), formatear como número entero limpio sin $
            elif 'Pts' in col_name or 'Ped.' in col_name or 'Ciclos' in col_name or 'Credito' in col_name or 'Crédito' in col_name:
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

        edited_df = st.data_editor(
            df_edit_styled,
            column_config=col_config,
            use_container_width=True,
            hide_index=True,
            key="mob_editor_tabla_tableau"
        )

        # Auto-guardado inteligente en segundo plano al modificar cualquier celda
        editor_state = st.session_state.get("mob_editor_tabla_tableau", {})
        edited_rows = editor_state.get("edited_rows", {})
        
        if edited_rows:
            dict_autoguardar = {}
            for row_idx_str, row_changes in edited_rows.items():
                if "Notas / Comentarios Líder" in row_changes:
                    try:
                        row_idx = int(row_idx_str)
                        if row_idx < len(df_edit_view):
                            codigo_key = str(df_edit_view.iloc[row_idx].get('Código CB', '')).strip()
                            nueva_nota = str(row_changes["Notas / Comentarios Líder"]).strip()
                            if codigo_key:
                                dict_autoguardar[codigo_key] = nueva_nota
                    except Exception:
                        pass
            
            if dict_autoguardar:
                guardar_todos_comentarios(dict_autoguardar)
                st.toast(f"💾 Guardado: {len(dict_autoguardar)} nota(s) actualizada(s)", icon="✅")

        # Botón de guardado manual
        col_s1, col_s2 = st.columns([1.5, 2.5])
        with col_s1:
            if st.button("💾 Guardar Manualmente", type="primary", use_container_width=True, key="mob_save_manual_btn"):
                dict_guardar = {}
                for idx, row in edited_df.iterrows():
                    codigo_key = str(row.get('Código CB', '')).strip()
                    nota_val = str(row.get('Notas / Comentarios Líder', '')).strip()
                    if codigo_key:
                        dict_guardar[codigo_key] = nota_val
                
                if guardar_todos_comentarios(dict_guardar):
                    st.success("✅ ¡Todas las notas han sido guardadas exitosamente!")
                    st.rerun()
        with col_s2:
            st.caption("🟢 **Guardado automático activo**: Al escribir una nota y pulsar `Enter`, se guarda de forma instantánea.")

# ==============================================================================
# TAB 3: MIS LÍDERES (TODAS LAS TABLAS Y SEGUIMIENTO COMPARATIVO)
# ==============================================================================
with tab_lideres:
    st.markdown("##### 👑 Mis Líderes - Diagnóstico y Comparativo")
    
    if df_cv_all is None or df_cv_all.empty:
        st.info("ℹ️ No hay datos de 'Cómo Vamos' disponibles para mostrar las tablas dinámicas de líderes.")
    else:
        df_diag = df_cv_all.copy()
        col_lider = 'Nombre de consultora' if 'Nombre de consultora' in df_diag.columns else (df_diag.columns[0] if len(df_diag.columns) > 0 else '')

        # Clasificación: Semillas (Desafío <= 0) vs Líderes (Desafío > 0)
        col_obj_fact_chk = 'Objetivo Facturación' if 'Objetivo Facturación' in df_diag.columns else None
        if col_obj_fact_chk:
            df_diag['Tipo_Red'] = df_diag[col_obj_fact_chk].apply(
                lambda v: '👑 Líder' if limpiar_numero(v, 0.0) > 0 else '🌱 Semilla'
            )
        else:
            df_diag['Tipo_Red'] = '👑 Líder'

        count_tot = len(df_diag)
        count_lideres = int((df_diag['Tipo_Red'] == '👑 Líder').sum())
        count_semillas = int((df_diag['Tipo_Red'] == '🌱 Semilla').sum())

        filtro_segmento = st.radio(
            "🎯 **Filtrar por Tipo de Red:**",
            options=[
                f"🌟 Todas ({count_tot})",
                f"👑 Líderes ({count_lideres})",
                f"🌱 Semillas ({count_semillas})"
            ],
            horizontal=True,
            key="mob_filtro_segmento_red"
        )

        if "👑 Líderes" in filtro_segmento:
            df_diag = df_diag[df_diag['Tipo_Red'] == '👑 Líder'].copy()
        elif "🌱 Semillas" in filtro_segmento:
            df_diag = df_diag[df_diag['Tipo_Red'] == '🌱 Semilla'].copy()

        # --- 1. TABLA DE FACTURACIÓN Y CUMPLIMIENTO ---
        st.markdown("---")
        st.markdown("###### 💰 1. Tabla de Facturación y Cumplimiento")
        
        cols_fact_exactas = [
            col_lider, 'Tipo_Red', 'Objetivo Facturación', 'Real Facturación', 'Cumplimiento Facturación',
            'Avance % Facturación', 'Productividad', 'Falta para el 100%', 'Falta para el 110%', 'Ganancia estimada'
        ]
        cols_presentes = [c for c in cols_fact_exactas if c in df_diag.columns]
        
        if 'Cumplimiento Facturación' in df_diag.columns:
            df_fact_sorted = df_diag.sort_values(by='Cumplimiento Facturación', ascending=False)
        else:
            df_fact_sorted = df_diag
            
        df_fact_view = df_fact_sorted[cols_presentes].copy()
        
        nombres_clery = {
            col_lider: 'LÍDER DE NEGOCIOS',
            'Tipo_Red': 'TIPO',
            'Objetivo Facturación': 'DESAFÍO FACTURACIÓN',
            'Real Facturación': 'FACTURACIÓN A HOY',
            'Cumplimiento Facturación': 'CUMPLIMIENTO DE FACTURACIÓN',
            'Avance % Facturación': 'AVANCE %',
            'Productividad': 'PRODUCTIVIDAD',
            'Falta para el 100%': 'FALTA PARA EL 100%',
            'Falta para el 110%': 'CUÁNTO FALTA PARA EL 110%',
            'Ganancia estimada': 'GANANCIA ESTIMADA TOTAL'
        }
        df_fact_view = df_fact_view.rename(columns=nombres_clery)
        
        df_fact_formatted = df_fact_view.copy()
        if 'DESAFÍO FACTURACIÓN' in df_fact_formatted.columns:
            df_fact_formatted['DESAFÍO FACTURACIÓN'] = df_fact_formatted['DESAFÍO FACTURACIÓN'].apply(formato_cop)
        if 'FACTURACIÓN A HOY' in df_fact_formatted.columns:
            df_fact_formatted['FACTURACIÓN A HOY'] = df_fact_formatted['FACTURACIÓN A HOY'].apply(formato_cop)
        if 'CUMPLIMIENTO DE FACTURACIÓN' in df_fact_formatted.columns:
            df_fact_formatted['CUMPLIMIENTO DE FACTURACIÓN'] = df_fact_formatted['CUMPLIMIENTO DE FACTURACIÓN'].apply(formato_porcentaje)
        if 'AVANCE %' in df_fact_formatted.columns:
            df_fact_formatted['AVANCE %'] = df_fact_formatted['AVANCE %'].apply(formato_porcentaje)
        if 'PRODUCTIVIDAD' in df_fact_formatted.columns:
            df_fact_formatted['PRODUCTIVIDAD'] = df_fact_formatted['PRODUCTIVIDAD'].apply(formato_cop)
        if 'FALTA PARA EL 100%' in df_fact_formatted.columns:
            df_fact_formatted['FALTA PARA EL 100%'] = df_fact_formatted['FALTA PARA EL 100%'].apply(formato_cop)
        if 'CUÁNTO FALTA PARA EL 110%' in df_fact_formatted.columns:
            df_fact_formatted['CUÁNTO FALTA PARA EL 110%'] = df_fact_formatted['CUÁNTO FALTA PARA EL 110%'].apply(formato_cop)
        if 'GANANCIA ESTIMADA TOTAL' in df_fact_formatted.columns:
            df_fact_formatted['GANANCIA ESTIMADA TOTAL'] = df_fact_formatted['GANANCIA ESTIMADA TOTAL'].apply(formato_cop)
            
        styler_fact = df_fact_formatted.style

        def _estilo_tipo(val_str):
            if 'Líder' in str(val_str):
                return 'background-color: #dbeafe; color: #1e40af; font-weight: bold;'
            elif 'Semilla' in str(val_str):
                return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
            return ''

        def _estilo_cump_fact(val_str):
            try:
                num = float(str(val_str).replace('%', '').strip())
                if num >= 100.0:
                    return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                elif num >= 90.0:
                    return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                else:
                    return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
            except Exception:
                return ''

        def _estilo_avance_pct_fact(val_str):
            try:
                num = float(str(val_str).replace('%', '').strip())
                if num >= 90.0:
                    return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                elif num >= 80.0:
                    return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                else:
                    return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
            except Exception:
                return ''

        def _estilo_falta_dinero(val_str):
            try:
                s = str(val_str)
                if '-' in s or '$0' in s:
                    return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                else:
                    return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
            except Exception:
                return ''

        def _estilo_ganancia_total(val_str):
            try:
                s = str(val_str)
                if s and '$0' not in s and '$' in s:
                    return 'background-color: #e0f2fe; color: #0369a1; font-weight: bold;'
                return ''
            except Exception:
                return ''

        if 'TIPO' in df_fact_formatted.columns:
            styler_fact = aplicar_mapa_styler(styler_fact, _estilo_tipo, subset=['TIPO'])
        if 'CUMPLIMIENTO DE FACTURACIÓN' in df_fact_formatted.columns:
            styler_fact = aplicar_mapa_styler(styler_fact, _estilo_cump_fact, subset=['CUMPLIMIENTO DE FACTURACIÓN'])
        if 'AVANCE %' in df_fact_formatted.columns:
            styler_fact = aplicar_mapa_styler(styler_fact, _estilo_avance_pct_fact, subset=['AVANCE %'])
        if 'FALTA PARA EL 100%' in df_fact_formatted.columns:
            styler_fact = aplicar_mapa_styler(styler_fact, _estilo_falta_dinero, subset=['FALTA PARA EL 100%'])
        if 'CUÁNTO FALTA PARA EL 110%' in df_fact_formatted.columns:
            styler_fact = aplicar_mapa_styler(styler_fact, _estilo_falta_dinero, subset=['CUÁNTO FALTA PARA EL 110%'])
        if 'GANANCIA ESTIMADA TOTAL' in df_fact_formatted.columns:
            styler_fact = aplicar_mapa_styler(styler_fact, _estilo_ganancia_total, subset=['GANANCIA ESTIMADA TOTAL'])

        st.dataframe(styler_fact, use_container_width=True)

        # --- 2. TABLA DE ACTIVAS / PEDIDOS ---
        st.markdown("---")
        st.markdown("###### 👥 2. Tabla de Activas / Pedidos")
        
        if 'Objetivo Activas' in df_diag.columns and 'Real Activas' in df_diag.columns:
            obj_a_num = df_diag['Objetivo Activas'].apply(lambda v: limpiar_numero(v, 0.0))
            real_a_num = df_diag['Real Activas'].apply(lambda v: limpiar_numero(v, 0.0))
            df_diag['Cumplimiento Activas'] = (real_a_num / obj_a_num.replace(0, pd.NA) * 100.0).fillna(0.0)

        cols_act_exactas = [
            col_lider, 'Tipo_Red', 'Objetivo Activas', 'Real Activas', 'Cumplimiento Activas',
            'Saldo', 'Disponibles', 'Inicios', 'Reinicios', 'Recuperos'
        ]
        cols_act_presentes = [c for c in cols_act_exactas if c in df_diag.columns]
        
        if 'Cumplimiento Activas' in df_diag.columns:
            df_act_sorted = df_diag.sort_values(by='Cumplimiento Activas', ascending=False)
        elif 'Real Activas' in df_diag.columns:
            df_act_sorted = df_diag.sort_values(by='Real Activas', ascending=False)
        else:
            df_act_sorted = df_diag
            
        df_act_view = df_act_sorted[cols_act_presentes].copy()
        nombres_clery_act = {
            col_lider: 'LÍDER DE NEGOCIOS',
            'Tipo_Red': 'TIPO',
            'Objetivo Activas': 'ACTIVAS METAS',
            'Real Activas': 'ACTIVAS HOY (PEDIDOS)',
            'Cumplimiento Activas': 'CUMPLIMIENTO ACTIVAS',
            'Saldo': 'SALDO ACTIVAS',
            'Disponibles': 'DISPONIBLES',
            'Inicios': 'INICIOS HOY',
            'Reinicios': 'REINICIOS HOY',
            'Recuperos': 'RECUPEROS HOY'
        }
        df_act_view = df_act_view.rename(columns=nombres_clery_act)
        
        df_act_formatted = df_act_view.copy()
        if 'ACTIVAS METAS' in df_act_formatted.columns:
            df_act_formatted['ACTIVAS METAS'] = df_act_formatted['ACTIVAS METAS'].apply(lambda v: f"{int(limpiar_numero(v))}")
        if 'ACTIVAS HOY (PEDIDOS)' in df_act_formatted.columns:
            df_act_formatted['ACTIVAS HOY (PEDIDOS)'] = df_act_formatted['ACTIVAS HOY (PEDIDOS)'].apply(lambda v: f"{int(limpiar_numero(v))}")
        if 'CUMPLIMIENTO ACTIVAS' in df_act_formatted.columns:
            df_act_formatted['CUMPLIMIENTO ACTIVAS'] = df_act_formatted['CUMPLIMIENTO ACTIVAS'].apply(formato_porcentaje)
        if 'SALDO ACTIVAS' in df_act_formatted.columns:
            df_act_formatted['SALDO ACTIVAS'] = df_act_formatted['SALDO ACTIVAS'].apply(formato_saldo_entero)
        if 'DISPONIBLES' in df_act_formatted.columns:
            df_act_formatted['DISPONIBLES'] = df_act_formatted['DISPONIBLES'].apply(lambda v: f"{int(limpiar_numero(v))}")
        if 'INICIOS HOY' in df_act_formatted.columns:
            df_act_formatted['INICIOS HOY'] = df_act_formatted['INICIOS HOY'].apply(lambda v: f"{int(limpiar_numero(v))}")
        if 'REINICIOS HOY' in df_act_formatted.columns:
            df_act_formatted['REINICIOS HOY'] = df_act_formatted['REINICIOS HOY'].apply(lambda v: f"{int(limpiar_numero(v))}")
        if 'RECUPEROS HOY' in df_act_formatted.columns:
            df_act_formatted['RECUPEROS HOY'] = df_act_formatted['RECUPEROS HOY'].apply(lambda v: f"{int(limpiar_numero(v))}")

        styler_act = df_act_formatted.style

        def _estilo_cump_act(val_str):
            try:
                num = float(str(val_str).replace('%', '').strip())
                if num >= 100.0:
                    return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                elif num >= 90.0:
                    return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                else:
                    return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
            except Exception:
                return ''

        def _estilo_saldo_act(val_str):
            try:
                num = float(limpiar_numero(val_str, 0))
                if num < 0:
                    return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
                else:
                    return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
            except Exception:
                return ''

        def _estilo_ingresos_act(val_str):
            try:
                num = int(limpiar_numero(val_str, 0))
                if num > 0:
                    return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                return ''
            except Exception:
                return ''

        if 'TIPO' in df_act_formatted.columns:
            styler_act = aplicar_mapa_styler(styler_act, _estilo_tipo, subset=['TIPO'])
        if 'CUMPLIMIENTO ACTIVAS' in df_act_formatted.columns:
            styler_act = aplicar_mapa_styler(styler_act, _estilo_cump_act, subset=['CUMPLIMIENTO ACTIVAS'])
        if 'SALDO ACTIVAS' in df_act_formatted.columns:
            styler_act = aplicar_mapa_styler(styler_act, _estilo_saldo_act, subset=['SALDO ACTIVAS'])
        for col_ing_sub in ['INICIOS HOY', 'REINICIOS HOY', 'RECUPEROS HOY']:
            if col_ing_sub in df_act_formatted.columns:
                styler_act = aplicar_mapa_styler(styler_act, _estilo_ingresos_act, subset=[col_ing_sub])

        st.dataframe(styler_act, use_container_width=True)

        # --- 3. CUADRO RESUMEN DE DISPONIBLES ---
        st.markdown("---")
        st.markdown("###### 📋 3. Cuadro Resumen de Disponibles (Desafío vs. Avance por Día)")
        dia_corte = st.number_input("📅 Día de Avance (Editable):", min_value=1, max_value=21, value=14, step=1, key="mob_dia_corte_14_key")
        nombre_col_dia = f"Dia {dia_corte}"

        if col_lider and col_lider in df_diag.columns:
            df_disp_prep = df_diag.copy()
            col_grp_diag = next((c for c in df_disp_prep.columns if any(k in str(c).lower() for k in ['código de grupo', 'codigo de grupo', 'cód. grupo', 'cod grupo', 'grupo'])), None)
            
            mapa_grp_disp = mapa_arte.get('por_grupo', {})
            mapa_nom_disp = mapa_arte.get('por_nombre', {})

            col_disp_actual = 'Disponibles' if 'Disponibles' in df_disp_prep.columns else ('Real Activas' if 'Real Activas' in df_disp_prep.columns else None)

            if col_disp_actual:
                df_disp_calc = pd.DataFrame()
                df_disp_calc['LÍDER DE NEGOCIOS'] = df_disp_prep[col_lider].astype(str)

                def _obtener_desafio_disp_row(row):
                    g = str(row.get(col_grp_diag, '')).strip().split('.')[0] if col_grp_diag else ''
                    nom = str(row.get(col_lider, '')).strip().lower()
                    target = mapa_grp_disp.get(g) or mapa_nom_disp.get(nom)
                    if target:
                        val = target.get('disponibles_proyectadas', 0) or target.get('disponibles_esperadas', 0)
                        if val > 0:
                            return int(val)
                    if 'Meta Disponibles Esperadas' in row and int(limpiar_numero(row['Meta Disponibles Esperadas'], 0)) > 0:
                        return int(limpiar_numero(row['Meta Disponibles Esperadas'], 0))
                    return int(limpiar_numero(row.get(col_disp_actual, 0), 0))

                df_disp_calc['Disponibles Proyectadas'] = df_disp_prep.apply(_obtener_desafio_disp_row, axis=1)
                df_disp_calc[nombre_col_dia] = df_disp_prep[col_disp_actual].apply(lambda v: int(limpiar_numero(v, 0)))
                
                df_disp_calc['% Cump LN'] = df_disp_calc.apply(
                    lambda r: (r[nombre_col_dia] / r['Disponibles Proyectadas'] * 100.0) if r['Disponibles Proyectadas'] > 0 else 0.0,
                    axis=1
                )
                df_disp_calc['falta'] = df_disp_calc.apply(
                    lambda r: max(0, r['Disponibles Proyectadas'] - r[nombre_col_dia]),
                    axis=1
                )

                df_disp_calc = df_disp_calc.sort_values(by='% Cump LN', ascending=False).reset_index(drop=True)

                tot_desafios = int(df_disp_calc['Disponibles Proyectadas'].sum())
                tot_dia = int(df_disp_calc[nombre_col_dia].sum())
                tot_cump = (tot_dia / tot_desafios * 100.0) if tot_desafios > 0 else 0.0
                tot_falta = max(0, tot_desafios - tot_dia)

                row_total = pd.DataFrame([{
                    'LÍDER DE NEGOCIOS': 'TOTAL GENERAL',
                    'Disponibles Proyectadas': tot_desafios,
                    nombre_col_dia: tot_dia,
                    '% Cump LN': tot_cump,
                    'falta': tot_falta
                }])
                df_disp_final = pd.concat([df_disp_calc, row_total], ignore_index=True)

                df_disp_formatted = df_disp_final.copy()
                df_disp_formatted['Disponibles Proyectadas'] = df_disp_formatted['Disponibles Proyectadas'].apply(lambda v: f"{int(v):,}".replace(",", "."))
                df_disp_formatted[nombre_col_dia] = df_disp_formatted[nombre_col_dia].apply(lambda v: f"{int(v):,}".replace(",", "."))
                df_disp_formatted['% Cump LN'] = df_disp_formatted['% Cump LN'].apply(lambda v: f"{v:.1f}%")
                df_disp_formatted['falta'] = df_disp_formatted['falta'].apply(lambda v: f"{int(v):,}".replace(",", "."))

                def _estilo_cump_val(val_str):
                    try:
                        num = float(str(val_str).replace('%', '').strip())
                        if num >= 95.0:
                            return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                        elif num >= 90.0:
                            return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                        else:
                            return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
                    except Exception:
                        return ''

                def _estilo_falta_val(val_str):
                    try:
                        num = int(str(val_str).replace('.', '').strip())
                        if num == 0:
                            return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                        elif num <= 10:
                            return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                        else:
                            return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
                    except Exception:
                        return ''

                styler_disp = df_disp_formatted.style
                if hasattr(styler_disp, 'map'):
                    styler_disp = styler_disp.map(_estilo_cump_val, subset=['% Cump LN']).map(_estilo_falta_val, subset=['falta'])
                elif hasattr(styler_disp, 'applymap'):
                    styler_disp = styler_disp.applymap(_estilo_cump_val, subset=['% Cump LN']).applymap(_estilo_falta_val, subset=['falta'])

                st.dataframe(styler_disp, use_container_width=True)

        # --- 4. CUADRO RESUMEN DE INICIOS + REINICIOS ---
        st.markdown("---")
        st.markdown("###### 🚀 4. Cuadro Resumen de Inicios + Reinicios")

        if col_lider and col_lider in df_diag.columns:
            df_ing_prep = df_diag.copy()
            col_inicios = 'Inicios' if 'Inicios' in df_ing_prep.columns else None
            col_reinicios = 'Reinicios' if 'Reinicios' in df_ing_prep.columns else None
            col_meta_ing = next((c for c in df_ing_prep.columns if any(k in str(c).lower() for k in ['meta inicios + reinicios', 'meta_inicios_reinicios', 'meta inicios', 'meta_inicios', 'inicios + reinicios'])), None)

            df_ing_calc = pd.DataFrame()
            df_ing_calc['LÍDER DE NEGOCIOS'] = df_ing_prep[col_lider].astype(str)

            val_inicios = df_ing_prep[col_inicios].apply(lambda v: limpiar_numero(v, 0)) if col_inicios else 0
            val_reinicios = df_ing_prep[col_reinicios].apply(lambda v: limpiar_numero(v, 0)) if col_reinicios else 0
            df_ing_calc['Hoy'] = (val_inicios + val_reinicios).astype(int)

            if col_meta_ing:
                df_ing_calc['Meta'] = df_ing_prep[col_meta_ing].apply(lambda v: int(limpiar_numero(v, 0)))
                df_ing_calc['Meta'] = df_ing_calc.apply(lambda r: r['Meta'] if r['Meta'] > 0 else max(5, int(r['Hoy'] + 3)), axis=1)
            else:
                df_ing_calc['Meta'] = df_ing_calc['Hoy'].apply(lambda h: max(5, int(h + 3)))

            df_ing_calc['Avance'] = df_ing_calc.apply(
                lambda r: (r['Hoy'] / r['Meta'] * 100.0) if r['Meta'] > 0 else 0.0,
                axis=1
            )
            df_ing_calc['para activar!'] = df_ing_calc.apply(
                lambda r: max(0, r['Meta'] - r['Hoy']),
                axis=1
            )

            df_ing_calc = df_ing_calc.sort_values(by='Avance', ascending=False).reset_index(drop=True)

            tot_meta_ing = int(df_ing_calc['Meta'].sum())
            tot_hoy_ing = int(df_ing_calc['Hoy'].sum())
            tot_av_ing = (tot_hoy_ing / tot_meta_ing * 100.0) if tot_meta_ing > 0 else 0.0
            tot_act_ing = max(0, tot_meta_ing - tot_hoy_ing)

            row_total_ing = pd.DataFrame([{
                'LÍDER DE NEGOCIOS': 'TOTAL GENERAL',
                'Hoy': tot_hoy_ing,
                'Meta': tot_meta_ing,
                'Avance': tot_av_ing,
                'para activar!': tot_act_ing
            }])
            df_ing_final = pd.concat([df_ing_calc, row_total_ing], ignore_index=True)

            df_ing_formatted = df_ing_final[['LÍDER DE NEGOCIOS', 'Hoy', 'Meta', 'Avance', 'para activar!']].copy()
            df_ing_formatted['Hoy'] = df_ing_formatted['Hoy'].apply(lambda v: f"{int(v):,}".replace(",", "."))
            df_ing_formatted['Meta'] = df_ing_formatted['Meta'].apply(lambda v: f"{int(v):,}".replace(",", "."))
            df_ing_formatted['Avance'] = df_ing_formatted['Avance'].apply(lambda v: f"{v:.1f}%")
            df_ing_formatted['para activar!'] = df_ing_formatted['para activar!'].apply(lambda v: f"{int(v):,}".replace(",", "."))

            def _estilo_avance_magenta(val_str):
                return 'background-color: #e3007b; color: #ffffff; font-weight: bold;'

            def _estilo_para_activar(val_str):
                try:
                    num = int(str(val_str).replace('.', '').strip())
                    if num <= 3:
                        return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                    elif num <= 6:
                        return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                    else:
                        return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
                except Exception:
                    return ''

            styler_ing = df_ing_formatted.style
            if hasattr(styler_ing, 'map'):
                styler_ing = styler_ing.map(_estilo_avance_magenta, subset=['Avance']).map(_estilo_para_activar, subset=['para activar!'])
            elif hasattr(styler_ing, 'applymap'):
                styler_ing = styler_ing.applymap(_estilo_avance_magenta, subset=['Avance']).applymap(_estilo_para_activar, subset=['para activar!'])

            st.dataframe(styler_ing, use_container_width=True)

        # --- 5. CUADRO RESUMEN DE RECUPEROS ---
        st.markdown("---")
        st.markdown("###### 🎯 5. Cuadro Resumen de Recuperos")

        if col_lider and col_lider in df_diag.columns:
            df_rec_prep = df_diag.copy()
            col_recuperos = 'Recuperos' if 'Recuperos' in df_rec_prep.columns else None
            col_meta_rec = next((c for c in df_rec_prep.columns if any(k in str(c).lower() for k in ['meta recuperos', 'meta_recuperos', 'recuperos_meta'])), None)

            df_rec_calc = pd.DataFrame()
            df_rec_calc['LÍDER DE NEGOCIOS'] = df_rec_prep[col_lider].astype(str)

            val_rec = df_rec_prep[col_recuperos].apply(lambda v: limpiar_numero(v, 0)) if col_recuperos else 0
            df_rec_calc['Hoy'] = val_rec.astype(int)

            if col_meta_rec:
                df_rec_calc['Meta'] = df_rec_prep[col_meta_rec].apply(lambda v: int(limpiar_numero(v, 0)))
                df_rec_calc['Meta'] = df_rec_calc.apply(lambda r: r['Meta'] if r['Meta'] > 0 else max(4, int(r['Hoy'] + 2)), axis=1)
            else:
                df_rec_calc['Meta'] = df_rec_calc['Hoy'].apply(lambda h: max(4, int(h + 2)))

            df_rec_calc['Avance'] = df_rec_calc.apply(
                lambda r: (r['Hoy'] / r['Meta'] * 100.0) if r['Meta'] > 0 else 0.0,
                axis=1
            )
            df_rec_calc['para activar!'] = df_rec_calc.apply(
                lambda r: max(0, r['Meta'] - r['Hoy']),
                axis=1
            )

            df_rec_calc = df_rec_calc.sort_values(by='Avance', ascending=False).reset_index(drop=True)

            tot_meta_rec = int(df_rec_calc['Meta'].sum())
            tot_hoy_rec = int(df_rec_calc['Hoy'].sum())
            tot_av_rec = (tot_hoy_rec / tot_meta_rec * 100.0) if tot_meta_rec > 0 else 0.0
            tot_act_rec = max(0, tot_meta_rec - tot_hoy_rec)

            row_total_rec = pd.DataFrame([{
                'LÍDER DE NEGOCIOS': 'TOTAL GENERAL',
                'Hoy': tot_hoy_rec,
                'Meta': tot_meta_rec,
                'Avance': tot_av_rec,
                'para activar!': tot_act_rec
            }])
            df_rec_final = pd.concat([df_rec_calc, row_total_rec], ignore_index=True)

            df_rec_formatted = df_rec_final[['LÍDER DE NEGOCIOS', 'Hoy', 'Meta', 'Avance', 'para activar!']].copy()
            df_rec_formatted['Hoy'] = df_rec_formatted['Hoy'].apply(lambda v: f"{int(v):,}".replace(",", "."))
            df_rec_formatted['Meta'] = df_rec_formatted['Meta'].apply(lambda v: f"{int(v):,}".replace(",", "."))
            df_rec_formatted['Avance'] = df_rec_formatted['Avance'].apply(lambda v: f"{v:.1f}%")
            df_rec_formatted['para activar!'] = df_rec_formatted['para activar!'].apply(lambda v: f"{int(v):,}".replace(",", "."))

            styler_rec = df_rec_formatted.style
            if hasattr(styler_rec, 'map'):
                styler_rec = styler_rec.map(_estilo_avance_magenta, subset=['Avance']).map(_estilo_para_activar, subset=['para activar!'])
            elif hasattr(styler_rec, 'applymap'):
                styler_rec = styler_rec.applymap(_estilo_avance_magenta, subset=['Avance']).applymap(_estilo_para_activar, subset=['para activar!'])

            st.dataframe(styler_rec, use_container_width=True)

        # --- 6. CUADRO RESUMEN DE RETENCIÓN I2 ---
        st.markdown("---")
        st.markdown("###### 🔄 6. Cuadro Resumen de Retención I2 (Meta 8% Máx. Fuga I2)")

        if col_lider and col_lider in df_diag.columns:
            df_i2_prep = df_diag.copy()
            col_disp_i2 = 'Disponibles' if 'Disponibles' in df_i2_prep.columns else None
            col_i2 = next((c for c in df_i2_prep.columns if str(c).lower().strip() in ['inactiva 2', 'inactiva_2', 'inactivas 2', 'inactivas_2', 'i2']), None)
            col_i2_ant = next((c for c in df_i2_prep.columns if 'inactiva 2_anterior' in str(c).lower() or 'inactivas 2_anterior' in str(c).lower()), None)

            if col_disp_i2 and col_i2:
                df_i2_calc = pd.DataFrame()
                df_i2_calc['LÍDER DE NEGOCIOS'] = df_i2_prep[col_lider].astype(str)

                val_disp2 = df_i2_prep[col_disp_i2].apply(lambda v: limpiar_numero(v, 0.0))
                val_i2 = df_i2_prep[col_i2].apply(lambda v: limpiar_numero(v, 0.0))

                df_i2_calc['META RETENCIÓN I2'] = (val_disp2 * 0.08).round().astype(int)
                df_i2_calc['FALTA I2 ACTIVARSE'] = (val_i2 - df_i2_calc['META RETENCIÓN I2']).round().astype(int)
                df_i2_calc['% RETENCIÓN META 8%'] = (val_i2 / val_disp2.replace(0, pd.NA) * 100.0).fillna(0.0)

                if col_i2_ant:
                    val_i2_ant = df_i2_prep[col_i2_ant].apply(lambda v: limpiar_numero(v, 0.0))
                    df_i2_calc['AVANCE RETENCION I2'] = (val_i2_ant - val_i2).fillna(0).astype(int)
                else:
                    df_i2_calc['AVANCE RETENCION I2'] = 0

                df_i2_calc = df_i2_calc.sort_values(by='% RETENCIÓN META 8%', ascending=True).reset_index(drop=True)

                tot_disp_i2 = float(val_disp2.sum())
                tot_i2 = float(val_i2.sum())
                tot_meta_i2 = int(df_i2_calc['META RETENCIÓN I2'].sum())
                tot_falta_i2 = int(df_i2_calc['FALTA I2 ACTIVARSE'].sum())
                tot_pct_i2 = (tot_i2 / tot_disp_i2 * 100.0) if tot_disp_i2 > 0 else 0.0
                tot_av_i2 = int(df_i2_calc['AVANCE RETENCION I2'].sum())

                row_tot_i2 = pd.DataFrame([{
                    'LÍDER DE NEGOCIOS': 'TOTAL GENERAL',
                    'META RETENCIÓN I2': tot_meta_i2,
                    'FALTA I2 ACTIVARSE': tot_falta_i2,
                    '% RETENCIÓN META 8%': tot_pct_i2,
                    'AVANCE RETENCION I2': tot_av_i2
                }])
                df_i2_final = pd.concat([df_i2_calc, row_tot_i2], ignore_index=True)

                df_i2_formatted = df_i2_final[['LÍDER DE NEGOCIOS', 'META RETENCIÓN I2', 'FALTA I2 ACTIVARSE', '% RETENCIÓN META 8%', 'AVANCE RETENCION I2']].copy()
                df_i2_formatted['META RETENCIÓN I2'] = df_i2_formatted['META RETENCIÓN I2'].apply(lambda v: f"{int(v):,}".replace(",", "."))
                df_i2_formatted['FALTA I2 ACTIVARSE'] = df_i2_formatted['FALTA I2 ACTIVARSE'].apply(lambda v: f"{int(v):,}".replace(",", "."))
                df_i2_formatted['% RETENCIÓN META 8%'] = df_i2_formatted['% RETENCIÓN META 8%'].apply(lambda v: f"{v:.1f}%")
                df_i2_formatted['AVANCE RETENCION I2'] = df_i2_formatted['AVANCE RETENCION I2'].apply(lambda v: f"{int(v):,}".replace(",", "."))

                def _estilo_falta_retencion(val_str):
                    try:
                        num = int(str(val_str).replace('.', '').strip())
                        if num <= 0:
                            return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                        elif num <= 5:
                            return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                        else:
                            return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
                    except Exception:
                        return ''

                def _estilo_pct_retencion_8(val_str):
                    try:
                        num = float(str(val_str).replace('%', '').strip())
                        if num <= 8.0:
                            return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                        elif num <= 10.0:
                            return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                        else:
                            return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
                    except Exception:
                        return ''

                styler_i2 = df_i2_formatted.style
                if hasattr(styler_i2, 'map'):
                    styler_i2 = styler_i2.map(_estilo_falta_retencion, subset=['FALTA I2 ACTIVARSE']).map(_estilo_pct_retencion_8, subset=['% RETENCIÓN META 8%'])
                elif hasattr(styler_i2, 'applymap'):
                    styler_i2 = styler_i2.applymap(_estilo_falta_retencion, subset=['FALTA I2 ACTIVARSE']).applymap(_estilo_pct_retencion_8, subset=['% RETENCIÓN META 8%'])

                st.dataframe(styler_i2, use_container_width=True)

st.markdown("---")
st.markdown("<p style='text-align:center; font-size:10px; color:#94a3b8; margin:0;'>App Matices Móvil • Natura & Avon • Diseñada para Celulares y Tablets</p>", unsafe_allow_html=True)


