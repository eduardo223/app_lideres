import streamlit as st
import pandas as pd
import urllib.parse
import os

import procesador
from procesador import (
    autenticar_usuario,
    cargar_usuarios,
    consultar_tableau_sql,
    consultar_geral_sql,
    calcular_metas_ciclo,
    limpiar_numero,
    cambiar_password_usuario,
    extraer_telefonos_colombia
)

# 1. Configuración de página ultra-optimizada para Celulares y Tablets
st.set_page_config(
    page_title="App Matices - Móvil",
    page_icon="📱",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. CSS Ultra-Compacto y Responsivo para Smartphones y Tablets
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
    
    /* Pestañas compactas estilo App móvil */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px !important;
        background-color: rgba(15, 23, 42, 0.05);
        border-radius: 12px;
        padding: 3px;
        margin-bottom: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 11.5px !important;
        font-weight: 700 !important;
        padding: 6px 8px !important;
        border-radius: 8px;
        color: #64748b;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0284c7 !important;
        color: #ffffff !important;
    }

    /* Cards KPI 2x2 Condensadas */
    .kpi-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 6px;
        margin-bottom: 8px;
    }
    .kpi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 8px 10px;
        text-align: center;
        box-shadow: 0 1px 4px rgba(0,0,0,0.04);
    }
    .kpi-title {
        font-size: 10px;
        color: #64748b;
        text-transform: uppercase;
        font-weight: 700;
        margin-bottom: 2px;
    }
    .kpi-val {
        font-size: 15px;
        font-weight: 800;
        color: #0f172a;
        line-height: 1.2;
    }
    .kpi-sub {
        font-size: 10px;
        font-weight: 600;
        margin-top: 2px;
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
        padding: 3px 8px;
        border-radius: 6px;
        text-decoration: none !important;
    }

    /* Badges de estado */
    .badge-pill {
        display: inline-block;
        font-size: 10px;
        font-weight: 700;
        padding: 2px 6px;
        border-radius: 9999px;
    }
    
    /* Header compacto */
    .header-mobile {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 4px 0 8px 0;
        border-bottom: 1px solid #e2e8f0;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

# 3. Control de Sesión y Login Móvil
if 'user' not in st.session_state:
    st.session_state['user'] = None

# Auto-login por parámetro de URL si existe
if st.session_state['user'] is None:
    session_param = st.query_params.get('user')
    if session_param:
        todos_u = cargar_usuarios()
        u_cl = str(session_param).strip().lower()
        if u_cl in todos_u:
            u_data = todos_u[u_cl].copy()
            u_data['username'] = u_cl
            st.session_state['user'] = u_data

# Pantalla de Login Móvil si no está autenticado
if st.session_state['user'] is None:
    st.markdown("<h3 style='text-align:center; margin-bottom:2px;'>📱 App Matices Móvil</h3>", unsafe_allow_html=True)
    st.caption("<p style='text-align:center; margin-bottom:12px;'>Portal exclusivo para Líderes de Negocio en Celulares y Tablets</p>", unsafe_allow_html=True)
    
    with st.form("form_login_mobile"):
        input_u = st.text_input("👤 Usuario o Correo", placeholder="Ej. lider9334 o correo@...").strip().lower()
        input_p = st.text_input("🔒 Contraseña", type="password", placeholder="••••••••")
        btn_log = st.form_submit_button("🚀 Entrar al Sistema", type="primary", use_container_width=True)
        
        if btn_log:
            u_auth = autenticar_usuario(input_u, input_p)
            if u_auth:
                st.session_state['user'] = u_auth
                st.query_params['user'] = u_auth['username']
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

# 5. Carga de Datos Relacionales (Tableau, Geral y Cómo Vamos)
df_tab = consultar_tableau_sql(grupo=grupo_activo if grupo_activo else None)
df_geral = consultar_geral_sql(grupo=grupo_activo if grupo_activo else None)

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

# 6. Pestañas Principales Móviles (3 Pestañas Condensadas)
tab_tab, tab_geral, tab_cv = st.tabs([
    "📋 Red Tableau",
    "💳 Cobranza Hoy",
    "📈 Cómo Vamos"
])

# ==============================================================================
# TAB 1: RED TABLEAU (ULTRA-RESUMIDO CON SOLO 2 FILTROS)
# ==============================================================================
with tab_tab:
    # 2 Filtros Esenciales
    f_c1, f_c2 = st.columns(2)
    with f_c1:
        opciones_sit = ["Todas"]
        if not df_tab.empty and 'Sit. Comercial' in df_tab.columns:
            sits_unicas = [str(x) for x in df_tab['Sit. Comercial'].dropna().unique() if str(x).strip()]
            opciones_sit += sorted(sits_unicas)
        filtro_sit = st.selectbox("🏷️ Sit. Comercial", options=opciones_sit, key="mob_f_sit")
        
    with f_c2:
        filtro_mora = st.selectbox("💳 Deuda en Mora", options=["Todas", "🔴 Solo con Mora", "🟢 Al Día"], key="mob_f_mora")

    # Búsqueda rápida por nombre
    busq_nom = st.text_input("🔍 Buscar consultora...", placeholder="Escribe un nombre...", key="mob_b_nom").strip()

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
                
        if busq_nom and 'Asesora / Consultora' in df_tab_filtrado.columns:
            df_tab_filtrado = df_tab_filtrado[df_tab_filtrado['Asesora / Consultora'].astype(str).str.contains(busq_nom, case=False, na=False)]

    # Mini Resumen
    tot_c = len(df_tab_filtrado)
    
    col_sit_check_m = 'Sit. Comercial' if 'Sit. Comercial' in df_tab_filtrado.columns else ('Situación' if 'Situación' in df_tab_filtrado.columns else None)
    if col_sit_check_m and not df_tab_filtrado.empty:
        s_vals_lower_m = df_tab_filtrado[col_sit_check_m].astype(str).str.strip().str.lower()
        mask_disp_m = s_vals_lower_m.apply(
            lambda s: any(k in s for k in ['activa', 'activas', 'inactiva 1', 'inactiva 2', 'inactiva 3', 'i1', 'i2', 'i3']) and not any(k in s for k in ['inactiva 4', 'inactiva 5', 'inactiva 6', 'i4', 'i5', 'i6'])
        )
        tot_disp_m = int(mask_disp_m.sum())
    else:
        tot_disp_m = 0

    tot_mora_cop = float(df_tab_filtrado['Deuda_Num'].sum()) if 'Deuda_Num' in df_tab_filtrado.columns else 0.0
    
    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-card">
            <div class="kpi-title">Consultoras</div>
            <div class="kpi-val">{tot_c}</div>
            <div class="kpi-sub" style="color:#0284c7;">🎯 {tot_disp_m} Disponibles</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-title">Deuda Mora</div>
            <div class="kpi-val" style="color:{'#ef4444' if tot_mora_cop > 0 else '#10b981'};">${tot_mora_cop:,.0f}</div>
            <div class="kpi-sub" style="color:#64748b;">Filtro Activo</div>
        </div>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)

    if df_tab_filtrado.empty:
        st.info("ℹ️ No hay consultoras con los filtros seleccionados.")
    else:
        # Renderizado de lista compacta en tarjetas móviles
        for _, row in df_tab_filtrado.head(60).iterrows():
            nom = str(row.get('Asesora / Consultora', 'Sin Nombre')).strip()
            sit = str(row.get('Sit. Comercial', 'N/D')).strip()
            pts = int(limpiar_numero(row.get('Pts Acum', 0)))
            mora = float(limpiar_numero(row.get('Deuda Mora', 0.0)))
            cel_raw = str(row.get('celular', '')).strip()
            m1, _ = extraer_telefonos_colombia(cel_raw)
            
            # Badge de situación comercial
            if 'activa' in sit.lower():
                color_bg, color_fg = "#dcfce7", "#15803d"
            elif '1' in sit:
                color_bg, color_fg = "#fef9c3", "#a16207"
            else:
                color_bg, color_fg = "#fee2e2", "#b91c1c"

            # Enlace de WhatsApp
            wa_btn_html = ""
            if m1 and len(m1) == 10:
                msg_wa = urllib.parse.quote(f"Hola {nom}, te saludo de tu equipo Natura & Avon. ¿Cómo estás?")
                wa_url = f"https://wa.me/57{m1}?text={msg_wa}"
                wa_btn_html = f'<a href="{wa_url}" target="_blank" class="btn-wa-link">📲 WhatsApp</a>'

            st.markdown(f"""
            <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:6px 10px; margin-bottom:4px; display:flex; align-items:center; justify-content:space-between;">
                <div style="flex:1; min-width:0; padding-right:6px;">
                    <div style="font-size:12px; font-weight:700; color:#0f172a; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{nom}</div>
                    <div style="font-size:10px; color:#64748b; margin-top:1px;">
                        <span class="badge-pill" style="background:{color_bg}; color:{color_fg};">{sit}</span>
                        <span style="margin-left:4px;">⭐ {pts} pts</span>
                        {f'<span style="color:#ef4444; font-weight:700; margin-left:4px;">• Mora: ${mora:,.0f}</span>' if mora > 0 else ''}
                    </div>
                </div>
                <div>{wa_btn_html}</div>
            </div>
            """.replace(",", "."), unsafe_allow_html=True)
            
        if len(df_tab_filtrado) > 60:
            st.caption(f"Mostrando 60 de {len(df_tab_filtrado)} consultoras. Usa el buscador para filtrar.")

# ==============================================================================
# TAB 2: CRÉDITO & COBRANZA GERAL (SOLO SI HAY DEUDA DEL DÍA - 3 COLUMNAS)
# ==============================================================================
with tab_geral:
    st.markdown("##### 💳 Cartera & Cobranza Prioritaria")
    
    df_g_pend = pd.DataFrame()
    if not df_geral.empty:
        df_geral['saldo_num'] = df_geral['saldo_total'].apply(lambda x: limpiar_numero(x, 0.0))
        # Filtrar solo casos con saldo pendiente mayor a 0
        df_g_pend = df_geral[df_geral['saldo_num'] > 0].copy()
        
    if df_g_pend.empty:
        st.success("🎉 **¡Excelente!** Tu grupo no tiene deudas en mora pendientes de cobro el día de hoy.")
    else:
        # Ordenar de mayor a menor deuda
        df_g_pend = df_g_pend.sort_values(by='saldo_num', ascending=False)
        tot_deuda_g = float(df_g_pend['saldo_num'].sum())
        
        st.markdown(f"""
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-title">Casos en Cobro</div>
                <div class="kpi-val" style="color:#ef4444;">{len(df_g_pend)}</div>
                <div class="kpi-sub">Títulos con saldo</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Total a Cobrar</div>
                <div class="kpi-val" style="color:#ef4444;">${tot_deuda_g:,.0f}</div>
                <div class="kpi-sub">Cartera activa</div>
            </div>
        </div>
        """.replace(",", "."), unsafe_allow_html=True)

        st.caption("📋 **Listado de Cobranza (Nombre • Deuda • Vencimiento & Botón WhatsApp):**")

        for _, row in df_g_pend.iterrows():
            c_nom = str(row.get('nombre', 'Sin Nombre')).strip()
            c_saldo = float(row.get('saldo_num', 0.0))
            c_venc = str(row.get('fecha_vencimiento', ''))[:10]
            c_ped = str(row.get('numero_pedido', '')).strip().split('.')[0]
            c_movil1 = str(row.get('telefono_movil', '')).strip()
            
            # Formato de mensaje WhatsApp de cobro amable
            wa_cobro_html = ""
            if c_movil1 and len(c_movil1) == 10:
                msg_cobro = urllib.parse.quote(
                    f"Hola {c_nom}, cordial saludo de tu Líder Natura & Avon. "
                    f"Te recuerdo que tienes un saldo pendiente de ${c_saldo:,.0f} COP (Pedido #{c_ped}) "
                    f"con fecha de vencimiento {c_venc}. Por favor confírmanos tu pago para mantener tu crédito al día. ¡Gracias!"
                )
                wa_cobro_url = f"https://wa.me/57{c_movil1}?text={msg_cobro}"
                wa_cobro_html = f'<a href="{wa_cobro_url}" target="_blank" class="btn-wa-link" style="background:#ef4444;">📲 Cobrar</a>'

            st.markdown(f"""
            <div style="background:#ffffff; border:1px solid #fee2e2; border-radius:10px; padding:8px 10px; margin-bottom:5px; display:flex; align-items:center; justify-content:space-between;">
                <div style="flex:1; min-width:0; padding-right:6px;">
                    <div style="font-size:12px; font-weight:700; color:#0f172a; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{c_nom}</div>
                    <div style="font-size:11px; margin-top:2px;">
                        <span style="color:#b91c1c; font-weight:800;">${c_saldo:,.0f} COP</span>
                        <span style="color:#64748b; font-size:10px; margin-left:4px;">• Vence: {c_venc}</span>
                    </div>
                </div>
                <div>{wa_cobro_html}</div>
            </div>
            """.replace(",", "."), unsafe_allow_html=True)

# ==============================================================================
# TAB 3: CÓMO VAMOS (FACTURACIÓN, CUMPLIMIENTO & SIMULADOR COMPACTO)
# ==============================================================================
with tab_cv:
    st.markdown("##### 📈 Desempeño & Cómo Vamos")
    
    if df_cv.empty:
        st.info(f"ℹ️ **Metas del ciclo pendientes para el Grupo {grupo_activo if grupo_activo else ''}:**\n\nEl archivo actual de metas aún no contiene datos registrados para este grupo en el ciclo actual. Puedes gestionar tu red en **'📋 Red Tableau'** y consultar tu cartera en **'💳 Cobranza Hoy'**.")
    else:
        row_cv = df_cv.iloc[0]
        
        f_real = float(limpiar_numero(row_cv.get('Real Facturación', 0.0)))
        f_obj = float(limpiar_numero(row_cv.get('Objetivo Facturación', 0.0)))
        cump_f = (f_real / f_obj * 100.0) if f_obj > 0 else 0.0
        
        a_real = int(limpiar_numero(row_cv.get('Real Activas', 0)))
        a_obj = int(limpiar_numero(row_cv.get('Objetivo Activas', 0)))
        cump_a = (a_real / a_obj * 100.0) if a_obj > 0 else 0.0
        
        saldo_cv = int(limpiar_numero(row_cv.get('Saldo', 0)))
        inicios_cv = int(limpiar_numero(row_cv.get('Inicios', 0)))
        reinicios_cv = int(limpiar_numero(row_cv.get('Reinicios', 0)))
        ganancia_cop = float(limpiar_numero(row_cv.get('Ganancia estimada', 0.0)))

        # Tarjetas 2x2 Ultra-Compactas
        st.markdown(f"""
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-title">Facturación</div>
                <div class="kpi-val">${f_real:,.0f}</div>
                <div class="kpi-sub" style="color:{'#10b981' if cump_f >= 100 else '#f59e0b'};">
                    {cump_f:.1f}% Cump. (Meta: ${f_obj:,.0f})
                </div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Activas Reales</div>
                <div class="kpi-val">{a_real} / {a_obj}</div>
                <div class="kpi-sub" style="color:{'#10b981' if cump_a >= 100 else '#f59e0b'};">
                    {cump_a:.1f}% Cumplimiento
                </div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Saldo Comercial</div>
                <div class="kpi-val" style="color:{'#10b981' if saldo_cv >= 2 else '#ef4444'};">{saldo_cv:+d}</div>
                <div class="kpi-sub" style="color:#64748b;">Inicios: {inicios_cv} • Reinicios: {reinicios_cv}</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-title">Ganancia Estimada</div>
                <div class="kpi-val" style="color:#0284c7;">${ganancia_cop:,.0f}</div>
                <div class="kpi-sub" style="color:#10b981;">Comisión + Potencializador</div>
            </div>
        </div>
        """.replace(",", "."), unsafe_allow_html=True)

        # Barra de progreso visual de Facturación
        st.caption(f"🎯 **Avance de Facturación:** ({cump_f:.1f}% del objetivo)")
        st.progress(min(1.0, cump_f / 100.0))

st.markdown("---")
st.markdown("<p style='text-align:center; font-size:10px; color:#94a3b8; margin:0;'>App Matices Móvil • Diseñada para Celulares y Tablets</p>", unsafe_allow_html=True)

