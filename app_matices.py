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
    extraer_telefonos_colombia
)

# 1. Configuración de página ultra-optimizada para Celulares y Tablets
st.set_page_config(
    page_title="App Matices - Móvil",
    page_icon="📱",
    layout="centered",
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
tab_cv, tab_tab, tab_geral = st.tabs([
    "🎯 MIS DESAFÍOS",
    "📋 MI LISTADO",
    "💳 COBRANZA HOY"
])

# ==============================================================================
# TAB 1: MIS DESAFÍOS & CÓMO VAMOS (OBJETIVOS ARTE + CÓMO VAMOS)
# ==============================================================================
with tab_cv:
    st.markdown("##### 🎯 Cuadro de Mando de Desafíos Líder")
    
    if df_cv.empty:
        st.info(f"ℹ️ **Metas del ciclo para el Grupo {grupo_activo if grupo_activo else ''}:**\n\nEl archivo de metas se sincronizará automáticamente. Puedes gestionar tu red en **'📋 MI LISTADO'** y tu cartera en **'💳 COBRANZA HOY'**.")
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
# TAB 2: MI LISTADO (ULTRA-RESUMIDO CON FILTROS, DOCUMENTO GPP, CRÉDITO Y CELULAR)
# ==============================================================================
with tab_tab:
    # Filtros Esenciales
    f_c1, f_c2 = st.columns(2)
    with f_c1:
        opciones_sit = ["Todas"]
        if not df_tab.empty and 'Sit. Comercial' in df_tab.columns:
            sits_unicas = [str(x) for x in df_tab['Sit. Comercial'].dropna().unique() if str(x).strip()]
            opciones_sit += sorted(sits_unicas)
        filtro_sit = st.selectbox("🏷️ Sit. Comercial", options=opciones_sit, key="mob_f_sit")
        
    with f_c2:
        filtro_mora = st.selectbox("💳 Deuda en Mora", options=["Todas", "🔴 Solo con Mora", "🟢 Al Día"], key="mob_f_mora")

    # Búsqueda rápida por nombre o documento
    busq_nom = st.text_input("🔍 Buscar consultora o documento...", placeholder="Escribe un nombre o cédula...", key="mob_b_nom").strip()

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
                
        if busq_nom:
            mask_busq = pd.Series(False, index=df_tab_filtrado.index)
            if 'Asesora / Consultora' in df_tab_filtrado.columns:
                mask_busq = mask_busq | df_tab_filtrado['Asesora / Consultora'].astype(str).str.contains(busq_nom, case=False, na=False)
            if 'DocumentoGPP' in df_tab_filtrado.columns:
                mask_busq = mask_busq | df_tab_filtrado['DocumentoGPP'].astype(str).str.contains(busq_nom, case=False, na=False)
            if 'Código CB' in df_tab_filtrado.columns:
                mask_busq = mask_busq | df_tab_filtrado['Código CB'].astype(str).str.contains(busq_nom, case=False, na=False)
            df_tab_filtrado = df_tab_filtrado[mask_busq]

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
    tot_mora_fmt = f"${tot_mora_cop:,.0f}".replace(",", ".")
    
    st.markdown(f"""<div class="kpi-grid">
<div class="kpi-card">
<div class="kpi-title">Consultoras</div>
<div class="kpi-val">{tot_c}</div>
<div class="kpi-sub kpi-sub-blue">🎯 {tot_disp_m} Disponibles</div>
</div>
<div class="kpi-card">
<div class="kpi-title">Deuda Mora</div>
<div class="kpi-val" style="color:{'#ef4444' if tot_mora_cop > 0 else '#10b981'};">{tot_mora_fmt}</div>
<div class="kpi-sub {'kpi-sub-red' if tot_mora_cop > 0 else 'kpi-sub-green'}">Filtro Activo</div>
</div>
</div>""", unsafe_allow_html=True)

    if df_tab_filtrado.empty:
        st.info("ℹ️ No hay consultoras con los filtros seleccionados.")
    else:
        # Renderizado de lista compacta en tarjetas móviles
        for _, row in df_tab_filtrado.head(60).iterrows():
            nom = str(row.get('Asesora / Consultora', 'Sin Nombre')).strip()
            doc_gpp = str(row.get('DocumentoGPP', '')).replace('.0', '').strip()
            sit = str(row.get('Sit. Comercial', 'N/D')).strip()
            pts = int(limpiar_numero(row.get('Pts Acum', 0)))
            cred_disp = int(limpiar_numero(row.get('Credito Disponible', 0)))
            mora = float(limpiar_numero(row.get('Deuda Mora', 0.0)))
            ped_pend = int(limpiar_numero(row.get('Ped. Pendientes', 0)))
            cel_raw = str(row.get('celular', row.get('Celular', ''))).strip()
            m1, _ = extraer_telefonos_colombia(cel_raw)
            
            # Badge de situación comercial
            if 'activa' in sit.lower():
                color_bg, color_fg = "#dcfce7", "#15803d"
            elif '1' in sit:
                color_bg, color_fg = "#fef9c3", "#a16207"
            else:
                color_bg, color_fg = "#fee2e2", "#b91c1c"

            # Formatos de texto
            mora_fmt = f"${mora:,.0f}".replace(",", ".")
            doc_html = f'<span style="font-size:10px; color:#64748b; font-weight:600; margin-left:6px;">Doc: {doc_gpp}</span>' if doc_gpp and doc_gpp not in ['nan', 'None', '0'] else ''
            cred_html = f'<span style="margin-left:4px; color:#0284c7; font-weight:700;">• Créd: {cred_disp} pts</span>' if cred_disp > 0 else ''
            mora_html = f'<span style="color:#ef4444; font-weight:700; margin-left:4px;">• Mora: {mora_fmt}</span>' if mora > 0 else ''
            ped_html = f'<span style="color:#ca8a04; font-weight:700; margin-left:4px;">• ⌛ Ped. Pend: {ped_pend}</span>' if ped_pend > 0 else ''
            tel_html = f'<span style="color:#475569; font-size:10px; margin-left:4px;">• 📱 {m1}</span>' if m1 else ''

            # Enlace de WhatsApp
            wa_btn_html = ""
            if m1 and len(m1) == 10:
                msg_wa = urllib.parse.quote(f"Hola {nom}, te saludo de tu equipo Natura & Avon. ¿Cómo estás?")
                wa_url = f"https://wa.me/57{m1}?text={msg_wa}"
                wa_btn_html = f'<a href="{wa_url}" target="_blank" class="btn-wa-link">📲 WA</a>'

            card_html = (
                f'<div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:8px 10px; margin-bottom:6px; display:flex; align-items:center; justify-content:space-between; box-shadow: 0 1px 3px rgba(0,0,0,0.04);">'
                f'<div style="flex:1; min-width:0; padding-right:6px;">'
                f'<div style="font-size:12px; font-weight:700; color:#0f172a; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{nom}{doc_html}</div>'
                f'<div style="font-size:10px; color:#64748b; margin-top:2px;">'
                f'<span class="badge-pill" style="background:{color_bg}; color:{color_fg};">{sit}</span>'
                f'<span style="margin-left:4px;">⭐ {pts} pts</span>'
                f'{cred_html}'
                f'{mora_html}'
                f'{ped_html}'
                f'{tel_html}'
                f'</div>'
                f'</div>'
                f'<div>{wa_btn_html}</div>'
                f'</div>'
            )
            st.markdown(card_html, unsafe_allow_html=True)
            
        if len(df_tab_filtrado) > 60:
            st.caption(f"Mostrando 60 de {len(df_tab_filtrado)} consultoras. Usa el buscador para filtrar.")

# ==============================================================================
# TAB 3: CRÉDITO & COBRANZA GERAL (SOLO SI HAY DEUDA DEL DÍA - DIRECTO A WHATSAPP)
# ==============================================================================
with tab_geral:
    st.markdown("##### 💳 Cartera & Cobranza Prioritaria")
    
    df_g_pend = pd.DataFrame()
    if not df_geral.empty:
        df_geral['saldo_num'] = df_geral['saldo_total'].apply(lambda x: limpiar_numero(x, 0.0))
        df_g_pend = df_geral[df_geral['saldo_num'] > 0].copy()
        
    if df_g_pend.empty:
        st.success("🎉 **¡Excelente!** Tu grupo no tiene deudas en mora pendientes de cobro el día de hoy.")
    else:
        df_g_pend = df_g_pend.sort_values(by='saldo_num', ascending=False)
        tot_deuda_g = float(df_g_pend['saldo_num'].sum())
        tot_deuda_g_fmt = f"${tot_deuda_g:,.0f}".replace(",", ".")
        
        st.markdown(f"""<div class="kpi-grid">
<div class="kpi-card">
<div class="kpi-title">Casos en Cobro</div>
<div class="kpi-val" style="color:#ef4444;">{len(df_g_pend)}</div>
<div class="kpi-sub kpi-sub-red">Títulos con saldo</div>
</div>
<div class="kpi-card">
<div class="kpi-title">Total a Cobrar</div>
<div class="kpi-val" style="color:#ef4444;">{tot_deuda_g_fmt}</div>
<div class="kpi-sub kpi-sub-red">Cartera activa</div>
</div>
</div>""", unsafe_allow_html=True)

        st.caption("📋 **Listado de Cobranza (Nombre • Deuda • Vencimiento & Botón WhatsApp):**")

        for _, row in df_g_pend.iterrows():
            c_nom = str(row.get('nombre', 'Sin Nombre')).strip()
            c_saldo = float(row.get('saldo_num', 0.0))
            c_venc = str(row.get('fecha_vencimiento', ''))[:10]
            c_ped = str(row.get('numero_pedido', '')).strip().split('.')[0]
            c_movil1 = str(row.get('telefono_movil', '')).strip()
            saldo_row_fmt = f"${c_saldo:,.0f}".replace(",", ".")
            
            wa_cobro_html = ""
            if c_movil1 and len(c_movil1) == 10:
                msg_cobro = urllib.parse.quote(
                    f"Hola {c_nom}, cordial saludo de tu Líder Natura & Avon. "
                    f"Te recuerdo que tienes un saldo pendiente de {saldo_row_fmt} COP (Pedido #{c_ped}) "
                    f"con fecha de vencimiento {c_venc}. Por favor confírmanos tu pago para mantener tu crédito al día. ¡Gracias!"
                )
                wa_cobro_url = f"https://wa.me/57{c_movil1}?text={msg_cobro}"
                wa_cobro_html = f'<a href="{wa_cobro_url}" target="_blank" class="btn-wa-link" style="background:#ef4444;">📲 Cobrar</a>'

            cobro_card_html = (
                f'<div style="background:#ffffff; border:1px solid #fee2e2; border-radius:10px; padding:8px 10px; margin-bottom:6px; display:flex; align-items:center; justify-content:space-between; box-shadow: 0 1px 3px rgba(0,0,0,0.04);">'
                f'<div style="flex:1; min-width:0; padding-right:6px;">'
                f'<div style="font-size:12px; font-weight:700; color:#0f172a; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">{c_nom}</div>'
                f'<div style="font-size:11px; margin-top:2px;">'
                f'<span style="color:#b91c1c; font-weight:800;">{saldo_row_fmt} COP</span>'
                f'<span style="color:#64748b; font-size:10px; margin-left:4px;">• Vence: {c_venc}</span>'
                f'</div>'
                f'</div>'
                f'<div>{wa_cobro_html}</div>'
                f'</div>'
            )
            st.markdown(cobro_card_html, unsafe_allow_html=True)

st.markdown("---")
st.markdown("<p style='text-align:center; font-size:10px; color:#94a3b8; margin:0;'>App Matices Móvil • Natura & Avon • Diseñada para Celulares y Tablets</p>", unsafe_allow_html=True)


