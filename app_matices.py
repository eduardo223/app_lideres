import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import importlib
import contextlib

import procesador
importlib.reload(procesador)

from procesador import (
    autenticar_usuario,
    cargar_usuarios,
    consultar_tableau_sql,
    limpiar_numero,
    cambiar_password_usuario,
    cargar_configuracion
)

formato_cop = lambda v: f"${float(limpiar_numero(v, 0)):,.0f}".replace(",", ".")

# Configuración de página optimizada para Smartphones (Centered layout, sidebar colapsada)
st.set_page_config(
    page_title="App Matices - Móvil",
    page_icon="📱",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Recuperación de sesión / F5 Browser Refresh
query_params = st.query_params
if 'user' in query_params and ('user' not in st.session_state or st.session_state['user'] is None):
    uname = query_params['user']
    usuarios = cargar_usuarios()
    if uname in usuarios:
        user_rec = usuarios[uname].copy()
        user_rec['username'] = uname
        st.session_state['user'] = user_rec

if 'user' not in st.session_state or st.session_state['user'] is None:
    st.markdown("<h2 style='text-align:center;'>📱 App Matices Móvil</h2>", unsafe_allow_html=True)
    st.caption("Ingresa con tus credenciales para acceder a la vista rápida desde tu teléfono.")
    
    with st.form("form_login_mobile"):
        input_user = st.text_input("Usuario (Login o Correo)").strip().lower()
        input_pass = st.text_input("Contraseña", type="password")
        btn_login = st.form_submit_button("🚀 Iniciar Sesión en App Móvil", use_container_width=True)
        
        if btn_login:
            user_auth = autenticar_usuario(input_user, input_pass)
            if user_auth:
                st.session_state['user'] = user_auth
                st.query_params['user'] = user_auth['username']
                st.rerun()
            else:
                st.error("❌ Usuario o contraseña incorrectos.")
    st.stop()

current_user = st.session_state.get('user') or {}
user_nombre = current_user.get('nombre', 'Usuario')
user_rol = current_user.get('rol', 'asesor')
user_grupo = str(current_user.get('codigo_grupo', '')).strip() if current_user.get('codigo_grupo') else ""
user_sector = str(current_user.get('codigo_sector', '')).strip() if current_user.get('codigo_sector') else ""

# Guard de Cambio Obligatorio de Contraseña
if current_user.get('debe_cambiar_password', False):
    st.warning("🔒 **Cambio Obligatorio de Contraseña Inicial**")
    st.info("Por seguridad, debes crear una nueva contraseña para continuar.")
    with st.form("form_pwd_mobile"):
        pn = st.text_input("Nueva Contraseña", type="password")
        pc = st.text_input("Confirmar Nueva Contraseña", type="password")
        if st.form_submit_button("💾 Guardar y Continuar", use_container_width=True):
            if len(pn.strip()) < 4 or pn != pc:
                st.error("Verifica que las contraseñas coincidan y tengan al menos 4 caracteres.")
            else:
                ok, msg = cambiar_password_usuario(current_user['username'], pn)
                if ok:
                    current_user['debe_cambiar_password'] = False
                    st.session_state['user'] = current_user
                    st.rerun()
    st.stop()

# Header de App Móvil
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.markdown(f"### 📱 App Matices")
    st.caption(f"👤 **{user_nombre}** ({user_rol.capitalize()})")
with col_h2:
    if st.button("🚪 Salir", key="btn_logout_mob"):
        st.session_state['user'] = None
        if 'user' in st.query_params:
            del st.query_params['user']
        st.rerun()

# Selector de Tema Móvil
modo_tema = st.radio("🎨 Tema", options=["🌙 Oscuro Neón", "☀️ Modo Claro"], horizontal=True, key="mob_theme")
is_dark = (modo_tema == "🌙 Oscuro Neón")

if is_dark:
    st.markdown("""
    <style>
    .stApp { background-color: #0b0f19 !important; color: #f8fafc !important; }
    .mob-card {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.95), rgba(30, 41, 59, 0.95));
        border: 1px solid rgba(56, 189, 248, 0.3);
        border-radius: 16px;
        padding: 14px;
        text-align: center;
        margin-bottom: 10px;
    }
    .mob-title { font-size: 11px; color: #94a3b8; text-transform: uppercase; font-weight: 700; }
    .mob-val { font-size: 22px; font-weight: 800; color: #38bdf8; }
    .mob-sub { font-size: 11px; color: #10b981; font-weight: 600; }
    </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <style>
    .stApp { background-color: #f8fafc !important; color: #0f172a !important; }
    .mob-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 14px;
        text-align: center;
        margin-bottom: 10px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }
    .mob-title { font-size: 11px; color: #64748b; text-transform: uppercase; font-weight: 700; }
    .mob-val { font-size: 22px; font-weight: 800; color: #0284c7; }
    .mob-sub { font-size: 11px; color: #059669; font-weight: 600; }
    </style>
    """, unsafe_allow_html=True)

# Cargar Datos
df_tab = consultar_tableau_sql(
    grupo=(user_grupo if user_rol == 'lider' else None),
    sector=(user_sector if (user_rol == 'gerente' and user_sector) else ('__INVALID_SECTOR__' if user_rol == 'gerente' else None))
)

pts_tot = 0
deuda_tot = 0.0

if not df_tab.empty:
    if 'Pts Acum' in df_tab.columns:
        pts_tot = int(df_tab['Pts Acum'].apply(lambda x: limpiar_numero(x, 0)).sum())
    if 'Deuda Mora' in df_tab.columns:
        deuda_tot = float(df_tab['Deuda Mora'].apply(lambda x: limpiar_numero(x, 0)).sum())

# Hero Cards 2x2 para Smartphone
m1, m2 = st.columns(2)
with m1:
    st.markdown(f"""
    <div class="mob-card">
        <div class="mob-title">👥 Consultoras</div>
        <div class="mob-val">{len(df_tab)}</div>
        <div class="mob-sub">Red Asignada</div>
    </div>
    """, unsafe_allow_html=True)
with m2:
    st.markdown(f"""
    <div class="mob-card">
        <div class="mob-title">⭐ Puntos / Mora</div>
        <div class="mob-val">{pts_tot:,} pts</div>
        <div class="mob-sub" style="color:#ef4444;">Mora: ${deuda_tot/1e6:.1f}M</div>
    </div>
    """.replace(",", "."), unsafe_allow_html=True)

# 3 PESTAÑAS MÓVILES APPS
t_mob1, t_mob2, t_mob3 = st.tabs([
    "📱 Puntos & Listado",
    "⚡ KPIs Tacómetros",
    "📈 Tabla Facturación"
])

with t_mob1:
    st.subheader("📋 Consultoras & Puntos")
    if df_tab.empty:
        st.info("No hay datos cargados.")
    else:
        cols_m = [c for c in ['Código CB', 'Líder / Grupo', 'Asesora / Consultora', 'Nivel / Color', 'Sit. Comercial', 'Pts Acum', 'Deuda Mora', 'Ped. Pendientes'] if c in df_tab.columns]
        st.dataframe(df_tab[cols_m] if cols_m else df_tab, use_container_width=True, height=400)

with t_mob2:
    st.subheader("⏱️ Medidores de Avance")
    fig_g = go.Figure(go.Indicator(
        mode="gauge+number",
        value=85.5,
        number={'suffix': '%'},
        title={'text': "Avance General (%)"},
        gauge={'axis': {'range': [0, 150]}, 'bar': {'color': "#38bdf8"}}
    ))
    fig_g.update_layout(height=260, margin=dict(l=10, r=10, t=30, b=10))
    st.plotly_chart(fig_g, use_container_width=True)

with t_mob3:
    st.subheader("📈 Ranking de Facturación")
    st.info("Vista optimizada de facturación para dispositivos móviles.")

st.markdown("---")
st.markdown("<p style='text-align:center; font-size:12px; color:#94a3b8;'>App Matices Móvil • Diseñada para Teléfonos Inteligentes</p>", unsafe_allow_html=True)
