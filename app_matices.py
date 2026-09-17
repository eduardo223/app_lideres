import streamlit as st
import pandas as pd
import numpy as np
pd.set_option("styler.render.max_elements", 3_000_000)
import urllib.parse
import os
import io
import time
import base64
from datetime import datetime
import requests

import procesador
from procesador import (
    autenticar_usuario,
    generar_token_sesion,
    validar_token_sesion,
    cargar_usuarios,
    cargar_objetivos_arte,
    obtener_metas_efectivas,
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
    color_pts_cierre_anterior,
    guardar_todos_comentarios,
    limpiar_codigo_cb_estandar,
    refrescar_perfil_usuario_en_sesion,
    registrar_evento_auditoria,
    obtener_nombre_sector_usuario,
    obtener_nombre_corto_sector,
    ruta_persistente,
    consultar_ce_plus_df,
    cargar_datos_ganancia_arte,
    obtener_cumpleanos_equipo,
    obtener_conexion_db
)
from ui_suscripcion import modal_pago_suscripcion_gerente, banner_alerta_suscripcion_gerente


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

def render_vista_movil(current_user=None, mostrar_salir=False):
    """
    Renderiza la interfaz móvil de App Matices optimizada para celulares y tablets.
    Puede ser llamada desde app.py (en el mismo contenedor) o de forma independiente.
    """
    if current_user is None:
        current_user = st.session_state.get('user') or {}

    if current_user and current_user.get('username'):
        current_user = refrescar_perfil_usuario_en_sesion(current_user)
        st.session_state['user'] = current_user

    user_nombre = current_user.get('nombre', 'Líder')
    user_rol = current_user.get('rol', 'lider')
    user_grupo = str(current_user.get('codigo_grupo', '')).strip().split('.')[0] if current_user.get('codigo_grupo') else ""
    user_sector = str(current_user.get('codigo_sector', '')).strip() if current_user.get('codigo_sector') else ""
    if not user_sector and user_grupo:
        try:
            usuarios_cat = cargar_usuarios()
            for u_k, u_v in usuarios_cat.items():
                if str(u_v.get('codigo_grupo', '')).strip() == str(user_grupo).strip() and u_v.get('codigo_sector'):
                    user_sector = str(u_v.get('codigo_sector')).strip()
                    break
        except Exception:
            pass

    user_sector_nombre = obtener_nombre_sector_usuario(current_user)
    nombre_sector_app = obtener_nombre_corto_sector(user_sector_nombre)

    # CSS Ultra-Compacto y Responsivo para Smartphones y Tablets (Tema Natura & Avon)
    st.markdown("""
    <style>
        header[data-testid="stHeader"] {
            background: transparent !important;
        }
        .block-container {
            padding-top: clamp(2.8rem, 4.5vw, 3.5rem) !important;
            padding-bottom: 2rem !important;
            padding-left: clamp(0.5rem, 2vw, 1rem) !important;
            padding-right: clamp(0.5rem, 2vw, 1rem) !important;
            max-width: 100% !important;
        }
        /* Tarjeta de Encabezado Móvil Premium */
        .mob-header-card {
            background: linear-gradient(135deg, rgba(255, 107, 0, 0.08) 0%, rgba(227, 0, 123, 0.08) 100%);
            border: 1px solid rgba(227, 0, 123, 0.22);
            border-radius: 14px;
            padding: 10px 14px;
            margin-bottom: 12px;
            backdrop-filter: blur(12px);
            box-shadow: 0 4px 15px -2px rgba(227, 0, 123, 0.08);
        }
        .mob-header-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 4px;
            gap: 8px;
        }
        .mob-header-title {
            display: flex;
            align-items: center;
            gap: 7px;
        }
        .mob-title-icon {
            font-size: 1.25rem;
            line-height: 1;
        }
        .mob-title-text {
            font-size: 1.25rem;
            font-weight: 800;
            letter-spacing: -0.02em;
            background: linear-gradient(135deg, #FF6B00 0%, #E3007B 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            line-height: 1.2;
        }
        .mob-badge-group {
            background: linear-gradient(135deg, rgba(255, 107, 0, 0.15) 0%, rgba(227, 0, 123, 0.15) 100%);
            color: #FF6B00;
            border: 1px solid rgba(255, 107, 0, 0.35);
            font-size: 11px;
            font-weight: 700;
            padding: 2.5px 9px;
            border-radius: 20px;
            white-space: nowrap;
        }
        .mob-header-sub {
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 11.5px;
            gap: 6px;
            flex-wrap: wrap;
        }
        .mob-user-text {
            opacity: 0.9;
            letter-spacing: -0.01em;
            font-size: 12px;
        }
        .mob-badge-mode {
            font-size: 10px;
            font-weight: 700;
            color: #E3007B;
            background: rgba(227, 0, 123, 0.1);
            padding: 2px 8px;
            border-radius: 10px;
            border: 1px solid rgba(227, 0, 123, 0.22);
            white-space: nowrap;
        }
        /* Pestañas Ultra-Estilizadas con Soporte Dark/Light */
        .stTabs [data-baseweb="tab-list"] {
            gap: 5px !important;
            background-color: rgba(227, 0, 123, 0.07) !important;
            border-radius: 14px !important;
            padding: 4px !important;
            margin-bottom: 12px !important;
            border: 1px solid rgba(227, 0, 123, 0.18) !important;
        }
        .stTabs [data-baseweb="tab"] {
            font-size: 11.5px !important;
            font-weight: 700 !important;
            padding: 8px 11px !important;
            border-radius: 10px !important;
            color: inherit !important;
            opacity: 0.75 !important;
            transition: all 0.2s ease !important;
            border: none !important;
        }
        .stTabs [data-baseweb="tab"]:hover {
            opacity: 1 !important;
            background-color: rgba(227, 0, 123, 0.12) !important;
        }
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #FF6B00 0%, #E3007B 100%) !important;
            color: #ffffff !important;
            opacity: 1 !important;
            font-weight: 800 !important;
            box-shadow: 0 4px 14px rgba(227, 0, 123, 0.35) !important;
        }
        .stTabs [data-baseweb="tab-highlight"],
        .stTabs [data-baseweb="tab-border"] {
            display: none !important;
        }
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
        .kpi-sub-green { background: #dcfce7; color: #15803d; }
        .kpi-sub-red { background: #fee2e2; color: #b91c1c; }
        .kpi-sub-blue { background: #e0f2fe; color: #0369a1; }
        .kpi-sub-orange { background: #ffedd5; color: #c2410c; }
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
        .badge-pill {
            display: inline-block;
            font-size: 10px;
            font-weight: 700;
            padding: 2px 6px;
            border-radius: 9999px;
        }
        /* Desplegadores (Selectbox & Multiselect) Compactos */
        div[data-testid="stSelectbox"],
        div[data-testid="stMultiSelect"] {
            margin-bottom: 2px !important;
        }
        div[data-testid="stSelectbox"] label,
        div[data-testid="stMultiSelect"] label,
        div[data-testid="stTextInput"] label {
            min-height: 0 !important;
            margin-bottom: 2px !important;
            padding: 0 !important;
        }
        div[data-testid="stSelectbox"] label p,
        div[data-testid="stMultiSelect"] label p,
        div[data-testid="stTextInput"] label p {
            font-size: 0.74rem !important;
            font-weight: 700 !important;
            line-height: 1.15 !important;
            white-space: nowrap !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
            margin: 0 !important;
        }
        div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
        div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div {
            min-height: 31px !important;
            height: 31px !important;
            padding-top: 1px !important;
            padding-bottom: 1px !important;
            padding-left: 8px !important;
            padding-right: 4px !important;
            border-radius: 8px !important;
            font-size: 0.80rem !important;
            line-height: 1.2 !important;
        }
        div[data-testid="stTextInput"] input {
            min-height: 31px !important;
            height: 31px !important;
            padding: 2px 8px !important;
            border-radius: 8px !important;
            font-size: 0.80rem !important;
        }
        div[data-testid="stSelectbox"] div[data-baseweb="select"] [data-testid="stSelectboxVirtualDropdown"],
        div[data-testid="stSelectbox"] div[data-baseweb="select"] div {
            font-size: 0.80rem !important;
            line-height: 1.2 !important;
        }
        div[data-baseweb="select"] button,
        div[data-testid="stSelectbox"] button,
        div[data-testid="stMultiSelect"] button {
            background: transparent !important;
            background-color: transparent !important;
            border: none !important;
            box-shadow: none !important;
            padding: 0 2px !important;
            margin: 0 !important;
            min-width: 0 !important;
            width: auto !important;
            height: auto !important;
            border-radius: 0 !important;
            transform: none !important;
        }
        div[data-baseweb="select"] button:hover,
        div[data-testid="stSelectbox"] button:hover,
        div[data-testid="stMultiSelect"] button:hover {
            background: transparent !important;
            background-color: transparent !important;
            border: none !important;
            box-shadow: none !important;
            transform: none !important;
        }
        div[data-baseweb="select"] svg,
        div[data-testid="stSelectbox"] svg {
            width: 13px !important;
            height: 13px !important;
        }
        div[data-baseweb="tag"] {
            height: 20px !important;
            font-size: 10.5px !important;
            padding: 0 4px !important;
            margin: 1px 2px !important;
        }
        ul[data-baseweb="menu"] li {
            font-size: 0.80rem !important;
            padding: 5px 8px !important;
        }

        /* ========================================================================= */
        /* RUEDA ONDULADA DINÁMICA DE CARGA (CUSTOM WAVY LOADER SPINNER MÓVIL)       */
        /* ========================================================================= */
        div[data-testid="stSpinner"],
        div.stSpinner {
            display: flex !important;
            justify-content: center !important;
            align-items: center !important;
            margin: 1rem auto !important;
            width: 100% !important;
        }

        div[data-testid="stSpinner"] > div,
        div.stSpinner > div {
            display: inline-flex !important;
            flex-direction: row !important;
            align-items: center !important;
            justify-content: center !important;
            gap: 14px !important;
            background: #FFFFFF !important;
            padding: 12px 20px !important;
            border-radius: 16px !important;
            box-shadow: 0 10px 25px -5px rgba(11, 87, 208, 0.15) !important;
            border: 1.5px solid rgba(96, 165, 250, 0.35) !important;
            color: #0F172A !important;
            font-weight: 600 !important;
            font-size: 0.88rem !important;
        }

        div[data-testid="stSpinner"] svg,
        div.stSpinner svg,
        div[data-testid="stSpinner"] i,
        div.stSpinner i {
            display: none !important;
        }

        div[data-testid="stSpinner"] > div::before,
        div.stSpinner > div::before {
            content: "" !important;
            display: inline-block !important;
            width: 38px !important;
            height: 38px !important;
            min-width: 38px !important;
            min-height: 38px !important;
            background-image: url("data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMjAgMTIwIiBmaWxsPSJub25lIj4KICA8ZGVmcz4KICAgIDxsaW5lYXJHcmFkaWVudCBpZD0id2F2eUdyYWQiIHgxPSIwJSIgeTE9IjAlIiB4Mj0iMTAwJSIgeTI9IjEwMCUiPgogICAgICA8c3RvcCBvZmZzZXQ9IjAlIiBzdG9wLWNvbG9yPSIjMGI1N2QwIiBzdG9wLW9wYWNpdHk9IjEiIC8+CiAgICAgIDxzdG9wIG9mZnNldD0iNDUlIiBzdG9wLWNvbG9yPSIjMWE3M2U4IiBzdG9wLW9wYWNpdHk9IjEiIC8+CiAgICAgIDxzdG9wIG9mZnNldD0iNzAlIiBzdG9wLWNvbG9yPSIjNjBhNWZhIiBzdG9wLW9wYWNpdHk9IjAuNSIgLz4KICAgICAgPHN0b3Agb2Zmc2V0PSIxMDAlIiBzdG9wLWNvbG9yPSIjOTNjNWZkIiBzdG9wLW9wYWNpdHk9IjAuMTUiIC8+CiAgICA8L2xpbmVhckdyYWRpZW50PgogIDwvZGVmcz4KICA8cGF0aCBkPSJNIDk2LjAwIDYwLjAwIEwgOTguMjAgNjEuNzIgTCA5OS40NSA2My41NSBMIDk5LjI2IDY1LjMyIEwgOTcuNjIgNjYuODMgTCA5NS4xMCA2OC4wMSBMIDkyLjU1IDY4Ljk4IEwgOTAuODAgNzAuMDEgTCA5MC4zMiA3MS4zOCBMIDkxLjA1IDczLjI3IEwgOTIuNDMgNzUuNjIgTCA5My42NyA3OC4xMiBMIDk0LjAxIDgwLjMyIEwgOTMuMDYgODEuODIgTCA5MC45MyA4Mi40NyBMIDg4LjE1IDgyLjQ1IEwgODUuNDMgODIuMjIgTCA4My40MSA4Mi4zOCBMIDgyLjM4IDgzLjQxIEwgODIuMjIgODUuNDMgTCA4Mi40NSA4OC4xNSBMIDgyLjQ3IDkwLjkzIEwgODEuODIgOTMuMDYgTCA4MC4zMiA5NC4wMSBMIDc4LjEyIDkzLjY3IEwgNzUuNjIgOTIuNDMgTCA3My4yNyA5MS4wNSBMIDcxLjM4IDkwLjMyIEwgNzAuMDEgOTAuODAgTCA2OC45OCA5Mi41NSBMIDY4LjAxIDk1LjEwIEwgNjYuODMgOTcuNjIgTCA2NS4zMiA5OS4yNiBMIDYzLjU1IDk5LjQ1IEwgNjEuNzIgOTguMjAgTCA2MC4wMCA5Ni4wMCBMIDU4LjQ5IDkzLjczIEwgNTcuMTAgOTIuMjYgTCA1NS42NSA5Mi4wOSBMIDUzLjk3IDkzLjIyIEwgNTEuOTkgOTUuMTAgTCA0OS44MyA5Ni44NiBMIDQ3Ljc2IDk3LjY4IEwgNDYuMDggOTcuMDkgTCA0NC45NyA5NS4xNiBMIDQ0LjM4IDkyLjQzIEwgNDQuMDAgODkuNzMgTCA0My4zOSA4Ny44MCBMIDQyLjE2IDg3LjAzIEwgNDAuMTUgODcuMzIgTCAzNy41NSA4OC4xNSBMIDM0Ljg0IDg4Ljc5IEwgMzIuNjIgODguNjMgTCAzMS4zNyA4Ny4zOCBMIDMxLjIxIDg1LjE2IEwgMzEuODUgODIuNDUgTCAzMi42OCA3OS44NSBMIDMyLjk3IDc3Ljg0IEwgMzIuMjAgNzYuNjEgTCAzMC4yNyA3Ni4wMCBMIDI3LjU3IDc1LjYyIEwgMjQuODQgNzUuMDMgTCAyMi45MSA3My45MiBMIDIyLjMyIDcyLjI0IEwgMjMuMTQgNzAuMTcgTCAyNC45MCA2OC4wMSBMIDI2Ljc4IDY2LjAzIEwgMjcuOTEgNjQuMzUgTCAyNy43NCA2Mi45MCBMIDI2LjI3IDYxLjUxIEwgMjQuMDAgNjAuMDAgTCAyMS44MCA1OC4yOCBMIDIwLjU1IDU2LjQ1IEwgMjAuNzQgNTQuNjggTCAyMi4zOCA1My4xNyBMIDI0LjkwIDUxLjk5IEwgMjcuNDUgNTEuMDIgTCAyOS4yMCA0OS45OSBMIDI5LjY4IDQ4LjYyIEwgMjguOTUgNDYuNzMgTCAyNy41NyA0NC4zOCBMIDI2LjMzIDQxLjg4IEwgMjUuOTkgMzkuNjggTCAyNi45NCAzOC4xOCBMIDI5LjA3IDM3LjUzIEwgMzEuODUgMzcuNTUgTCAzNC41NyAzNy43OCBMIDM2LjU5IDM3LjYyIEwgMzcuNjIgMzYuNTkgTCAzNy43OCAzNC41NyBMIDM3LjU1IDMxLjg1IEwgMzcuNTMgMjkuMDcgTCAzOC4xOCAyNi45NCBMIDM5LjY4IDI1Ljk5IEwgNDEuODggMjYuMzMgTCA0NC4zOCAyNy41NyBMIDQ2LjczIDI4Ljk1IEwgNDguNjIgMjkuNjggTCA0OS45OSAyOS4yMCBMIDUxLjAyIDI3LjQ1IEwgNTEuOTkgMjQuOTAgTCA1My4xNyAyMi4zOCBMIDU0LjY4IDIwLjc0IEwgNTYuNDUgMjAuNTUgTCA1OC4yOCAyMS44MCBMIDYwLjAwIDI0LjAwIEwgNjEuNTEgMjYuMjcgTCA2Mi45MCAyNy43NCBMIDY0LjM1IDI3LjkxIEwgNjYuMDMgMjYuNzggTCA2OC4wMSAyNC45MCBMIDcwLjE3IDIzLjE0IEwgNzIuMjQgMjIuMzIgTCA3My45MiAyMi45MSBMIDc1LjAzIDI0Ljg0IEwgNzUuNjIgMjcuNTcgTCA3Ni4wMCAzMC4yNyBMIDc2LjYxIDMyLjIwIEwgNzcuODQgMzIuOTcgTCA3OS44NSAzMi42OCBMIDgyLjQ1IDMxLjg1IEwgODUuMTYgMzEuMjEgTCA4Ny4zOCAzMS4zNyBMIDg4LjYzIDMyLjYyIEwgODguNzkgMzQuODQgTCA4OC4xNSAzNy41NSBMIDg3LjMyIDQwLjE1IEwgODcuMDMgNDIuMTYgTCA4Ny44MCA0My4zOSBMIDg5LjczIDQ0LjAwIEwgOTIuNDMgNDQuMzggTCA5NS4xNiA0NC45NyBMIDk3LjA5IDQ2LjA4IEwgOTcuNjggNDcuNzYgTCA5Ni44NiA0OS44MyBMIDk1LjEwIDUxLjk5IEwgOTMuMjIgNTMuOTcgTCA5Mi4wOSA1NS42NSBMIDkyLjI2IDU3LjEwIEwgOTMuNzMgNTguNDkgWiIgc3Ryb2tlPSJ1cmwoI3dhdnlHcmFkKSIgc3Ryb2tlLXdpZHRoPSI1LjUiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgZmlsbD0ibm9uZSIgLz4KPC9zdmc+") !important;
            background-size: contain !important;
            background-repeat: no-repeat !important;
            background-position: center !important;
            animation: custom-wavy-spin 1.1s linear infinite !important;
        }

        @keyframes custom-wavy-spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
    </style>
    """, unsafe_allow_html=True)

    # Header Móvil Premium
    grupo_str = f"Grupo {user_grupo}" if user_grupo else "General"

    if mostrar_salir:
        st.info("💡 **Portal Unificado**: Ahora puedes ingresar directamente a la dirección oficial: [https://metaseindicadores.up.railway.app](https://metaseindicadores.up.railway.app)")
        col_h1, col_h2 = st.columns([3, 1])
        with col_h1:
            st.markdown(f"""
            <div class="mob-header-card">
                <div class="mob-header-top">
                    <div class="mob-header-title">
                        <span class="mob-title-icon">📱</span>
                        <span class="mob-title-text">App {nombre_sector_app}</span>
                    </div>
                    <span class="mob-badge-group">🏷️ {grupo_str}</span>
                </div>
                <div class="mob-header-sub">
                    <span class="mob-user-text">👤 <b>{user_nombre}</b></span>
                    <span class="mob-badge-mode">✨ Vista Móvil</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_h2:
            st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)
            if user_rol in ['gerente', 'superadmin']:
                if st.button("💳 Pagos", key="btn_pagos_mob_standalone", use_container_width=True, help="Ver cuota, QRs de pago y referidas"):
                    modal_pago_suscripcion_gerente(current_user, user_sector, user_nombre)
            if st.button("🚪 Salir", key="btn_logout_mob_standalone", use_container_width=True):
                st.session_state['user'] = None
                st.query_params.clear()
                st.rerun()
    else:
        col_m_top1, col_m_top2 = st.columns([3, 1.2])
        with col_m_top1:
            st.markdown(f"""
            <div class="mob-header-card">
                <div class="mob-header-top">
                    <div class="mob-header-title">
                        <span class="mob-title-icon">📱</span>
                        <span class="mob-title-text">App {nombre_sector_app}</span>
                    </div>
                    <span class="mob-badge-group">🏷️ {grupo_str}</span>
                </div>
                <div class="mob-header-sub">
                    <span class="mob-user-text">👤 <b>{user_nombre}</b></span>
                    <span class="mob-badge-mode">✨ Vista Móvil</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_m_top2:
            st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)
            if st.button("💻 PC", key="btn_switch_to_desktop", use_container_width=True, help="Cambiar a la versión de Escritorio (PC)"):
                st.session_state['sb_segmented_vista'] = "💻 Escritorio"
                st.query_params['vista'] = 'escritorio'
                st.rerun()
            if user_rol in ['gerente', 'superadmin']:
                if st.button("💳 Pagos", key="btn_pagos_mob_top", use_container_width=True, help="Ver cuota, QRs de pago y referidas"):
                    modal_pago_suscripcion_gerente(current_user, user_sector, user_nombre)

    # Aviso no invasivo de suscripción en móvil (solo si faltan <= 3 días o está vencido)
    if user_rol in ['gerente', 'superadmin']:
        banner_alerta_suscripcion_gerente(current_user, user_sector, user_nombre, es_movil=True)

    # Si es Gerente o Admin, permitir seleccionar qué grupo auditar
    grupo_activo = user_grupo
    if user_rol in ['gerente', 'superadmin']:
        usuarios_todos = cargar_usuarios()
        grupos_disponibles = sorted(list(set(
            str(u.get('codigo_grupo')).strip().split('.')[0]
            for u in usuarios_todos.values()
            if u.get('codigo_grupo') and (user_rol == 'superadmin' or str(u.get('codigo_sector', '')).strip() == str(user_sector).strip())
        )))
        if grupos_disponibles:
            col_g_sel1, col_g_sel2 = st.columns([2, 1])
            with col_g_sel1:
                grupo_activo = st.selectbox("👩‍💼 Ver Grupo:", options=grupos_disponibles, index=0, key="mob_sel_grp_gerente")
            with col_g_sel2:
                st.caption("*(Vista Gerencial)*")

    # 5. Carga de Datos Relacionales (Tableau, Geral, Objetivos Arte y Cómo Vamos)
    df_tab = consultar_tableau_sql(
        grupo=grupo_activo if grupo_activo else None,
        sector=user_sector if (not grupo_activo and user_sector and user_rol != 'superadmin') else None
    )
    df_geral = consultar_geral_sql(
        grupo=grupo_activo if grupo_activo else None,
        sector=user_sector if (not grupo_activo and user_sector and user_rol != 'superadmin') else None
    )

    # Carga de Metas Efectivas (Objetivos Arte oficial + Ajustes Desafíos de Zona)
    mapa_arte = cargar_objetivos_arte()
    arte_lider = obtener_metas_efectivas(grupo=grupo_activo) if grupo_activo else {}

    # Carga de Cómo Vamos
    df_cv_all = None
    try:
        df_cv_all = calcular_metas_ciclo()
    except Exception:
        if os.path.exists('Base para el como vamos.xlsx'):
            try:
                df_cv_all = calcular_metas_ciclo('Base para el como vamos.xlsx')
            except Exception:
                pass

    # Aislamiento Multitenant Estricto por Sector en Móvil (tanto para Líder como para Gerente)
    if user_sector and user_rol != 'superadmin' and df_cv_all is not None and not df_cv_all.empty:
        col_sec_found = None
        for c in df_cv_all.columns:
            c_low = str(c).lower().replace('ó', 'o')
            if 'setor' in c_low or 'sector' in c_low:
                col_sec_found = c
                break
        if col_sec_found:
            s_vals = df_cv_all[col_sec_found].astype(str).str.strip().str.replace('.0', '', regex=False)
            df_cv_all = df_cv_all[s_vals == str(user_sector).strip()]
        else:
            grupos_sector = {str(u.get('codigo_grupo')).strip() for u in cargar_usuarios().values() if str(u.get('codigo_sector')).strip() == str(user_sector).strip() and u.get('codigo_grupo')}
            col_g_ref = next((c for c in df_cv_all.columns if 'grupo' in str(c).lower()), None)
            if col_g_ref and grupos_sector:
                g_vals = df_cv_all[col_g_ref].astype(str).str.split('.').str[0].str.strip()
                df_cv_all = df_cv_all[g_vals.isin(grupos_sector)]

    df_cv = pd.DataFrame()
    if df_cv_all is not None and not df_cv_all.empty and grupo_activo:
        col_g = next((c for c in df_cv_all.columns if 'grupo' in str(c).lower()), None)
        if col_g:
            df_cv = df_cv_all[df_cv_all[col_g].astype(str).str.split('.').str[0].str.strip() == str(grupo_activo).strip()]

    # --- MODALES INTERACTIVOS DE DRILL-DOWN (AUDITORÍA DE ORIGEN DE DATOS) ---
    if hasattr(st, 'dialog'):
        @st.dialog("🔍 Origen y Desglose de Activas Reales", width="large")
        def dialog_origen_activas(df_fuente, sector_nombre, real_act, obj_act, cump_act):
            st.markdown(f"### 📊 Auditoría y Desglose de Activas Reales")
            st.markdown(f"**Ámbito:** `{sector_nombre}` • **Total Activas Reportadas:** `{int(real_act)}`")
            
            st.info(
                "📁 **Origen oficial de los datos:** Archivo comercial **`Base para el como vamos.xlsx`** (Reporte oficial de metas de Natura).\n\n"
                "🔢 **Fórmula del cálculo:** Se toma la columna **`Real Activas`** y se suman las activas acumuladas de ciclo de cada una de las líderes de negocio que componen este sector."
            )

            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("👥 Activas Reales Logradas", f"{int(real_act)}")
            with c2:
                st.metric("🎯 Objetivo Total de Activas", f"{int(obj_act)}")
            with c3:
                st.metric("📈 Cumplimiento de Sector", f"{cump_act:.1f}%")

            st.markdown("---")
            st.markdown("#### 📋 Aporte Detallado por cada Líder de Negocio")

            col_grp = next((c for c in df_fuente.columns if 'grupo' in str(c).lower()), 'Código de grupo')
            col_nom = 'Nombre de consultora' if 'Nombre de consultora' in df_fuente.columns else ('Consultora' if 'Consultora' in df_fuente.columns else df_fuente.columns[0])

            df_det = df_fuente.copy()
            if col_nom in df_det.columns:
                df_det = df_det[df_det[col_nom].notna() & (~df_det[col_nom].astype(str).str.strip().str.lower().isin(['nan', 'none', '', '0']))]

            cols_sel = []
            rename_map = {}
            if col_grp in df_det.columns:
                cols_sel.append(col_grp)
                rename_map[col_grp] = 'Grupo'
            if col_nom in df_det.columns:
                cols_sel.append(col_nom)
                rename_map[col_nom] = 'Líder / Grupo'
            if 'Real Activas' in df_det.columns:
                cols_sel.append('Real Activas')
                rename_map['Real Activas'] = 'Activas Reales'
            if 'Objetivo Activas' in df_det.columns:
                cols_sel.append('Objetivo Activas')
                rename_map['Objetivo Activas'] = 'Meta Activas'
            if 'Cumplimiento Activas' in df_det.columns:
                cols_sel.append('Cumplimiento Activas')
                rename_map['Cumplimiento Activas'] = '% Cumplimiento'

            df_view = df_det[cols_sel].copy() if cols_sel else df_det.copy()
            if rename_map:
                df_view = df_view.rename(columns=rename_map)

            if 'Activas Reales' in df_view.columns:
                df_view['Activas Reales'] = pd.to_numeric(df_view['Activas Reales'], errors='coerce').fillna(0).astype(int)
                df_view = df_view.sort_values(by='Activas Reales', ascending=False)

            if 'Meta Activas' in df_view.columns:
                df_view['Meta Activas'] = pd.to_numeric(df_view['Meta Activas'], errors='coerce').fillna(0).astype(int)

            if '% Cumplimiento' in df_view.columns:
                df_view['% Cumplimiento'] = pd.to_numeric(df_view['% Cumplimiento'], errors='coerce').fillna(0).apply(lambda v: f"{v:.1f}%")

            st.dataframe(df_view, use_container_width=True, hide_index=True)

            st.caption(
                "💡 **Diferencia técnica con Tableau:** En la pestaña *'📋 MI LISTADO'*, se auditan las consultoras individuales con pedido en el corte específico de base. "
                "En cambio, este indicador de metas suma las activas oficiales que acumulan las líderes según las reglas del ciclo comercial de Natura."
            )

        @st.dialog("💰 Origen y Desglose de Facturación", width="large")
        def dialog_origen_facturacion(df_fuente, sector_nombre, r_fact, o_fact, c_fact):
            st.markdown(f"### 💰 Auditoría de Facturación — {sector_nombre}")
            st.info("📁 **Origen oficial:** Archivo comercial **`Base para el como vamos.xlsx`** (columna `Real Facturación`).")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Venta Real", f"${r_fact:,.0f} COP")
            with c2:
                st.metric("Meta Venta", f"${o_fact:,.0f} COP")
            with c3:
                st.metric("Cumplimiento", f"{c_fact:.1f}%")

            col_grp = next((c for c in df_fuente.columns if 'grupo' in str(c).lower()), 'Código de grupo')
            col_nom = 'Nombre de consultora' if 'Nombre de consultora' in df_fuente.columns else df_fuente.columns[0]
            df_f = df_fuente.copy()
            if col_nom in df_f.columns:
                df_f = df_f[df_f[col_nom].notna() & (~df_f[col_nom].astype(str).str.strip().str.lower().isin(['nan', 'none', '', '0']))]
            cols_f = [c for c in [col_grp, col_nom, 'Real Facturación', 'Objetivo Facturación', 'Cumplimiento Facturación', 'Ganancia estimada'] if c in df_f.columns]
            df_f_view = df_f[cols_f].copy()
            if 'Real Facturación' in df_f_view.columns:
                df_f_view['Real Facturación'] = df_f_view['Real Facturación'].apply(formato_cop)
            if 'Objetivo Facturación' in df_f_view.columns:
                df_f_view['Objetivo Facturación'] = df_f_view['Objetivo Facturación'].apply(formato_cop)
            if 'Cumplimiento Facturación' in df_f_view.columns:
                df_f_view['Cumplimiento Facturación'] = df_f_view['Cumplimiento Facturación'].apply(lambda v: f"{v:.1f}%")
            if 'Ganancia estimada' in df_f_view.columns:
                df_f_view['Ganancia estimada'] = df_f_view['Ganancia estimada'].apply(formato_cop)
            st.dataframe(df_f_view, use_container_width=True, hide_index=True)
    else:
        def dialog_origen_activas(*args, **kwargs): pass
        def dialog_origen_facturacion(*args, **kwargs): pass

    # 6. Pestañas Principales Móviles (4 Pestañas Condensadas en MAYÚSCULAS)
    tab_cv, tab_tab, tab_cartera, tab_lideres = st.tabs([
        "🎯 MIS DESAFÍOS",
        "📋 MI LISTADO",
        "💳 MI CARTERA",
        "👑 MIS LÍDERES"
    ])

    # ==============================================================================
    # TAB 1: MIS DESAFÍOS & CÓMO VAMOS (OBJETIVOS ARTE + CÓMO VAMOS)
    # ==============================================================================
    with tab_cv:
        st.markdown("##### 🎯 Cuadro de Mando de Desafíos Líder")
        if arte_lider.get('es_ajuste_zona'):
            st.caption(f"✨ *Metas calibradas por Gerencia de Zona ({arte_lider.get('campana', 'Campaña Activa')})*")

        # Banner de Cumpleaños Móvil (si hay cumpleañeras en el equipo)
        try:
            from procesador import obtener_cumpleanos_equipo
            plantilla_wa = st.session_state.get('plantilla_wa_cumpleanos', None)
            data_cumple = obtener_cumpleanos_equipo(df_tab, user_nombre=user_nombre, plantilla_wa=plantilla_wa)
            if data_cumple and data_cumple.get('total_mes', 0) > 0:
                h_c = len(data_cumple.get('hoy', []))
                s_c = len(data_cumple.get('semana', []))
                tot_c = data_cumple.get('total_mes', 0)
                mes_c = data_cumple.get('nombre_mes', '')
                
                if h_c > 0:
                    b_txt = f"🎂 ¡HOY HAY {h_c} CUMPLEAÑERA{'S' if h_c > 1 else ''}! 🎉"
                elif s_c > 0:
                    b_txt = f"🎁 {s_c} CUMPLEAÑOS EN PRÓXIMOS 7 DÍAS 📅"
                else:
                    b_txt = f"🗓️ {tot_c} CUMPLEAÑOS EN {mes_c.upper()} ✨"
                
                with st.expander(f"{b_txt} • Ver & Felicitar", expanded=False):
                    items_cumple = data_cumple.get('hoy', []) + data_cumple.get('semana', [])
                    if not items_cumple:
                        items_cumple = data_cumple.get('mes', [])[:12]
                    for itm in items_cumple[:12]:
                        c_c1, c_c2 = st.columns([2, 1])
                        with c_c1:
                            st.markdown(f"🌸 **{itm['nombre']}** (CB: `{itm['codigo_cb']}` • Día {itm['dia']})")
                        with c_c2:
                            if itm.get('link_wa'):
                                st.link_button("📲 Felicitar", url=itm['link_wa'], use_container_width=True)
        except Exception:
            pass

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

            # Botones de Auditoría Drill-Down Móvil
            col_aud1, col_aud2 = st.columns(2)
            with col_aud1:
                if st.button("🔍 Auditar Activas Reales", key="mob_drill_act", use_container_width=True, help="Auditar de dónde salen las activas y el aporte de cada líder"):
                    dialog_origen_activas(df_cv_all if (df_cv_all is not None and not df_cv_all.empty) else df_cv, user_sector if user_sector else (f"Grupo {grupo_activo}" if grupo_activo else "Sector"), act_r, act_m, act_pct)
            with col_aud2:
                if st.button("💰 Auditar Facturación", key="mob_drill_fact", use_container_width=True, help="Auditar el origen de la facturación y metas"):
                    dialog_origen_facturacion(df_cv_all if (df_cv_all is not None and not df_cv_all.empty) else df_cv, user_sector if user_sector else (f"Grupo {grupo_activo}" if grupo_activo else "Sector"), fact_r, fact_m, fact_pct)

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
        st.caption("Escribe las notas de gestión por cada consultora. Se guardarán de forma permanente por `Codigo CB`. Puedes usar el corrector del explorador (subrayado rojo y clic derecho) para sugerencias ortográficas directas.")

        # Filtros Rápidos
        f_c1, f_c2, f_c3, f_c4 = st.columns(4)
        with f_c1:
            opciones_sit = ["Todas"]
            if not df_tab.empty and 'Sit. Comercial' in df_tab.columns:
                sits_unicas = [str(x) for x in df_tab['Sit. Comercial'].dropna().unique() if str(x).strip()]
                opciones_sit += sorted(sits_unicas)
            filtro_sit = st.selectbox("🏷️ Sit. Comercial", options=opciones_sit, key="mob_f_sit")

        with f_c2:
            opciones_col = ["Todos"]
            col_nivel_tab = 'Nivel / Color' if 'Nivel / Color' in df_tab.columns else ('Color' if 'Color' in df_tab.columns else None)
            if not df_tab.empty and col_nivel_tab:
                cols_unicas = [str(x) for x in df_tab[col_nivel_tab].dropna().unique() if str(x).strip()]
                opciones_col += sorted(cols_unicas)
            filtro_col = st.selectbox("🏆 Nivel / Color", options=opciones_col, key="mob_f_col")

        with f_c3:
            filtro_mora = st.selectbox("💳 Deuda en Mora", options=["Todas", "🔴 Solo con Mora", "🟢 Al Día"], key="mob_f_mora")

        with f_c4:
            filtro_ped = st.selectbox("⌛ Pedidos Pendientes", options=["Todos", "Con Pedidos Pendientes (> 0)", "Sin Pedidos Pendientes (0)"], key="mob_f_ped")

        # Búsqueda rápida por nombre, documento o código
        busq_nom = st.text_input("🔍 Buscar consultora, documento o código CB...", placeholder="Escribe nombre, cédula o código...", key="mob_b_nom").strip()

        # Aplicar filtros
        df_tab_filtrado = df_tab.copy() if not df_tab.empty else pd.DataFrame()

        if not df_tab_filtrado.empty:
            if filtro_sit != "Todas" and 'Sit. Comercial' in df_tab_filtrado.columns:
                df_tab_filtrado = df_tab_filtrado[df_tab_filtrado['Sit. Comercial'].astype(str) == filtro_sit]

            if filtro_col != "Todos" and col_nivel_tab:
                df_tab_filtrado = df_tab_filtrado[df_tab_filtrado[col_nivel_tab].astype(str) == filtro_col]

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
                col_nom_filtro = 'Consultora' if 'Consultora' in df_tab_filtrado.columns else ('Asesora / Consultora' if 'Asesora / Consultora' in df_tab_filtrado.columns else None)
                if col_nom_filtro:
                    mask_busq = mask_busq | df_tab_filtrado[col_nom_filtro].astype(str).str.contains(busq_nom, case=False, na=False)
                if 'DocumentoGPP' in df_tab_filtrado.columns:
                    mask_busq = mask_busq | df_tab_filtrado['DocumentoGPP'].astype(str).str.contains(busq_nom, case=False, na=False)
                if 'Código CB' in df_tab_filtrado.columns:
                    mask_busq = mask_busq | df_tab_filtrado['Código CB'].astype(str).str.contains(busq_nom, case=False, na=False)
                df_tab_filtrado = df_tab_filtrado[mask_busq]

        # 1. Las 3 ventanitas KPI (Total cadastro, Total disponibles, Deuda Mora Total)
        tot_cadastro = len(df_tab_filtrado)
        
        col_sit_check = 'Sit. Comercial' if 'Sit. Comercial' in df_tab_filtrado.columns else ('Situación' if 'Situación' in df_tab_filtrado.columns else None)
        if col_sit_check and not df_tab_filtrado.empty:
            s_vals_lower = df_tab_filtrado[col_sit_check].astype(str).str.strip().str.lower()
            mask_disp = s_vals_lower.apply(
                lambda s: any(k in s for k in ['activa', 'activas', 'inactiva 1', 'inactiva 2', 'inactiva 3', 'i1', 'i2', 'i3']) and not any(k in s for k in ['inactiva 4', 'inactiva 5', 'inactiva 6', 'i4', 'i5', 'i6'])
            )
            tot_disponibles = int(mask_disp.sum())
        else:
            tot_disponibles = 0

        tot_mora = df_tab_filtrado['Deuda Mora'].apply(lambda x: limpiar_numero(x, 0.0)).sum() if ('Deuda Mora' in df_tab_filtrado.columns and not df_tab_filtrado.empty) else 0.0
        if tot_mora >= 1_000_000:
            mora_str = f"${tot_mora/1e6:.2f}M COP"
        else:
            mora_str = f"${tot_mora:,.0f} COP".replace(",", ".")

        st.markdown(f"""
        <div class="kpi-grid-3" style="margin: 8px 0 14px 0;">
            <div class="kpi-card" style="background: linear-gradient(135deg, rgba(255, 107, 0, 0.06) 0%, rgba(227, 0, 123, 0.06) 100%); border: 1.5px solid rgba(227, 0, 123, 0.28); border-radius: 16px; padding: 12px 6px; box-shadow: 0 4px 15px rgba(227, 0, 123, 0.08); text-align: center;">
                <div class="kpi-title" style="font-size: 11px; font-weight: 700; color: #334155; margin-bottom: 4px; display: flex; align-items: center; justify-content: center; gap: 4px;">👥 Total cadastro</div>
                <div class="kpi-val" style="font-size: 22px; font-weight: 800; background: linear-gradient(135deg, #FF6B00 0%, #E3007B 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; line-height: 1.2;">{tot_cadastro}</div>
            </div>
            <div class="kpi-card" style="background: linear-gradient(135deg, rgba(255, 107, 0, 0.06) 0%, rgba(227, 0, 123, 0.06) 100%); border: 1.5px solid rgba(227, 0, 123, 0.28); border-radius: 16px; padding: 12px 6px; box-shadow: 0 4px 15px rgba(227, 0, 123, 0.08); text-align: center;">
                <div class="kpi-title" style="font-size: 11px; font-weight: 700; color: #334155; margin-bottom: 4px; display: flex; align-items: center; justify-content: center; gap: 4px;">🎯 Total disponibles</div>
                <div class="kpi-val" style="font-size: 22px; font-weight: 800; background: linear-gradient(135deg, #FF6B00 0%, #E3007B 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; line-height: 1.2;">{tot_disponibles}</div>
            </div>
            <div class="kpi-card" style="background: linear-gradient(135deg, rgba(255, 107, 0, 0.06) 0%, rgba(227, 0, 123, 0.06) 100%); border: 1.5px solid rgba(227, 0, 123, 0.28); border-radius: 16px; padding: 12px 6px; box-shadow: 0 4px 15px rgba(227, 0, 123, 0.08); text-align: center;">
                <div class="kpi-title" style="font-size: 11px; font-weight: 700; color: #334155; margin-bottom: 4px; display: flex; align-items: center; justify-content: center; gap: 4px;">⚠️ Deuda Mora Total</div>
                <div class="kpi-val" style="font-size: 22px; font-weight: 800; background: linear-gradient(135deg, #FF6B00 0%, #E3007B 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; line-height: 1.2;">{mora_str}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        # Tarjetas Responsivas de Niveles (Bronce, Plata, Oro, Zafiro, Diamante) con Paleta Oficial
        if not df_tab_filtrado.empty:
            col_color_t = 'Color' if 'Color' in df_tab_filtrado.columns else ('Nivel / Color' if 'Nivel / Color' in df_tab_filtrado.columns else None)
            if col_color_t:
                tiers_mob = [
                    {"nombre": "Bronce", "color": "#CD7F32", "icon": "🥉", "bg": "rgba(205, 127, 50, 0.12)", "border": "rgba(205, 127, 50, 0.45)"},
                    {"nombre": "Plata", "color": "#64748B", "icon": "🥈", "bg": "rgba(100, 116, 139, 0.14)", "border": "rgba(100, 116, 139, 0.45)"},
                    {"nombre": "Oro", "color": "#D97706", "icon": "🥇", "bg": "rgba(245, 158, 11, 0.14)", "border": "rgba(217, 119, 6, 0.45)"},
                    {"nombre": "Zafiro", "color": "#2563EB", "icon": "💎", "bg": "rgba(37, 99, 235, 0.12)", "border": "rgba(37, 99, 235, 0.45)"},
                    {"nombre": "Diamante", "color": "#7C3AED", "icon": "👑", "bg": "rgba(124, 58, 237, 0.13)", "border": "rgba(124, 58, 237, 0.45)"}
                ]
                
                tot_mob_all = len(df_tab_filtrado)
                col_sit_mob = 'Sit. Comercial' if 'Sit. Comercial' in df_tab_filtrado.columns else ('Situación' if 'Situación' in df_tab_filtrado.columns else None)
                
                cards_mob_html = '<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(92px, 1fr)); gap: 6px; margin: 4px 0 12px 0;">'
                for tm in tiers_mob:
                    sub_tm = df_tab_filtrado[df_tab_filtrado[col_color_t].astype(str).str.strip().str.lower() == tm["nombre"].lower()]
                    c_tot = len(sub_tm)
                    c_act = len(sub_tm[sub_tm[col_sit_mob].astype(str).str.strip().str.lower() == 'activa']) if col_sit_mob else 0
                    pct_t = (c_tot / tot_mob_all * 100) if tot_mob_all > 0 else 0
                    
                    cards_mob_html += f'''<div style="background: {tm['bg']}; border: 1.5px solid {tm['border']}; border-top: 3.5px solid {tm['color']}; border-radius: 12px; padding: 7px 5px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.06);">
                        <div style="font-size: 10px; font-weight: 800; color: {tm['color']}; text-transform: uppercase;">{tm['icon']} {tm['nombre']}</div>
                        <div style="font-size: 17px; font-weight: 800; color: {tm['color']}; line-height: 1.2; margin: 2px 0;">{c_tot}</div>
                        <div style="font-size: 9px; color: #64748B; font-weight: 600;">{pct_t:.0f}% total</div>
                        <div style="font-size: 9.5px; color: #15803D; font-weight: 700; margin-top: 2px;">🟢 {c_act} act.</div>
                    </div>'''
                cards_mob_html += '</div>'
                st.markdown(cards_mob_html, unsafe_allow_html=True)

                # --- DESGLOSE ANALÍTICO POR NIVEL Y ESTADO COMERCIAL (MÓVIL) ---
        with st.expander("🎨 Ver Análisis de Niveles, Estados y Facturación", expanded=False):
            st.markdown("##### 🎨 Clasificación por Niveles y Estado Comercial")
            orden_niveles = ['Bronce', 'Plata', 'Oro', 'Zafiro', 'Diamante']
            orden_sit = ['Activa', 'Inactiva 1', 'Inactiva 2', 'Inactiva 3', 'Inactiva 4', 'Inactiva 5', 'Inactiva 6', 'Cesada', 'Posible Baja', 'Registrada']

            df_calc_tab_m = df_tab_filtrado.copy() if not df_tab_filtrado.empty else pd.DataFrame()
            if not df_calc_tab_m.empty:
                if 'Fact. Total' in df_calc_tab_m.columns:
                    df_calc_tab_m['__fact_val__'] = pd.to_numeric(df_calc_tab_m['Fact. Total'], errors='coerce').fillna(0.0)
                    if df_calc_tab_m['__fact_val__'].sum() == 0 and 'Fact. Natura' in df_calc_tab_m.columns:
                        df_calc_tab_m['__fact_val__'] = (
                            pd.to_numeric(df_calc_tab_m['Fact. Natura'], errors='coerce').fillna(0.0) +
                            pd.to_numeric(df_calc_tab_m.get('Fact. AVON', 0), errors='coerce').fillna(0.0) +
                            pd.to_numeric(df_calc_tab_m.get('Fact. C&E', 0), errors='coerce').fillna(0.0) +
                            pd.to_numeric(df_calc_tab_m.get('Fact. VOL', 0), errors='coerce').fillna(0.0)
                        )
                elif 'Fact. Natura' in df_calc_tab_m.columns:
                    df_calc_tab_m['__fact_val__'] = (
                        pd.to_numeric(df_calc_tab_m['Fact. Natura'], errors='coerce').fillna(0.0) +
                        pd.to_numeric(df_calc_tab_m.get('Fact. AVON', 0), errors='coerce').fillna(0.0) +
                        pd.to_numeric(df_calc_tab_m.get('Fact. C&E', 0), errors='coerce').fillna(0.0) +
                        pd.to_numeric(df_calc_tab_m.get('Fact. VOL', 0), errors='coerce').fillna(0.0)
                    )
                else:
                    df_calc_tab_m['__fact_val__'] = 0.0

                mask_fact_pos_m = df_calc_tab_m['__fact_val__'] > 0
                if mask_fact_pos_m.any():
                    if df_calc_tab_m.loc[mask_fact_pos_m, '__fact_val__'].quantile(0.9) < 10000:
                        df_calc_tab_m['__fact_val__'] = df_calc_tab_m['__fact_val__'] * 1000.0

                col_sit_m = 'Sit. Comercial' if 'Sit. Comercial' in df_calc_tab_m.columns else ('Situación' if 'Situación' in df_calc_tab_m.columns else None)
                df_calc_tab_m['__es_activa__'] = df_calc_tab_m[col_sit_m].astype(str).str.strip().str.lower() == 'activa' if col_sit_m else False

                col_color_m = 'Color' if 'Color' in df_calc_tab_m.columns else ('Nivel / Color' if 'Nivel / Color' in df_calc_tab_m.columns else None)
                if col_color_m:
                    df_calc_tab_m[col_color_m] = df_calc_tab_m[col_color_m].astype(str).str.strip()
                if col_sit_m:
                    df_calc_tab_m[col_sit_m] = df_calc_tab_m[col_sit_m].astype(str).str.strip()

                # Tabla 1: Niveles
                st.markdown("###### 🏆 Distribución por Nivel")
                if col_color_m:
                    df_calc_valid_c_m = df_calc_tab_m[~df_calc_tab_m[col_color_m].str.lower().isin(['nan', 'none', ''])]
                    df_color_group_m = df_calc_valid_c_m.groupby(col_color_m).agg(
                        Cantidad=(col_color_m, 'count'),
                        Activas=('__es_activa__', 'sum'),
                        Facturacion_Total=('__fact_val__', 'sum')
                    ).reset_index()

                    df_color_group_m['__orden__'] = df_color_group_m[col_color_m].apply(
                        lambda c: orden_niveles.index(c) if c in orden_niveles else 99
                    )
                    df_color_group_m = df_color_group_m.sort_values(by='__orden__').drop(columns=['__orden__'])

                    df_color_group_m['% Actividad'] = (df_color_group_m['Activas'] / df_color_group_m['Cantidad'] * 100).round(1)
                    df_color_group_m['Ticket Promedio'] = df_color_group_m.apply(
                        lambda r: r['Facturacion_Total'] / r['Activas'] if r['Activas'] > 0 else 0.0, axis=1
                    )

                    df_color_render_m = df_color_group_m.copy()
                    df_color_render_m['% Actividad'] = df_color_render_m['% Actividad'].apply(lambda x: f"{x:.1f}%")
                    df_color_render_m['Facturación Total'] = df_color_render_m['Facturacion_Total'].apply(formato_cop)
                    df_color_render_m['Ticket Promedio'] = df_color_render_m['Ticket Promedio'].apply(formato_cop)
                    df_color_render_m = df_color_render_m.rename(columns={
                        col_color_m: 'Nivel / Color',
                        'Cantidad': 'Total Red',
                        'Activas': 'Activas',
                        'Ticket Promedio': 'Productividad'
                    })

                    cols_c_render_m = ['Nivel / Color', 'Total Red', 'Activas', '% Actividad', 'Facturación Total', 'Productividad']
                    df_color_render_m = df_color_render_m[[c for c in cols_c_render_m if c in df_color_render_m.columns]].reset_index(drop=True)

                    st.dataframe(
                        df_color_render_m.style.map(color_nivel, subset=['Nivel / Color'] if 'Nivel / Color' in df_color_render_m.columns else []),
                        use_container_width=True,
                        hide_index=True
                    )

                # Tabla 2: Matriz Cruzada
                if col_color_m and col_sit_m:
                    st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
                    st.markdown("###### 🔍 Matriz Cruzada: Nivel vs. Situación Comercial")
                    df_calc_mat_m = df_calc_tab_m[
                        (~df_calc_tab_m[col_color_m].str.lower().isin(['nan', 'none', ''])) &
                        (~df_calc_tab_m[col_sit_m].str.lower().isin(['nan', 'none', '']))
                    ]
                    if not df_calc_mat_m.empty:
                        df_pivot_m = pd.crosstab(df_calc_mat_m[col_color_m], df_calc_mat_m[col_sit_m])
                        df_pivot_m.columns.name = None
                        cols_sit_en_mat_m = [s for s in orden_sit if s in df_pivot_m.columns] + [c for c in df_pivot_m.columns if c not in orden_sit]
                        df_pivot_m = df_pivot_m[cols_sit_en_mat_m].reset_index()
                        df_pivot_m['Total'] = df_pivot_m[cols_sit_en_mat_m].sum(axis=1)

                        if 'Activa' in df_pivot_m.columns:
                            df_pivot_m['% Actividad'] = (df_pivot_m['Activa'] / df_pivot_m['Total'] * 100).round(1).apply(lambda x: f"{x:.1f}%")
                        else:
                            df_pivot_m['% Actividad'] = "0.0%"

                        fact_x_color_m = df_calc_mat_m.groupby(col_color_m)['__fact_val__'].sum().to_dict()
                        df_pivot_m['Facturación Activas'] = df_pivot_m[col_color_m].map(fact_x_color_m).fillna(0.0).apply(formato_cop)

                        df_pivot_m['__orden__'] = df_pivot_m[col_color_m].apply(
                            lambda c: orden_niveles.index(c) if c in orden_niveles else 99
                        )
                        df_pivot_m = df_pivot_m.sort_values(by='__orden__').drop(columns=['__orden__'])

                        cols_mat_final_m = [col_color_m, 'Total']
                        if 'Activa' in df_pivot_m.columns:
                            cols_mat_final_m.append('Activa')
                        for s_col in cols_sit_en_mat_m:
                            if s_col != 'Activa' and s_col in df_pivot_m.columns:
                                cols_mat_final_m.append(s_col)
                        cols_mat_final_m.extend(['Facturación Activas', '% Actividad'])

                        df_mat_render_m = df_pivot_m[[c for c in cols_mat_final_m if c in df_pivot_m.columns]].copy()
                        df_mat_render_m = df_mat_render_m.rename(columns={col_color_m: 'Nivel / Color', 'Total': 'Total Red'})

                        mat_styler_m = df_mat_render_m.style.map(
                            color_nivel, subset=['Nivel / Color'] if 'Nivel / Color' in df_mat_render_m.columns else []
                        )
                        st.dataframe(mat_styler_m, use_container_width=True, hide_index=True)

        if df_tab_filtrado.empty:
            st.info("ℹ️ No hay consultoras con los filtros seleccionados.")
        else:
            # Limpiar, ordenar y estandarizar columnas para que coincidan exactamente con la base canónica (16 columnas)
            df_edit_view = limpiar_y_ordenar_columnas_tableau(df_tab_filtrado, {}, es_lider=True)

            # Limpiar cualquier flotante residual en todo el DataFrame para eliminar decimales (.000000)
            for c in df_edit_view.columns:
                if c not in ['DocumentoGPP', 'Celular', 'Código CB', 'Codigo CB'] and pd.api.types.is_float_dtype(df_edit_view[c]):
                    df_edit_view[c] = df_edit_view[c].fillna(0).round().astype('int64')

            # Detectar columnas de puntos y situación histórica del ciclo anterior
            cols_pts_hist_m = [c for c in df_edit_view.columns if any(k in str(c).lower() for k in ['(c-', 'ant', 'cierre']) and 'pts' in str(c).lower()]
            cols_sit_hist_m = [c for c in df_edit_view.columns if any(k in str(c).lower() for k in ['(c-', 'ant', 'cierre']) and ('sit' in str(c).lower() or 'comercial' in str(c).lower())]
            if cols_pts_hist_m or cols_sit_hist_m:
                ref_col_m = cols_pts_hist_m[0] if cols_pts_hist_m else cols_sit_hist_m[0]
                tag_cierre_m = ref_col_m.split('(')[-1].replace(')', '').strip() if '(' in ref_col_m else "Cierre"
                resaltar_cierre_mob = st.toggle(f"🎨 Resaltar Puntos Ciclo Anterior ({tag_cierre_m})", value=True, key="mob_toggle_resaltar_cierre")
            else:
                resaltar_cierre_mob = False

            # Usar st.data_editor para permitir editar notas directamente en la tabla
            col_config = {}
            for col_name in df_edit_view.columns:
                # Si es una columna de dinero (Deuda o Facturación), formatear con $
                if 'Deuda' in col_name or 'Fact.' in col_name:
                    col_config[col_name] = st.column_config.NumberColumn(col_name, format="$%d", disabled=True)
                # Código CB (sin inmovilizar en móvil para permitir desplazamiento horizontal cómodo)
                elif col_name in ['Código CB', 'Codigo CB']:
                    col_config[col_name] = st.column_config.TextColumn(str(col_name), disabled=True)
                # Asesora / Consultora (sin inmovilizar en móvil)
                elif col_name in ['Asesora / Consultora', 'Consultora', 'Nombre']:
                    col_config[col_name] = st.column_config.TextColumn(str(col_name), disabled=True)
                # Si es DocumentoGPP o Celular, formatear como texto limpio sin comas
                elif col_name in ['DocumentoGPP', 'Celular']:
                    col_config[col_name] = st.column_config.TextColumn(str(col_name), disabled=True)
                # Si es una columna numérica (Pts, Crédito, Pedidos, Ciclos), formatear como número entero limpio sin $
                elif 'Pts' in col_name or 'Ped.' in col_name or 'Ciclos' in col_name or 'Credito' in col_name or 'Crédito' in col_name:
                    col_config[col_name] = st.column_config.NumberColumn(col_name, format="%d", disabled=True)
                else:
                    col_config[col_name] = st.column_config.TextColumn(str(col_name), disabled=True)

            if "Notas / Comentarios Líder" in df_edit_view.columns:
                col_config["Notas / Comentarios Líder"] = st.column_config.TextColumn("Notas / Comentarios Líder", disabled=False)

            total_filas_edit = len(df_edit_view)
            if total_filas_edit > 300:
                df_data_render = df_edit_view.iloc[:300]
                st.caption(f"⚡ *Mostrando las primeras 300 de {total_filas_edit:,} consultoras para máxima velocidad. Usa el buscador o filtros para localizar a una consultora específica.*")
            else:
                df_data_render = df_edit_view

            num_celdas = len(df_data_render) * len(df_data_render.columns)
            if num_celdas <= 250_000:
                try:
                    subsets_sit_m = [c for c in df_data_render.columns if ('sit' in str(c).lower() or 'situación' in str(c).lower()) and 'inactividad' not in str(c).lower()]
                    styler_mob = df_data_render.style.map(
                        color_nivel, subset=['Nivel / Color'] if 'Nivel / Color' in df_data_render.columns else []
                    ).map(
                        color_situacion, subset=subsets_sit_m
                    ).map(
                        color_deuda_mora, subset=['Deuda Mora'] if 'Deuda Mora' in df_data_render.columns else []
                    )
                    if resaltar_cierre_mob and cols_pts_hist_m:
                        styler_mob = styler_mob.map(
                            color_pts_cierre_anterior, subset=cols_pts_hist_m
                        )
                    df_data_to_edit = styler_mob
                except Exception:
                    df_data_to_edit = df_data_render
            else:
                df_data_to_edit = df_data_render

            edited_df = st.data_editor(
                df_data_to_edit,
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
                            if row_idx < len(df_data_render):
                                codigo_key = limpiar_codigo_cb_estandar(df_data_render.iloc[row_idx].get('Código CB', ''))
                                nueva_nota = str(row_changes["Notas / Comentarios Líder"]).strip()
                                if codigo_key:
                                    dict_autoguardar[codigo_key] = nueva_nota
                        except Exception:
                            pass

                if dict_autoguardar:
                    guardar_todos_comentarios(dict_autoguardar)
                    registrar_evento_auditoria(
                        current_user,
                        categoria="💬 Gestión Comercial",
                        accion="Guardado de Notas Móvil",
                        detalle=f"Autoguardado de {len(dict_autoguardar)} notas en móvil",
                        dispositivo="📱 Móvil / Tablet"
                    )
                    st.toast(f"💾 Guardado: {len(dict_autoguardar)} nota(s) actualizada(s)", icon="✅")

            # Botón de guardado manual
            col_s1, col_s2 = st.columns([1.5, 2.5])
            with col_s1:
                if st.button("💾 Guardar Manualmente", type="primary", use_container_width=True, key="mob_save_manual_btn"):
                    dict_guardar = {}
                    for idx, row in edited_df.iterrows():
                        codigo_key = limpiar_codigo_cb_estandar(row.get('Código CB', ''))
                        nota_val = str(row.get('Notas / Comentarios Líder', '')).strip()
                        if codigo_key:
                            dict_guardar[codigo_key] = nota_val

                    if guardar_todos_comentarios(dict_guardar):
                        registrar_evento_auditoria(
                            current_user,
                            categoria="💬 Gestión Comercial",
                            accion="Guardado de Notas Móvil",
                            detalle=f"Guardado manual de {len(dict_guardar)} notas en móvil",
                            dispositivo="📱 Móvil / Tablet"
                        )
                        st.success("✅ ¡Todas las notas han sido guardadas exitosamente!")
                        st.rerun()
            with col_s2:
                st.caption("🟢 **Guardado automático activo**: Al escribir una nota y pulsar `Enter`, se guarda de forma instantánea.")

            # --- SECCIÓN DE MENSAJERÍA WHATSAPP DIRECTA DEL LISTADO FILTRADO (MÓVIL) ---
            st.markdown("---")
            with st.expander(f"📲 Contacto & WhatsApp del Listado Filtrado ({len(df_edit_view)} Consultoras)", expanded=False):
                st.markdown("##### 📲 Envíos y Campaña de WhatsApp sobre el Listado Filtrado")
                st.caption("Contacta a las consultoras que acabas de filtrar en tu tabla. Puedes enviarles tus notas personalizadas, recordatorios o promociones en 1 clic.")

                # Identificar columnas canónicas
                c_col_cb = 'Código CB' if 'Código CB' in df_edit_view.columns else ('Codigo CB' if 'Codigo CB' in df_edit_view.columns else None)
                c_col_nom = 'Consultora' if 'Consultora' in df_edit_view.columns else ('Asesora / Consultora' if 'Asesora / Consultora' in df_edit_view.columns else ('Nombre' if 'Nombre' in df_edit_view.columns else None))
                c_col_cel = 'Celular' if 'Celular' in df_edit_view.columns else ('celular' if 'celular' in df_edit_view.columns else None)
                c_col_sit = 'Sit. Comercial' if 'Sit. Comercial' in df_edit_view.columns else None
                c_col_col = 'Nivel / Color' if 'Nivel / Color' in df_edit_view.columns else None
                c_col_nota = 'Notas / Comentarios Líder' if 'Notas / Comentarios Líder' in df_edit_view.columns else None
                c_col_ped = 'Ped. Pendientes' if 'Ped. Pendientes' in df_edit_view.columns else None
                c_col_mora = 'Deuda Mora' if 'Deuda Mora' in df_edit_view.columns else None
                c_col_pts = 'Pts Acum' if 'Pts Acum' in df_edit_view.columns else None

                if c_col_cb and c_col_nom:
                    # Mapeo de consultoras disponibles
                    mapa_wa_mob = {}
                    for _, r_w in df_edit_view.iterrows():
                        k_cb = str(r_w.get(c_col_cb, '')).strip()
                        n_consultora = str(r_w.get(c_col_nom, '')).strip()
                        n_color = str(r_w.get(c_col_col, 'Nivel')) if c_col_col else ''
                        n_sit = str(r_w.get(c_col_sit, 'Estado')) if c_col_sit else ''
                        n_nota = str(r_w.get(c_col_nota, '')).strip() if c_col_nota else ''
                        
                        etiqueta = f"[{n_color}] [{n_sit}] {n_consultora} (CB: {k_cb})"
                        if n_nota and n_nota.lower() not in ['nan', 'none']:
                            etiqueta += f' — 💬 "{n_nota[:24]}..."'
                        mapa_wa_mob[k_cb] = etiqueta

                    # Subgrupos para selección rápida
                    cbs_todas_m = list(mapa_wa_mob.keys())
                    cbs_con_notas = [str(r.get(c_col_cb, '')).strip() for _, r in df_edit_view.iterrows() if str(r.get(c_col_nota, '')).strip() and str(r.get(c_col_nota, '')).strip().lower() not in ['nan', 'none']] if c_col_nota else []
                    cbs_inactivas = [str(r.get(c_col_cb, '')).strip() for _, r in df_edit_view.iterrows() if 'inactiva' in str(r.get(c_col_sit, '')).lower()] if c_col_sit else []
                    cbs_con_ped = [str(r.get(c_col_cb, '')).strip() for _, r in df_edit_view.iterrows() if float(limpiar_numero(r.get(c_col_ped, 0))) > 0] if c_col_ped else []
                    cbs_con_mora = [str(r.get(c_col_cb, '')).strip() for _, r in df_edit_view.iterrows() if float(limpiar_numero(r.get(c_col_mora, 0))) > 0] if c_col_mora else []

                    # Inicializar estado de selección de casillas
                    if 'cbs_sel_mob_wa' not in st.session_state or st.session_state.get('cbs_sel_mob_wa') is None:
                        st.session_state['cbs_sel_mob_wa'] = set(cbs_todas_m)
                    elif isinstance(st.session_state['cbs_sel_mob_wa'], list):
                        st.session_state['cbs_sel_mob_wa'] = set(st.session_state['cbs_sel_mob_wa'])
                    if 'editor_ver_mob' not in st.session_state:
                        st.session_state['editor_ver_mob'] = 0

                    st.markdown("<p style='font-size: 0.84rem; font-weight: 700; color: #334155; margin: 4px 0 2px 0;'>🎯 Segmentación Rápida con Casillas (O marca/desmarca directamente en la lista abajo):</p>", unsafe_allow_html=True)
                    num_bcols = 5 + (1 if cbs_con_notas else 0)
                    b_cols_m = st.columns(num_bcols)
                    b_idx = 0
                    with b_cols_m[b_idx]:
                        if st.button(f"👥 Todas ({len(df_edit_view)})", key="btn_sel_todas_mob_wa", use_container_width=True):
                            st.session_state['cbs_sel_mob_wa'] = set(cbs_todas_m)
                            st.session_state['editor_ver_mob'] = st.session_state.get('editor_ver_mob', 0) + 1
                            st.rerun()
                    b_idx += 1
                    with b_cols_m[b_idx]:
                        if st.button(f"🌸 Inactivas ({len(cbs_inactivas)})", key="btn_sel_inact_mob_wa", use_container_width=True):
                            st.session_state['cbs_sel_mob_wa'] = set(cbs_inactivas)
                            st.session_state['editor_ver_mob'] = st.session_state.get('editor_ver_mob', 0) + 1
                            st.rerun()
                    b_idx += 1
                    with b_cols_m[b_idx]:
                        if st.button(f"⌛ Con Pedidos ({len(cbs_con_ped)})", key="btn_sel_ped_mob_wa", use_container_width=True):
                            st.session_state['cbs_sel_mob_wa'] = set(cbs_con_ped)
                            st.session_state['editor_ver_mob'] = st.session_state.get('editor_ver_mob', 0) + 1
                            st.rerun()
                    b_idx += 1
                    with b_cols_m[b_idx]:
                        if st.button(f"🚨 Con Mora ({len(cbs_con_mora)})", key="btn_sel_mora_mob_wa", use_container_width=True):
                            st.session_state['cbs_sel_mob_wa'] = set(cbs_con_mora)
                            st.session_state['editor_ver_mob'] = st.session_state.get('editor_ver_mob', 0) + 1
                            st.rerun()
                    b_idx += 1
                    if cbs_con_notas:
                        with b_cols_m[b_idx]:
                            if st.button(f"💬 Con Notas ({len(cbs_con_notas)})", key="btn_sel_notas_mob_wa", use_container_width=True):
                                st.session_state['cbs_sel_mob_wa'] = set(cbs_con_notas)
                                st.session_state['editor_ver_mob'] = st.session_state.get('editor_ver_mob', 0) + 1
                                st.rerun()
                        b_idx += 1
                    with b_cols_m[b_idx]:
                        if st.button("🧹 Ninguna (0)", key="btn_desel_todas_mob_wa", use_container_width=True):
                            st.session_state['cbs_sel_mob_wa'] = set()
                            st.session_state['editor_ver_mob'] = st.session_state.get('editor_ver_mob', 0) + 1
                            st.rerun()

                    # Flyer de Campaña Opcional (Formato ultra compacto con 1-clic copiado)
                    flyer_mob_file = None
                    with st.expander("🖼️ Adjuntar Imagen o Flyer de Campaña (Opcional)", expanded=False):
                        flyer_mob_file = st.file_uploader(
                            "Subir imagen promocional o volante para WhatsApp",
                            type=["jpg", "jpeg", "png", "webp"],
                            key="flyer_mob_wa_uploader",
                            help="Sube un volante o flyer. Podrás copiarlo al portapapeles con 1 clic para pegarlo directamente en WhatsApp (Ctrl + V)."
                        )
                        if flyer_mob_file is not None:
                            b64_flyer_m = base64.b64encode(flyer_mob_file.getvalue()).decode('utf-8')
                            mime_flyer_m = flyer_mob_file.type or 'image/jpeg'
                            flyer_kb_m = len(flyer_mob_file.getvalue()) // 1024

                            try:
                                dir_campanas = os.path.join('data', 'campanas')
                                os.makedirs(dir_campanas, exist_ok=True)
                                with open(os.path.join(dir_campanas, flyer_mob_file.name), 'wb') as f_fly:
                                    f_fly.write(flyer_mob_file.getvalue())
                            except Exception:
                                pass

                            st.markdown(f"""
                            <div style="background: rgba(37, 211, 102, 0.08); border: 1px solid rgba(37, 211, 102, 0.35); border-radius: 10px; padding: 8px 12px; margin-bottom: 6px;">
                                <div style="font-weight: 700; color: #128C7E; font-size: 0.85rem; display: flex; align-items: center; justify-content: space-between;">
                                    <span>🖼️ Flyer Listo: <strong>{flyer_mob_file.name}</strong> ({flyer_kb_m} KB)</span>
                                    <span style="background: #25D366; color: white; padding: 2px 7px; border-radius: 8px; font-size: 0.68rem; font-weight: 700;">LISTO</span>
                                </div>
                                <div style="font-size: 0.75rem; color: #475569; margin-top: 2px;">
                                    Pulsa <strong>Copiar Flyer</strong> y luego en WhatsApp presiona <strong>Ctrl + V</strong> para enviar imagen y texto juntos.
                                </div>
                            </div>
                            """, unsafe_allow_html=True)

                            col_btn_m1, col_btn_m2 = st.columns([1.5, 1])
                            with col_btn_m1:
                                copy_script_mob = f"""
                                <div style="display: flex; flex-direction: column; gap: 4px; font-family: sans-serif;">
                                    <button onclick="copiarFlyerMob()" style="
                                        background: linear-gradient(135deg, #25D366 0%, #128C7E 100%);
                                        color: white; border: none; padding: 7px 10px; border-radius: 8px;
                                        font-weight: 700; font-size: 0.80rem; cursor: pointer; width: 100%;
                                        display: flex; align-items: center; justify-content: center; gap: 6px;
                                        box-shadow: 0 2px 6px rgba(37, 211, 102, 0.25);
                                    ">
                                        📋 Copiar Flyer al Portapapeles (Ctrl + V)
                                    </button>
                                    <div id="statusCopiarFlyerMob" style="font-size: 0.72rem; color: #128C7E; font-weight: 600; text-align: center; display: none;"></div>
                                </div>
                                <script>
                                async function copiarFlyerMob() {{
                                    try {{
                                        const b64 = "{b64_flyer_m}";
                                        const byteCharacters = atob(b64);
                                        const byteNumbers = new Array(byteCharacters.length);
                                        for (let i = 0; i < byteCharacters.length; i++) {{
                                            byteNumbers[i] = byteCharacters.charCodeAt(i);
                                        }}
                                        const byteArray = new Uint8Array(byteNumbers);
                                        const blob = new Blob([byteArray], {{ type: "{mime_flyer_m}" }});
                                        await navigator.clipboard.write([
                                            new ClipboardItem({{ [blob.type]: blob }})
                                        ]);
                                        const el = document.getElementById('statusCopiarFlyerMob');
                                        el.innerText = '✅ ¡Flyer copiado! Abre WhatsApp y presiona Ctrl + V';
                                        el.style.display = 'block';
                                    }} catch (e) {{
                                        const el = document.getElementById('statusCopiarFlyerMob');
                                        el.innerHTML = '💡 <em>Para copiar: Clic derecho sobre la imagen ➔ "Copiar imagen"</em>';
                                        el.style.display = 'block';
                                    }}
                                }}
                                </script>
                                """
                                st.components.v1.html(copy_script_mob, height=48)
                            with col_btn_m2:
                                st.download_button(
                                    "📥 Descargar Flyer",
                                    data=flyer_mob_file.getvalue(),
                                    file_name=flyer_mob_file.name,
                                    mime=mime_flyer_m,
                                    use_container_width=True,
                                    key="btn_dl_flyer_mob_wa"
                                )
                            col_prev1, col_prev2 = st.columns([1, 2])
                            with col_prev1:
                                st.image(flyer_mob_file, caption="Vista previa", width=140)

                    tipo_camp_mob = st.selectbox(
                        "Tipo de Plantilla de Mensaje:",
                        options=[
                            "💬 1. Usar mis Notas / Comentarios",
                            "🎁 2. Reactivación Comercial (Inactivas)",
                            "🌟 3. Impulso de Puntos & Nivel",
                            "📦 4. Pedido Pendiente / Retenido",
                            "🚨 5. Cobro Amable / Recordatorio de Pago",
                            "⏰ 6. Cierre de Campaña Urgente",
                            "🌸 7. Saludo & Acompañamiento General",
                            "✍️ 8. Mensaje Libre / Personalizado"
                        ],
                        index=0 if cbs_con_notas else 1,
                        key="sel_tipo_camp_mob_widget"
                    )
                    remitente_mob_wa = user_nombre if user_nombre else "Tu Líder"

                    # Plantilla predeterminada según tipo
                    if "1. Usar mis Notas" in tipo_camp_mob:
                        tpl_mob_def = (
                            "Hola *{primer_nombre}* 🌸, te saluda tu Líder {remitente} de *Natura & Avon*.\n\n"
                            "Te contacto para contarte: *{nota}*.\n\n"
                            "¡Quedo muy atenta a lo que necesites para apoyarte! ✨"
                        )
                    elif "2. Reactivación" in tipo_camp_mob:
                        tpl_mob_def = (
                            "¡Hola *{primer_nombre}*! 🌸 Te extrañamos mucho en nuestro equipo de *Natura & Avon*.\n\n"
                            "En este ciclo tenemos promociones exclusivas y descuentos pensados para ti.\n\n"
                            "¿Te gustaría que te comparta el catálogo virtual interactivo de este ciclo? 📖✨"
                        )
                    elif "3. Impulso" in tipo_camp_mob:
                        tpl_mob_def = (
                            "¡Hola *{primer_nombre}*! 🌟 Felicitaciones por tus *{pts_acum} puntos* acumulados en tu nivel *{nivel}*.\n\n"
                            "Estás muy cerca de tu siguiente meta de premios de este ciclo. ¡Pasa tu pedido y gana más con Natura & Avon! 🎁✨"
                        )
                    elif "4. Pedido" in tipo_camp_mob:
                        tpl_mob_def = (
                            "Hola *{primer_nombre}* 🛍️, te saluda tu Líder {remitente} de *Natura & Avon*.\n\n"
                            "Tienes *{pedidos} pedido(s)* en espera de despacho por saldo de *{deuda_mora}*.\n\n"
                            "Al poner al día tu pago hoy, tu pedido saldrá de inmediato para entrega. ¡Quedo atenta para ayudarte! 📦✨"
                        )
                    elif "5. Cobro Amable" in tipo_camp_mob:
                        tpl_mob_def = (
                            "Hola *{primer_nombre}* 🌸, te saluda tu Líder {remitente} de *Natura & Avon*.\n\n"
                            "Te recuerdo amablemente que registras un saldo pendiente de *{deuda_mora}*.\n\n"
                            "Ponte al día hoy mismo para mantener tu crédito activo y no perder tus premios y beneficios de este ciclo. ¡Avísame si necesitas el enlace de pago PSE! 💳✨"
                        )
                    elif "6. Cierre" in tipo_camp_mob:
                        tpl_mob_def = (
                            "¡URGENTE *{primer_nombre}*! ⏰ Quedan pocas horas para el cierre de campaña de *Natura & Avon*.\n\n"
                            "No te quedes sin tus puntos ({pts_acum} acumulados) y aprovecha las ofertas de liquidación de este ciclo. ¡Pasa tu pedido ya y gana más! 🚀✨"
                        )
                    elif "7. Saludo" in tipo_camp_mob:
                        tpl_mob_def = (
                            "Hola *{primer_nombre}* 🌸, te saluda tu Líder {remitente} de *Natura & Avon*.\n\n"
                            "Quería saludarte y desearte muchos éxitos en tus ventas de este ciclo. ¡Cuenta conmigo para cualquier apoyo! ✨"
                        )
                    else:
                        tpl_mob_def = "Hola *{primer_nombre}* 🌸, te escribe tu Líder {remitente}.\n\n"

                    texto_plantilla_mob = st.text_area(
                        "✏️ Personaliza la Plantilla:",
                        value=tpl_mob_def,
                        height=100,
                        key=f"txt_tpl_mob_{tipo_camp_mob[:2]}"
                    )
                    st.caption("Variables: `{primer_nombre}`, `{nombre}`, `{nota}`, `{nivel}`, `{pts_acum}`, `{pedidos}`, `{deuda_mora}`, `{remitente}`")

                    # Generar filas de mensajes para todo el grupo filtrado con casillas
                    filas_wa_mob = []
                    set_sel_mob_actual = st.session_state.get('cbs_sel_mob_wa', set(cbs_todas_m))
                    for _, r_t in df_edit_view.iterrows():
                        cb_val_m = str(r_t.get(c_col_cb, '')).strip()
                        n_full = str(r_t.get(c_col_nom, '')).strip()
                        p_nom = n_full.split()[0].title() if n_full else "Consultora"
                        cel_raw = str(r_t.get(c_col_cel, '')).strip().replace(' ', '').replace('-', '').replace('+', '')
                        cel_val = cel_raw.split('.')[0] if '.' in cel_raw else cel_raw
                        
                        nota_val = str(r_t.get(c_col_nota, '')).strip() if c_col_nota else ''
                        if not nota_val or nota_val.lower() in ['nan', 'none']:
                            nota_val = ""
                        sit_val_m = str(r_t.get(c_col_sit, '')) if c_col_sit else ''
                        nivel_val = str(r_t.get(c_col_col, 'Consultora')) if c_col_col else 'Consultora'
                        pts_val = str(r_t.get(c_col_pts, '0')) if c_col_pts else '0'
                        ped_val = str(r_t.get(c_col_ped, '0')) if c_col_ped else '0'
                        mora_val = formato_cop(r_t.get(c_col_mora, 0)) if c_col_mora else '$0'

                        # Reemplazar variables
                        msg_m = (
                            texto_plantilla_mob
                            .replace("{primer_nombre}", p_nom)
                            .replace("{nombre}", n_full.title())
                            .replace("{nota}", nota_val if nota_val else "tenemos novedades especiales para ti")
                            .replace("{nivel}", nivel_val)
                            .replace("{pts_acum}", pts_val)
                            .replace("{pedidos}", ped_val)
                            .replace("{deuda_mora}", mora_val)
                            .replace("{remitente}", remitente_mob_wa)
                        )

                        link_m = f"https://api.whatsapp.com/send?phone=57{cel_val}&text={urllib.parse.quote(msg_m)}" if cel_val and len(cel_val) >= 10 else ""

                        fila_m = {
                            '✅ Enviar': cb_val_m in set_sel_mob_actual,
                            'Consultora': n_full,
                            'Código CB': cb_val_m,
                            'Sit. Comercial': sit_val_m,
                            'Celular': cel_val if cel_val else "Sin celular",
                        }
                        if flyer_mob_file is not None:
                            fila_m['📎 Flyer Listo'] = '🖼️ Adjunto'
                        fila_m.update({
                            'Nota Líder': nota_val if nota_val else "-",
                            'Enlace WhatsApp': link_m,
                            'Mensaje': msg_m
                        })
                        filas_wa_mob.append(fila_m)

                    df_campana_mob_out = pd.DataFrame(filas_wa_mob)

                    st.caption("👇 **Toca las casillas '✅ Enviar'** para incluir o quitar a cualquier consultora de tu lista, o pulsa **'Abrir WhatsApp'** para chatear directamente.")

                    cols_editor_mob = ['✅ Enviar', 'Consultora', 'Sit. Comercial', 'Celular']
                    if flyer_mob_file is not None:
                        cols_editor_mob.append('📎 Flyer Listo')
                    cols_editor_mob.extend(['Nota Líder', 'Enlace WhatsApp'])

                    cols_disabled_mob = [c for c in cols_editor_mob if c != '✅ Enviar']

                    col_cfg_mob = {
                        '✅ Enviar': st.column_config.CheckboxColumn(
                            "✅ Enviar",
                            help="Marca o desmarca la casilla para incluirla en el envío",
                            default=True
                        ),
                        "Enlace WhatsApp": st.column_config.LinkColumn(
                            "📲 Enviar WhatsApp",
                            display_text="Abrir WhatsApp"
                        )
                    }
                    if flyer_mob_file is not None:
                        col_cfg_mob['📎 Flyer Listo'] = st.column_config.TextColumn(
                            "📎 Flyer",
                            help="Indica que el flyer está listo para enviar junto con el mensaje",
                            width="small"
                        )

                    df_editado_mob = st.data_editor(
                        df_campana_mob_out[cols_editor_mob],
                        column_config=col_cfg_mob,
                        disabled=cols_disabled_mob,
                        use_container_width=True,
                        hide_index=True,
                        key=f"editor_campana_mob_{st.session_state.get('editor_ver_mob', 0)}"
                    )

                    # Resumen de marcadas
                    df_marcadas_m = df_editado_mob[df_editado_mob['✅ Enviar'] == True]
                    n_marc_m = len(df_marcadas_m)
                    n_act_m2 = len(df_marcadas_m[df_marcadas_m['Sit. Comercial'].astype(str).str.strip().str.lower() == 'activa']) if not df_marcadas_m.empty else 0
                    n_inact_m2 = len(df_marcadas_m[df_marcadas_m['Sit. Comercial'].astype(str).str.contains('inactiva', case=False, na=False)]) if not df_marcadas_m.empty else 0
                    n_mora_m = len(df_marcadas_m[df_marcadas_m['Consultora'].isin([r['Consultora'] for r in filas_wa_mob if float(limpiar_numero(r.get('deuda_mora', 0))) > 0])]) if not df_marcadas_m.empty else 0

                    st.markdown(f"""
                    <div style="background: #F8FAFC; border: 1.5px solid #CBD5E1; border-radius: 10px; padding: 6px 12px; margin-top: 6px; display: flex; align-items: center; justify-content: space-between;">
                        <div>
                            <span style="font-size: 0.85rem; font-weight: 800; color: #0F172A;">🎯 Marcadas con casilla: </span>
                            <span style="font-size: 0.95rem; font-weight: 800; color: #EA580C;">{n_marc_m}</span> 
                            <span style="font-size: 0.70rem; color: #64748B;">de {len(df_campana_mob_out)}</span>
                        </div>
                        <div style="display: flex; gap: 4px; font-size: 0.68rem; font-weight: 700;">
                            <span style="background: #F0FDF4; color: #166534; padding: 2px 5px; border-radius: 5px; border: 1px solid #BBF7D0;">🟢 {n_act_m2} Act.</span>
                            <span style="background: #FFF7ED; color: #C2410C; padding: 2px 5px; border-radius: 5px; border: 1px solid #FFEDD5;">🌸 {n_inact_m2} Inact.</span>
                            <span style="background: #FEF2F2; color: #DC2626; padding: 2px 5px; border-radius: 5px; border: 1px solid #FECACA;">🚨 {n_mora_m} Mora</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                    # Botones de exportación rápida para las consultoras marcadas
                    if not df_marcadas_m.empty:
                        idx_marcadas = df_editado_mob[df_editado_mob['✅ Enviar'] == True].index
                        df_marcadas_exp = df_campana_mob_out.loc[idx_marcadas]

                        st.markdown("<div style='margin-top: 6px;'></div>", unsafe_allow_html=True)
                        c_dl1, c_dl2 = st.columns(2)
                        with c_dl1:
                            buf_csv = io.BytesIO()
                            cols_exp = [c for c in ['Consultora', 'Código CB', 'Sit. Comercial', 'Celular', 'Nota Líder', 'Mensaje'] if c in df_marcadas_exp.columns]
                            df_marcadas_exp[cols_exp].to_csv(buf_csv, index=False, encoding='utf-8-sig')
                            st.download_button(
                                f"📥 Descargar CSV ({len(df_marcadas_exp)})",
                                data=buf_csv.getvalue(),
                                file_name=f"campana_wa_mob_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                                mime="text/csv",
                                use_container_width=True,
                                key="btn_dl_csv_mob_wa"
                            )
                        with c_dl2:
                            buf_xls = io.BytesIO()
                            try:
                                with pd.ExcelWriter(buf_xls, engine='openpyxl') as writer:
                                    df_marcadas_exp[cols_exp].to_excel(writer, index=False, sheet_name='Campana WA')
                                st.download_button(
                                    f"📊 Descargar Excel ({len(df_marcadas_exp)})",
                                    data=buf_xls.getvalue(),
                                    file_name=f"campana_wa_mob_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,
                                    key="btn_dl_xls_mob_wa"
                                )
                            except Exception:
                                pass
                else:
                    st.warning("⚠️ No se encontraron las columnas necesarias en la tabla.")

    # ==============================================================================
    # TAB 3: MI CARTERA & COBRANZA PREVENTIVA
    # ==============================================================================
    with tab_cartera:
        st.markdown("##### 💳 Mi Cartera & Cobranza Preventiva")
        st.caption("Control de facturas por vencer y cobranza preventiva por WhatsApp para tu grupo.")

        if df_geral is None or df_geral.empty:
            st.info("🎉 ¡Excelente noticia! No hay facturas de cobranza pendientes en tu grupo.")
        else:
            df_g_mob = df_geral.copy()
            # 1. Limpiar números y calcular días
            if 'saldo_total' in df_g_mob.columns:
                df_g_mob['saldo_num'] = df_g_mob['saldo_total'].apply(lambda x: float(limpiar_numero(x, 0.0)))
            else:
                df_g_mob['saldo_num'] = 0.0

            dias_col = 'dias_para_vencer' if 'dias_para_vencer' in df_g_mob.columns else None
            if dias_col:
                df_g_mob['dias_num'] = df_g_mob[dias_col].apply(lambda x: int(limpiar_numero(x, 0)))
            else:
                df_g_mob['dias_num'] = 0

            # KPIs Superiores Móviles
            total_facturas_g = len(df_g_mob)
            total_deuda_g = df_g_mob['saldo_num'].sum()
            mora_g = df_g_mob[df_g_mob['dias_num'] < 0]['saldo_num'].sum()
            manana_g = len(df_g_mob[df_g_mob['dias_num'] == 1])
            pasado_g = len(df_g_mob[df_g_mob['dias_num'] == 2])

            st.markdown(f"""
            <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 8px; margin-bottom: 12px;">
                <div style="background: white; border-radius: 10px; padding: 10px; border: 1px solid #E2E8F0; text-align: center;">
                    <div style="font-size: 11px; color: #64748B; font-weight: 700;">💳 FACTURAS PENDIENTES</div>
                    <div style="font-size: 18px; font-weight: 800; color: #0F172A;">{total_facturas_g} ({formato_cop(total_deuda_g)})</div>
                </div>
                <div style="background: white; border-radius: 10px; padding: 10px; border: 1px solid #FEE2E2; text-align: center;">
                    <div style="font-size: 11px; color: #DC2626; font-weight: 700;">🔴 EN MORA</div>
                    <div style="font-size: 18px; font-weight: 800; color: #DC2626;">{formato_cop(mora_g)}</div>
                </div>
                <div style="background: white; border-radius: 10px; padding: 10px; border: 1px solid #FEF3C7; text-align: center;">
                    <div style="font-size: 11px; color: #D97706; font-weight: 700;">🟡 VENCEN MAÑANA</div>
                    <div style="font-size: 18px; font-weight: 800; color: #D97706;">{manana_g} facturas</div>
                </div>
                <div style="background: white; border-radius: 10px; padding: 10px; border: 1px solid #DCFCE7; text-align: center;">
                    <div style="font-size: 11px; color: #16A34A; font-weight: 700;">🟢 PASADO MAÑANA</div>
                    <div style="font-size: 18px; font-weight: 800; color: #16A34A;">{pasado_g} facturas</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Filtro por vencimiento
            filtro_venc_mob = st.radio(
                "📅 **Filtrar por Vencimiento:**",
                options=[
                    "🗓️ Todas las Facturas",
                    "🟡 Vencen Mañana",
                    "🟢 Pasado Mañana",
                    "🔴 En Mora",
                    "📅 Próximos 7 Días"
                ],
                horizontal=True,
                key="radio_venc_geral_mob"
            )

            df_g_filt = df_g_mob.copy()
            if "Mañana" in filtro_venc_mob:
                df_g_filt = df_g_filt[df_g_filt['dias_num'] == 1]
            elif "Pasado Mañana" in filtro_venc_mob:
                df_g_filt = df_g_filt[df_g_filt['dias_num'] == 2]
            elif "En Mora" in filtro_venc_mob:
                df_g_filt = df_g_filt[df_g_filt['dias_num'] < 0]
            elif "Próximos 7" in filtro_venc_mob:
                df_g_filt = df_g_filt[(df_g_filt['dias_num'] >= 0) & (df_g_filt['dias_num'] <= 7)]

            if df_g_filt.empty:
                st.info("🎉 No hay facturas en esta categoría.")
            else:
                st.markdown(f"###### 📋 {len(df_g_filt)} Facturas Seleccionadas:")
                # Selector de cobro individual
                nombres_cartera = [str(r.get('nombre', '')).strip() for _, r in df_g_filt.iterrows()]
                sel_consultora_cobro = st.selectbox("Elige la consultora para enviar recordatorio de pago:", options=nombres_cartera, key="sel_asesora_cobro_mob")
                row_cobro = df_g_filt[df_g_filt['nombre'] == sel_consultora_cobro].iloc[0]

                # Armar mensaje de cobro
                nom_c_p = str(row_cobro.get('nombre', '')).split()[0].title()
                val_deuda_str = formato_cop(row_cobro.get('saldo_total', 0))
                d_rest = int(limpiar_numero(row_cobro.get('dias_para_vencer', 0)))
                fact_num = str(row_cobro.get('numero_factura', ''))
                
                if d_rest < 0:
                    estado_cobro = f"tu factura {fact_num} presenta *{abs(d_rest)} día(s) de mora*"
                elif d_rest == 0:
                    estado_cobro = f"tu factura {fact_num} *vence hoy*"
                elif d_rest == 1:
                    estado_cobro = f"tu factura {fact_num} *vence mañana*"
                else:
                    estado_cobro = f"tu factura {fact_num} vencerá en *{d_rest} días*"

                remitente_cobro = user_nombre if user_nombre else "Tu Líder"
                msg_cobro = (
                    f"Hola *{nom_c_p}* 🌸, te saluda tu Líder {remitente_cobro} de *Natura & Avon*.\n\n"
                    f"Queremos recordarte cordialmente que {estado_cobro} por valor de *{val_deuda_str}*.\n\n"
                    f"Te invitamos a realizar tu pago oportunamente para mantener tu crédito activo y seguir disfrutando de tus beneficios. ¡Quedo atenta para apoyarte! 💳✨"
                )

                tel_raw = str(row_cobro.get('telefono_movil', row_cobro.get('telefono_movil_2', ''))).strip().replace(' ', '').replace('-', '').replace('+', '')
                tel_c = tel_raw.split('.')[0] if '.' in tel_raw else tel_raw
                link_cobro_wa = f"https://api.whatsapp.com/send?phone=57{tel_c}&text={urllib.parse.quote(msg_cobro)}" if tel_c and len(tel_c) >= 10 else ""

                if link_cobro_wa:
                    st.link_button(f"📲 Cobrar por WhatsApp a {nom_c_p} ({val_deuda_str})", url=link_cobro_wa, type="primary", use_container_width=True)
                else:
                    st.warning("⚠️ Esta consultora no tiene teléfono móvil registrado.")

                # Tabla resumen de cartera
                cols_g_disp = ['nombre', 'codigo_cb', 'sit_comercial', 'numero_factura', 'fecha_vencimiento', 'dias_num', 'saldo_num']
                cols_g_exist = [c for c in cols_g_disp if c in df_g_filt.columns]
                df_g_view = df_g_filt[cols_g_exist].copy()
                df_g_view['saldo_num'] = df_g_view['saldo_num'].apply(formato_cop)
                df_g_view.rename(columns={
                    'nombre': 'Consultora',
                    'codigo_cb': 'Código CB',
                    'sit_comercial': 'Sit. Comercial',
                    'numero_factura': 'Factura',
                    'fecha_vencimiento': 'Vence',
                    'dias_num': 'Días Rest.',
                    'saldo_num': 'Saldo Deuda'
                }, inplace=True)
                st.dataframe(df_g_view, use_container_width=True, hide_index=True)

    # ==============================================================================
    # TAB 4: MIS LÍDERES (TODAS LAS TABLAS Y AVANCE COMPARATIVO)
    # ==============================================================================
    with tab_lideres:
        st.markdown("##### 👑 Mis Líderes - Diagnóstico y Comparativo")

        if df_cv_all is None or df_cv_all.empty:
            st.info("ℹ️ No hay datos de 'Cómo Vamos' disponibles para mostrar las tablas dinámicas de líderes.")
        else:
            df_diag = df_cv_all.copy()
            col_lider = 'Nombre de consultora' if 'Nombre de consultora' in df_diag.columns else (df_diag.columns[0] if len(df_diag.columns) > 0 else '')

            # Clasificación Oficial Estratégica:
            try:
                arte_papa_dict = cargar_objetivos_arte()
                grps_oficiales_arte = set(str(k).strip() for k in arte_papa_dict.get('por_grupo', {}).keys())
            except Exception:
                grps_oficiales_arte = set()

            try:
                gan_arte_dict = cargar_datos_ganancia_arte()
                ce_grps_gan = set(str(x.get('grupo_ce')).strip().split('.')[0] for x in gan_arte_dict.get('ce_plus', []) if x.get('grupo_ce'))
                ce_cods_gan = set(str(x.get('cod_ce')).strip().split('.')[0] for x in gan_arte_dict.get('ce_plus', []) if x.get('cod_ce'))
            except Exception:
                ce_grps_gan = set()
                ce_cods_gan = set()

            col_grp_diag_tipo = next((c for c in df_diag.columns if any(k in str(c).lower() for k in ['código de grupo', 'codigo de grupo', 'cód. grupo', 'cod grupo', 'grupo'])), None)
            col_cod_diag_tipo = next((c for c in df_diag.columns if any(k in str(c).lower() for k in ['código de consultora', 'codigo de consultora', 'cód. consultora', 'cod consultora'])), None)

            def _resolver_tipo_red_mob(row):
                g = str(row.get(col_grp_diag_tipo, '')).strip().split('.')[0] if col_grp_diag_tipo else ''
                c = str(row.get(col_cod_diag_tipo, '')).strip().split('.')[0] if col_cod_diag_tipo else ''
                if ce_grps_gan and (g in ce_grps_gan or c in ce_cods_gan):
                    return '🌱 CE+'
                if grps_oficiales_arte:
                    return '👑 LN' if g in grps_oficiales_arte else '🌱 CE+'
                col_obj_f = 'Objetivo Facturación' if 'Objetivo Facturación' in row else None
                if col_obj_f:
                    return '👑 LN' if limpiar_numero(row.get(col_obj_f, 0.0), 0.0) >= 25000000.0 else '🌱 CE+'
                return '👑 LN'

            if not df_diag.empty:
                df_diag['Tipo_Red'] = df_diag.apply(_resolver_tipo_red_mob, axis=1)
            else:
                df_diag['Tipo_Red'] = pd.Series(dtype=str)

            count_tot = len(df_diag)
            count_lideres = int((df_diag['Tipo_Red'] == '👑 LN').sum())
            count_emprendedoras = int((df_diag['Tipo_Red'] == '🌱 CE+').sum())

            def _renderizar_tablas_lideres_mob(df_mob_suite, key_prefix="mob_ln"):
                if df_mob_suite is None or df_mob_suite.empty:
                    st.info("ℹ️ No hay datos disponibles para mostrar en este segmento.")
                    return
                df_diag = df_mob_suite.copy()

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

                df_fact_view = df_fact_sorted[cols_presentes].copy().reset_index(drop=True)

                nombres_clery = {
                    col_lider: 'LÍDER DE NEGOCIOS',
                    'Tipo_Red': 'TIPO',
                    'Objetivo Facturación': 'Desafío\nFacturación',
                    'Real Facturación': 'Facturación\na Hoy',
                    'Cumplimiento Facturación': 'Cumplimiento de\nFacturación',
                    'Avance % Facturación': 'Avance %',
                    'Productividad': 'Productividad',
                    'Falta para el 100%': 'Falta para\nel 100%',
                    'Falta para el 110%': 'Falta para\nel 110%',
                    'Ganancia estimada': 'Ganancia\nEstimada'
                }
                df_fact_view = df_fact_view.rename(columns=nombres_clery)

                df_fact_formatted = df_fact_view.copy()
                if 'Desafío\nFacturación' in df_fact_formatted.columns:
                    df_fact_formatted['Desafío\nFacturación'] = df_fact_formatted['Desafío\nFacturación'].apply(formato_cop)
                if 'Facturación\na Hoy' in df_fact_formatted.columns:
                    df_fact_formatted['Facturación\na Hoy'] = df_fact_formatted['Facturación\na Hoy'].apply(formato_cop)
                if 'Cumplimiento de\nFacturación' in df_fact_formatted.columns:
                    df_fact_formatted['Cumplimiento de\nFacturación'] = df_fact_formatted['Cumplimiento de\nFacturación'].apply(formato_porcentaje)
                if 'Avance %' in df_fact_formatted.columns:
                    df_fact_formatted['Avance %'] = df_fact_formatted['Avance %'].apply(formato_porcentaje)
                if 'Productividad' in df_fact_formatted.columns:
                    df_fact_formatted['Productividad'] = df_fact_formatted['Productividad'].apply(formato_cop)
                if 'Falta para\nel 100%' in df_fact_formatted.columns:
                    df_fact_formatted['Falta para\nel 100%'] = df_fact_formatted['Falta para\nel 100%'].apply(formato_cop)
                if 'Falta para\nel 110%' in df_fact_formatted.columns:
                    df_fact_formatted['Falta para\nel 110%'] = df_fact_formatted['Falta para\nel 110%'].apply(formato_cop)
                if 'Ganancia\nEstimada' in df_fact_formatted.columns:
                    df_fact_formatted['Ganancia\nEstimada'] = df_fact_formatted['Ganancia\nEstimada'].apply(formato_cop)

                styler_fact = df_fact_formatted.style

                def _estilo_tipo(val_str):
                    if 'LN' in str(val_str) or 'Líder' in str(val_str):
                        return 'background-color: #dbeafe; color: #1e40af; font-weight: bold;'
                    elif 'CE+' in str(val_str) or 'Emprendedora' in str(val_str) or 'Semilla' in str(val_str):
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
                if 'Cumplimiento de\nFacturación' in df_fact_formatted.columns:
                    styler_fact = aplicar_mapa_styler(styler_fact, _estilo_cump_fact, subset=['Cumplimiento de\nFacturación'])
                if 'Avance %' in df_fact_formatted.columns:
                    styler_fact = aplicar_mapa_styler(styler_fact, _estilo_avance_pct_fact, subset=['Avance %'])
                if 'Falta para\nel 100%' in df_fact_formatted.columns:
                    styler_fact = aplicar_mapa_styler(styler_fact, _estilo_falta_dinero, subset=['Falta para\nel 100%'])
                if 'Falta para\nel 110%' in df_fact_formatted.columns:
                    styler_fact = aplicar_mapa_styler(styler_fact, _estilo_falta_dinero, subset=['Falta para\nel 110%'])
                if 'Ganancia\nEstimada' in df_fact_formatted.columns:
                    styler_fact = aplicar_mapa_styler(styler_fact, _estilo_ganancia_total, subset=['Ganancia\nEstimada'])

                st.dataframe(
                    styler_fact,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "LÍDER DE NEGOCIOS": st.column_config.Column("LÍDER DE NEGOCIOS", pinned=True)
                    }
                )

                # --- 2. TABLA DE ACTIVAS / PEDIDOS ---
                st.markdown("---")
                col_t_act1, col_t_act2 = st.columns([2, 1])
                with col_t_act1:
                    st.markdown("###### 👥 2. Tabla de Activas / Pedidos")
                with col_t_act2:
                    if st.button("🔍 Auditar Activas", key=f"mob_drill_tab4_act_{key_prefix}", use_container_width=True):
                        tot_r_act = df_diag['Real Activas'].apply(lambda v: limpiar_numero(v, 0.0)).sum() if 'Real Activas' in df_diag.columns else 0
                        tot_o_act = df_diag['Objetivo Activas'].apply(lambda v: limpiar_numero(v, 0.0)).sum() if 'Objetivo Activas' in df_diag.columns else 0
                        cump_t = (tot_r_act / tot_o_act * 100.0) if tot_o_act > 0 else 0.0
                        dialog_origen_activas(df_diag, user_sector if user_sector else "Sector", tot_r_act, tot_o_act, cump_t)

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

                df_act_view = df_act_sorted[cols_act_presentes].copy().reset_index(drop=True)
                nombres_clery_act = {
                    col_lider: 'LÍDER DE NEGOCIOS',
                    'Tipo_Red': 'TIPO',
                    'Objetivo Activas': 'Meta\nActivas',
                    'Real Activas': 'Activas\nHoy',
                    'Cumplimiento Activas': 'Cumplimiento\nActivas',
                    'Saldo': 'Saldo\nActivas',
                    'Disponibles': 'Disponibles',
                    'Inicios': 'Inicios\nHoy',
                    'Reinicios': 'Reinicios\nHoy',
                    'Recuperos': 'Recuperos\nHoy'
                }
                df_act_view = df_act_view.rename(columns=nombres_clery_act)

                df_act_formatted = df_act_view.copy()
                if 'Meta\nActivas' in df_act_formatted.columns:
                    df_act_formatted['Meta\nActivas'] = df_act_formatted['Meta\nActivas'].apply(lambda v: f"{int(limpiar_numero(v))}")
                if 'Activas\nHoy' in df_act_formatted.columns:
                    df_act_formatted['Activas\nHoy'] = df_act_formatted['Activas\nHoy'].apply(lambda v: f"{int(limpiar_numero(v))}")
                if 'Cumplimiento\nActivas' in df_act_formatted.columns:
                    df_act_formatted['Cumplimiento\nActivas'] = df_act_formatted['Cumplimiento\nActivas'].apply(formato_porcentaje)
                if 'Saldo\nActivas' in df_act_formatted.columns:
                    df_act_formatted['Saldo\nActivas'] = df_act_formatted['Saldo\nActivas'].apply(formato_saldo_entero)
                if 'Disponibles' in df_act_formatted.columns:
                    df_act_formatted['Disponibles'] = df_act_formatted['Disponibles'].apply(lambda v: f"{int(limpiar_numero(v))}")
                if 'Inicios\nHoy' in df_act_formatted.columns:
                    df_act_formatted['Inicios\nHoy'] = df_act_formatted['Inicios\nHoy'].apply(lambda v: f"{int(limpiar_numero(v))}")
                if 'Reinicios\nHoy' in df_act_formatted.columns:
                    df_act_formatted['Reinicios\nHoy'] = df_act_formatted['Reinicios\nHoy'].apply(lambda v: f"{int(limpiar_numero(v))}")
                if 'Recuperos\nHoy' in df_act_formatted.columns:
                    df_act_formatted['Recuperos\nHoy'] = df_act_formatted['Recuperos\nHoy'].apply(lambda v: f"{int(limpiar_numero(v))}")

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
                if 'Cumplimiento\nActivas' in df_act_formatted.columns:
                    styler_act = aplicar_mapa_styler(styler_act, _estilo_cump_act, subset=['Cumplimiento\nActivas'])
                if 'Saldo\nActivas' in df_act_formatted.columns:
                    styler_act = aplicar_mapa_styler(styler_act, _estilo_saldo_act, subset=['Saldo\nActivas'])
                for col_ing_sub in ['Inicios\nHoy', 'Reinicios\nHoy', 'Recuperos\nHoy']:
                    if col_ing_sub in df_act_formatted.columns:
                        styler_act = aplicar_mapa_styler(styler_act, _estilo_ingresos_act, subset=[col_ing_sub])

                st.dataframe(
                    styler_act,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "LÍDER DE NEGOCIOS": st.column_config.Column("LÍDER DE NEGOCIOS", pinned=True)
                    }
                )

                # --- 3. CUADRO RESUMEN DE DISPONIBLES ---
                st.markdown("---")
                st.markdown("###### 📋 3. Cuadro Resumen de Disponibles (Desafío vs. Avance por Día)")
                dia_corte = st.number_input("📅 Día de Avance (Editable):", min_value=1, max_value=21, value=14, step=1, key=f"mob_dia_corte_14_{key_prefix}")
                nombre_col_dia = f"Dia {dia_corte}"

                if col_lider and col_lider in df_diag.columns and not df_diag.empty:
                    df_disp_prep = df_diag.copy()
                    col_grp_diag = next((c for c in df_disp_prep.columns if any(k in str(c).lower() for k in ['código de grupo', 'codigo de grupo', 'cód. grupo', 'cod grupo', 'grupo'])), None)

                    mapa_grp_disp = mapa_arte.get('por_grupo', {})
                    mapa_nom_disp = mapa_arte.get('por_nombre', {})

                    col_disp_actual = 'Disponibles' if 'Disponibles' in df_disp_prep.columns else ('Real Activas' if 'Real Activas' in df_disp_prep.columns else None)

                    if col_disp_actual and not df_disp_prep.empty:
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

                        df_disp_calc['% Cump LN'] = np.where(
                            df_disp_calc['Disponibles Proyectadas'] > 0,
                            (df_disp_calc[nombre_col_dia] / df_disp_calc['Disponibles Proyectadas'] * 100.0),
                            0.0
                        )
                        df_disp_calc['falta'] = (df_disp_calc['Disponibles Proyectadas'] - df_disp_calc[nombre_col_dia]).clip(lower=0).astype(int)

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

                        st.dataframe(
                            styler_disp,
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "LÍDER DE NEGOCIOS": st.column_config.Column("LÍDER DE NEGOCIOS", pinned=True)
                            }
                        )

                # --- 4. CUADRO RESUMEN DE INICIOS + REINICIOS ---
                st.markdown("---")
                st.markdown("###### 🚀 4. Cuadro Resumen de Inicios + Reinicios")

                if col_lider and col_lider in df_diag.columns and not df_diag.empty:
                    df_ing_prep = df_diag.copy()
                    col_inicios = 'Inicios' if 'Inicios' in df_ing_prep.columns else None
                    col_reinicios = 'Reinicios' if 'Reinicios' in df_ing_prep.columns else None
                    col_meta_ing = next((c for c in df_ing_prep.columns if any(k in str(c).lower() for k in ['meta inicios + reinicios', 'meta_inicios_reinicios', 'meta inicios', 'meta_inicios', 'inicios + reinicios'])), None)

                    df_ing_calc = pd.DataFrame()
                    df_ing_calc['LÍDER DE NEGOCIOS'] = df_ing_prep[col_lider].astype(str)

                    val_inicios = df_ing_prep[col_inicios].apply(lambda v: limpiar_numero(v, 0)) if col_inicios else pd.Series(0, index=df_ing_prep.index)
                    val_reinicios = df_ing_prep[col_reinicios].apply(lambda v: limpiar_numero(v, 0)) if col_reinicios else pd.Series(0, index=df_ing_prep.index)
                    df_ing_calc['Hoy'] = (val_inicios + val_reinicios).astype(int)

                    if col_meta_ing:
                        df_ing_calc['Meta'] = df_ing_prep[col_meta_ing].apply(lambda v: int(limpiar_numero(v, 0)))
                        meta_fallback = (df_ing_calc['Hoy'] + 3).clip(lower=5)
                        df_ing_calc['Meta'] = np.where(df_ing_calc['Meta'] > 0, df_ing_calc['Meta'], meta_fallback).astype(int)
                    else:
                        df_ing_calc['Meta'] = (df_ing_calc['Hoy'] + 3).clip(lower=5).astype(int)

                    df_ing_calc['Avance'] = np.where(
                        df_ing_calc['Meta'] > 0,
                        (df_ing_calc['Hoy'] / df_ing_calc['Meta'] * 100.0),
                        0.0
                    )
                    df_ing_calc['para activar!'] = (df_ing_calc['Meta'] - df_ing_calc['Hoy']).clip(lower=0).astype(int)

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

                    st.dataframe(
                        styler_ing,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "LÍDER DE NEGOCIOS": st.column_config.Column("LÍDER DE NEGOCIOS", pinned=True)
                        }
                    )

                # --- 5. CUADRO RESUMEN DE RECUPEROS ---
                st.markdown("---")
                st.markdown("###### 🎯 5. Cuadro Resumen de Recuperos")

                if col_lider and col_lider in df_diag.columns and not df_diag.empty:
                    df_rec_prep = df_diag.copy()
                    col_recuperos = 'Recuperos' if 'Recuperos' in df_rec_prep.columns else None
                    col_meta_rec = next((c for c in df_rec_prep.columns if any(k in str(c).lower() for k in ['meta recuperos', 'meta_recuperos', 'recuperos_meta'])), None)

                    df_rec_calc = pd.DataFrame()
                    df_rec_calc['LÍDER DE NEGOCIOS'] = df_rec_prep[col_lider].astype(str)

                    val_rec = df_rec_prep[col_recuperos].apply(lambda v: limpiar_numero(v, 0)) if col_recuperos else pd.Series(0, index=df_rec_prep.index)
                    df_rec_calc['Hoy'] = val_rec.astype(int)

                    if col_meta_rec:
                        df_rec_calc['Meta'] = df_rec_prep[col_meta_rec].apply(lambda v: int(limpiar_numero(v, 0)))
                        meta_rec_fallback = (df_rec_calc['Hoy'] + 2).clip(lower=4)
                        df_rec_calc['Meta'] = np.where(df_rec_calc['Meta'] > 0, df_rec_calc['Meta'], meta_rec_fallback).astype(int)
                    else:
                        df_rec_calc['Meta'] = (df_rec_calc['Hoy'] + 2).clip(lower=4).astype(int)

                    df_rec_calc['Avance'] = np.where(
                        df_rec_calc['Meta'] > 0,
                        (df_rec_calc['Hoy'] / df_rec_calc['Meta'] * 100.0),
                        0.0
                    )
                    df_rec_calc['para activar!'] = (df_rec_calc['Meta'] - df_rec_calc['Hoy']).clip(lower=0).astype(int)

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

                    st.dataframe(
                        styler_rec,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "LÍDER DE NEGOCIOS": st.column_config.Column("LÍDER DE NEGOCIOS", pinned=True)
                        }
                    )

                # --- 6. CUADRO RESUMEN DE RETENCIÓN I2 ---
                st.markdown("---")
                st.markdown("###### 🔄 6. Cuadro Resumen de Retención I2 (Meta 8% Máx. Fuga I2)")

                if col_lider and col_lider in df_diag.columns and not df_diag.empty:
                    df_i2_prep = df_diag.copy()
                    col_disp_i2 = 'Disponibles' if 'Disponibles' in df_i2_prep.columns else None
                    col_i2 = next((c for c in df_i2_prep.columns if str(c).lower().strip() in ['inactiva 2', 'inactiva_2', 'inactivas 2', 'inactivas_2', 'i2']), None)
                    col_i2_ant = next((c for c in df_i2_prep.columns if 'inactiva 2_anterior' in str(c).lower() or 'inactivas 2_anterior' in str(c).lower()), None)

                    if col_disp_i2 and col_i2 and not df_i2_prep.empty:
                        df_i2_calc = pd.DataFrame()
                        df_i2_calc['LÍDER DE NEGOCIOS'] = df_i2_prep[col_lider].astype(str)

                        val_disp2 = df_i2_prep[col_disp_i2].apply(lambda v: limpiar_numero(v, 0.0))
                        val_i2 = df_i2_prep[col_i2].apply(lambda v: limpiar_numero(v, 0.0))

                        df_i2_calc['MAX PANEL'] = (val_disp2 * 0.08).round().astype(int)
                        df_i2_calc['FALTA I2 ACTIVARSE'] = (val_i2 - df_i2_calc['MAX PANEL']).round().astype(int)
                        df_i2_calc['% RETENCIÓN META 8%'] = np.where(val_disp2 > 0, (val_i2 / val_disp2 * 100.0), 0.0)

                        if col_i2_ant:
                            val_i2_ant = df_i2_prep[col_i2_ant].apply(lambda v: limpiar_numero(v, 0.0))
                            df_i2_calc['AVANCE RETENCION I2'] = (val_i2_ant - val_i2).fillna(0).astype(int)
                        else:
                            df_i2_calc['AVANCE RETENCION I2'] = 0

                        df_i2_calc = df_i2_calc.sort_values(by='% RETENCIÓN META 8%', ascending=True).reset_index(drop=True)

                        tot_disp_i2 = float(val_disp2.sum())
                        tot_i2 = float(val_i2.sum())
                        tot_meta_i2 = int(df_i2_calc['MAX PANEL'].sum())
                        tot_falta_i2 = int(df_i2_calc['FALTA I2 ACTIVARSE'].sum())
                        tot_pct_i2 = (tot_i2 / tot_disp_i2 * 100.0) if tot_disp_i2 > 0 else 0.0
                        tot_av_i2 = int(df_i2_calc['AVANCE RETENCION I2'].sum())

                        row_tot_i2 = pd.DataFrame([{
                            'LÍDER DE NEGOCIOS': 'TOTAL GENERAL',
                            'MAX PANEL': tot_meta_i2,
                            'FALTA I2 ACTIVARSE': tot_falta_i2,
                            '% RETENCIÓN META 8%': tot_pct_i2,
                            'AVANCE RETENCION I2': tot_av_i2
                        }])
                        df_i2_final = pd.concat([df_i2_calc, row_tot_i2], ignore_index=True)

                        df_i2_formatted = df_i2_final[['LÍDER DE NEGOCIOS', 'MAX PANEL', 'FALTA I2 ACTIVARSE', '% RETENCIÓN META 8%', 'AVANCE RETENCION I2']].copy()
                        df_i2_formatted['MAX PANEL'] = df_i2_formatted['MAX PANEL'].apply(lambda v: f"{int(v):,}".replace(",", "."))
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

                        st.dataframe(
                            styler_i2,
                            use_container_width=True,
                            hide_index=True,
                            column_config={
                                "LÍDER DE NEGOCIOS": st.column_config.Column("LÍDER DE NEGOCIOS", pinned=True)
                            }
                        )



            # =========================================================================
            # MÓDULO DE REPORTE DE DESEMPEÑO POR WHATSAPP A LÍDER (CÓMO VAMOS)
            # =========================================================================
            with st.expander("📲 Enviar Reporte de Desempeño por WhatsApp a Líder", expanded=False):
                st.markdown("###### 📲 Reporte Oficial de Facturación, Activas y Desafíos")
                st.caption("Genera y envía por WhatsApp el reporte integral a cualquier líder con métricas de 95%, 100%, 110%, Inicios, Reinicios, Saldo y Disponibles.")

                # Construir mapa de celulares
                mapa_cel_lid_mob = {}
                try:
                    conn_cel_m = procesador.obtener_conexion_db(timeout=5.0)
                    cursor_cel_m = conn_cel_m.cursor()
                    cursor_cel_m.execute("SELECT codigo_cb, nombre, celular, grupo FROM consultoras_tableau WHERE celular IS NOT NULL AND TRIM(celular) != ''")
                    for r_cm in cursor_cel_m.fetchall():
                        cb_cm = str(r_cm[0]).strip().split('.')[0]
                        nom_cm = str(r_cm[1]).strip().upper()
                        cel_vm = str(r_cm[2]).strip()
                        grp_cm = str(r_cm[3]).strip().split('.')[0] if r_cm[3] else ""
                        if cb_cm and cb_cm != '0':
                            mapa_cel_lid_mob[f"cb_{cb_cm}"] = cel_vm
                        if nom_cm:
                            mapa_cel_lid_mob[f"nom_{nom_cm}"] = cel_vm
                        if grp_cm and grp_cm != '0' and f"grp_{grp_cm}" not in mapa_cel_lid_mob:
                            mapa_cel_lid_mob[f"grp_{grp_cm}"] = cel_vm
                    conn_cel_m.close()
                except Exception:
                    pass

                try:
                    usr_cat_m = cargar_usuarios()
                    for _, uv_m in usr_cat_m.items():
                        if isinstance(uv_m, dict) and uv_m.get('telefono'):
                            un_m = str(uv_m.get('nombre', '')).strip().upper()
                            ug_m = str(uv_m.get('codigo_grupo', '')).strip().split('.')[0]
                            tf_m = str(uv_m.get('telefono')).strip()
                            if un_m and tf_m:
                                mapa_cel_lid_mob[f"nom_{un_m}"] = tf_m
                            if ug_m and tf_m and f"grp_{ug_m}" not in mapa_cel_lid_mob:
                                mapa_cel_lid_mob[f"grp_{ug_m}"] = tf_m
                except Exception:
                    pass

                # Extraer opciones de líderes de df_diag
                opciones_lideres_rep = []
                for _, r_lid in df_diag.iterrows():
                    nom_lid = str(r_lid.get(col_lider, '')).strip()
                    if not nom_lid or nom_lid.lower() in ['nan', 'none', '', '0', 'null', 'total general']:
                        continue
                    cb_lid = str(r_lid.get('Código de consultora', '')).strip().split('.')[0]
                    if cb_lid == '0':
                        cb_lid = ''
                    col_g_tmp = next((c for c in r_lid.index if 'grupo' in str(c).lower()), None)
                    grp_lid = str(r_lid.get(col_g_tmp, '')).strip().split('.')[0] if col_g_tmp else ''

                    cel_lid = mapa_cel_lid_mob.get(f"cb_{cb_lid}") or mapa_cel_lid_mob.get(f"nom_{nom_lid.upper()}") or mapa_cel_lid_mob.get(f"grp_{grp_lid}") or ""
                    opciones_lideres_rep.append({
                        'nombre': nom_lid,
                        'grupo': grp_lid,
                        'codigo_cb': cb_lid,
                        'celular': cel_lid,
                        'row': r_lid
                    })

                if opciones_lideres_rep:
                    nombres_sel_rep = [f"{it['nombre']}{' (Grp ' + it['grupo'] + ')' if it['grupo'] else ''}" for it in opciones_lideres_rep]
                    idx_lid_rep = st.selectbox("👤 Selecciona la Líder:", range(len(opciones_lideres_rep)), format_func=lambda i: nombres_sel_rep[i], key="sel_lider_rep_wa_mob")
                    lid_sel_info = opciones_lideres_rep[idx_lid_rep]
                    r_sel = lid_sel_info['row']
                    nom_sel = lid_sel_info['nombre']
                    cel_sel = lid_sel_info['celular']

                    # Generar texto del reporte oficial con faltantes 95%, 100%, 110%
                    r_fact_num = float(limpiar_numero(r_sel.get('Real Facturación', 0.0), 0.0))
                    o_fact_num = float(limpiar_numero(r_sel.get('Objetivo Facturación', 0.0), 0.0))
                    c_fact_num = (r_fact_num / o_fact_num * 100.0) if o_fact_num > 0 else 0.0
                    c_fact_str = f"{c_fact_num:.1f}%"

                    meta_95_fact = o_fact_num * 0.95
                    falta_95_fact = max(0.0, meta_95_fact - r_fact_num)
                    falta_95_fact_str = "¡Logrado! 🎉" if falta_95_fact == 0 and o_fact_num > 0 and r_fact_num >= meta_95_fact else formato_cop(falta_95_fact)

                    falta_100_fact = max(0.0, o_fact_num - r_fact_num)
                    falta_100_fact_str = "¡Logrado! 🎉" if falta_100_fact == 0 and o_fact_num > 0 and r_fact_num >= o_fact_num else formato_cop(falta_100_fact)

                    meta_110_fact = o_fact_num * 1.10
                    falta_110_fact = max(0.0, meta_110_fact - r_fact_num)
                    falta_110_fact_str = "¡Logrado! 🎉" if falta_110_fact == 0 and o_fact_num > 0 and r_fact_num >= meta_110_fact else formato_cop(falta_110_fact)

                    r_act_num = float(limpiar_numero(r_sel.get('Real Activas', 0), 0))
                    o_act_num = float(limpiar_numero(r_sel.get('Objetivo Activas', 0), 0))
                    c_act_num = (r_act_num / o_act_num * 100.0) if o_act_num > 0 else 0.0
                    c_act_str = f"{c_act_num:.1f}%"

                    meta_95_act = int(round(o_act_num * 0.95))
                    falta_95_act = max(0, meta_95_act - int(r_act_num))
                    falta_95_act_str = "¡Logrado! 🎉" if falta_95_act == 0 and o_act_num > 0 and r_act_num >= meta_95_act else f"{falta_95_act} activas"

                    falta_100_act = max(0, int(o_act_num) - int(r_act_num))
                    falta_100_act_str = "¡Logrado! 🎉" if falta_100_act == 0 and o_act_num > 0 and r_act_num >= int(o_act_num) else f"{falta_100_act} activas"

                    meta_110_act = int(round(o_act_num * 1.10))
                    falta_110_act = max(0, meta_110_act - int(r_act_num))
                    falta_110_act_str = "¡Logrado! 🎉" if falta_110_act == 0 and o_act_num > 0 and r_act_num >= meta_110_act else f"{falta_110_act} activas"

                    sal_num = int(round(limpiar_numero(r_sel.get('Saldo', 0), 0)))
                    sal_str = f"{sal_num:+d}" if sal_num != 0 else "0"

                    disp_num = int(round(limpiar_numero(r_sel.get('Disponibles', 0), 0)))
                    inicios_num = int(round(limpiar_numero(r_sel.get('Inicios', 0), 0)))
                    reinicios_num = int(round(limpiar_numero(r_sel.get('Reinicios', 0), 0)))
                    tot_inicios_reinicios = inicios_num + reinicios_num
                    meta_ir_num = int(round(limpiar_numero(r_sel.get('Meta Inicios + Reinicios', 0), 0)))
                    meta_ir_suffix = f" (Meta: {meta_ir_num})" if meta_ir_num > 0 else ""

                    gan_l = formato_cop(r_sel.get('Ganancia estimada', 0))
                    sector_l = str(r_sel.get('Nombre Setor', user_sector or 'Sector')).strip()

                    msg_reporte_lider = (
                        f"📊 *REPORTE CÓMO VAMOS*\n"
                        f"👤 *Líder:* {nom_sel}\n"
                        f"📍 *Sector:* {sector_l}\n\n"
                        f"💰 *--- FACTURACIÓN ---*\n"
                        f"💵 *Facturación Real:* {formato_cop(r_fact_num)}\n"
                        f"🎯 *Objetivo Facturación:* {formato_cop(o_fact_num)}\n"
                        f"📈 *Cumplimiento Facturación:* {c_fact_str}\n"
                        f"⚡ *Falta para 95% (Mínimo):* {falta_95_fact_str}\n"
                        f"💵 *Falta para 100%:* {falta_100_fact_str}\n"
                        f"🚀 *Falta para 110%:* {falta_110_fact_str}\n\n"
                        f"👥 *--- ACTIVAS & DISPONIBLES ---*\n"
                        f"👥 *Activas Reales:* {int(r_act_num)}\n"
                        f"🎯 *Objetivo Activas:* {int(o_act_num)}\n"
                        f"📈 *Cumplimiento Activas:* {c_act_str}\n"
                        f"⚡ *Falta para 95% (Mínimo):* {falta_95_act_str}\n"
                        f"🌱 *Falta para 100%:* {falta_100_act_str}\n"
                        f"🚀 *Falta para 110%:* {falta_110_act_str}\n"
                        f"📋 *Disponibles en Red:* {disp_num} consultoras\n\n"
                        f"🚀 *--- INICIOS, REINICIOS & SALDO ---*\n"
                        f"🌱 *Inicios:* {inicios_num}\n"
                        f"🔄 *Reinicios:* {reinicios_num}\n"
                        f"✨ *Total Inicios + Reinicios:* {tot_inicios_reinicios}{meta_ir_suffix}\n"
                        f"⚖️ *Saldo Comercial:* {sal_str}\n"
                        f"💵 *Ganancia Estimada:* {gan_l}\n"
                    )

                    c_rep_t1, c_rep_t2 = st.columns([2, 1])
                    with c_rep_t1:
                        txt_rep_edit = st.text_area("✏️ Mensaje del Reporte:", value=msg_reporte_lider, height=180, key="txt_rep_lider_mob_edit")
                    with c_rep_t2:
                        tel_rep_input = st.text_input("📱 Celular Líder:", value=cel_sel, key="txt_tel_lider_mob_edit")
                        tel_clean_r = tel_rep_input.strip().replace(' ', '').replace('-', '').replace('+', '')
                        link_wa_rep = f"https://api.whatsapp.com/send?phone=57{tel_clean_r}&text={urllib.parse.quote(txt_rep_edit)}" if len(tel_clean_r) >= 10 else ""
                        if link_wa_rep:
                            st.link_button(f"📲 Enviar WhatsApp a {nom_sel.split()[0]}", url=link_wa_rep, type="primary", use_container_width=True)
                        else:
                            st.warning("⚠️ Ingresa un celular válido (10 dígitos)")
                else:
                    st.info("ℹ️ No hay líderes registradas en la vista actual.")

            # =========================================================================
            # SUBPESTAÑAS DE SEPARACIÓN ESTRATÉGICA: LN vs CE+ vs TODA LA RED
            # =========================================================================
            subtab_mob_ln, subtab_mob_ce, subtab_mob_red = st.tabs([
                f"👑 LN ({count_lideres})",
                f"🌱 CE+ ({count_emprendedoras})",
                f"🌟 TODA LA RED ({count_tot})"
            ])

            with subtab_mob_ln:
                st.markdown("###### 👑 Líderes de Negocio (LN Oficiales)")
                df_ln = df_diag[df_diag['Tipo_Red'] == '👑 LN'].copy() if not df_diag.empty else pd.DataFrame()

                if df_ln.empty:
                    if count_lideres == 0 and not df_diag.empty:
                        st.caption("ℹ️ Mostrando toda la red disponible (sube 'Objetivos Arte.xlsx' para separar automáticamente LN oficiales de CE+).")
                        _renderizar_tablas_lideres_mob(df_diag, key_prefix="mob_ln")
                    else:
                        st.info("ℹ️ No hay registros de Líderes de Negocio en esta vista.")
                else:
                    _renderizar_tablas_lideres_mob(df_ln, key_prefix="mob_ln")

            with subtab_mob_ce:
                st.markdown("###### 🌱 Metas de Crecimiento & Bonos Mentora (CE+)")
                sec_target = user_sector if user_rol == 'lider' else None
                df_ce_mob = consultar_ce_plus_df(sector=sec_target, df_como_vamos=df_diag)

                if user_rol == 'lider' and grupo_activo and not df_ce_mob.empty:
                    grp_u = str(grupo_activo).strip().split('.')[0]
                    mask_ce_u = (df_ce_mob['Cód. Grupo LN'].astype(str).str.strip() == grp_u) | \
                                (df_ce_mob['Grupo CE+'].astype(str).str.strip() == grp_u)
                    df_ce_mob = df_ce_mob[mask_ce_u]

                if df_ce_mob.empty:
                    st.info("ℹ️ No se encontraron Consultoras Emprende+ (CE+) asignadas para tu grupo o sector.")
                else:
                    tot_ce_m = len(df_ce_mob)
                    tot_act_m = sum(int(limpiar_numero(x, 0)) for x in df_ce_mob['Activas Hoy']) if 'Activas Hoy' in df_ce_mob.columns else 0
                    tot_crec_m = sum(int(limpiar_numero(x, 0)) for x in df_ce_mob['Crecimiento Activas']) if 'Crecimiento Activas' in df_ce_mob.columns else 0
                    tot_bono_m = sum(float(limpiar_numero(x, 0.0)) for x in df_ce_mob['Bono Mentora LN']) if 'Bono Mentora LN' in df_ce_mob.columns else 0.0

                    c_m1, c_m2, c_m3 = st.columns(3)
                    with c_m1:
                        st.metric("🌱 TOTAL CE+", f"{tot_ce_m}")
                    with c_m2:
                        st.metric("⚡ ACTIVAS", f"{tot_act_m}")
                    with c_m3:
                        st.metric("🎁 BONOS LN", formato_cop(tot_bono_m))

                    cols_deseadas_ce = [
                        'Grupo CE+', 'Consultora Emprende+', 'Activas Hoy', 'Crecimiento Activas',
                        'Meta 1+ (+150k)', 'Meta 3+ (+200k)', 'Meta 5+ (+300k)', 'Bono Mentora LN'
                    ]
                    cols_p_ce = [c for c in cols_deseadas_ce if c in df_ce_mob.columns]
                    df_ce_m_render = df_ce_mob[cols_p_ce].copy()

                    if 'Crecimiento Activas' in df_ce_m_render.columns:
                        def _formato_crec_signo_mob(v):
                            n = int(limpiar_numero(v, 0))
                            return f"{'+' if n > 0 else ''}{n}"
                        df_ce_m_render['Crecimiento Activas'] = df_ce_m_render['Crecimiento Activas'].apply(_formato_crec_signo_mob)

                    st.dataframe(df_ce_m_render, use_container_width=True, hide_index=True)

                df_ce_cv_mob = df_diag[df_diag['Tipo_Red'] == '🌱 CE+'].copy() if not df_diag.empty else pd.DataFrame()
                if not df_ce_cv_mob.empty:
                    with st.expander(f"📊 Desempeño Operativo de las {len(df_ce_cv_mob)} CE+ en Campaña", expanded=False):
                        _renderizar_tablas_lideres_mob(df_ce_cv_mob, key_prefix="mob_ce_op")

            with subtab_mob_red:
                st.markdown("###### 🌟 Toda la Red Consolidada (LN + CE+)")
                _renderizar_tablas_lideres_mob(df_diag, key_prefix="mob_red")


    st.markdown("---")
    st.markdown(f"""
    <div style="text-align: center; padding: 10px 0 15px 0; color: #94A3B8; font-size: 0.82rem; letter-spacing: 0.3px;">
        <span>📱 <b>App {nombre_sector_app} Móvil</b></span>
        <span style="margin: 0 8px; opacity: 0.4;">•</span>
        <span>Desarrollado por <b>Tao-System</b></span>
        <span style="margin: 0 6px; opacity: 0.4;">|</span>
        <span style="color: #64748B; font-weight: 500;">Powered by <b>XYZ</b></span>
    </div>
    """, unsafe_allow_html=True)



if __name__ == '__main__':
    try:
        st.set_page_config(
            page_title="App Móvil - Gestión Líderes",
            page_icon="📱",
            layout="wide",
            initial_sidebar_state="collapsed"
        )
    except Exception:
        pass

    # Control de Sesión y Login Móvil Standalone
    if 'user' not in st.session_state:
        st.session_state['user'] = None

    if st.session_state['user'] is None:
        token_url = st.query_params.get('session')
        if token_url:
            u_rest = validar_token_sesion(token_url)
            if u_rest:
                st.session_state['user'] = u_rest

    if st.session_state['user'] is None:
        st.markdown("<h3 style='text-align:center; margin-bottom:2px;'>📱 App Móvil Líderes</h3>", unsafe_allow_html=True)
        st.caption("<p style='text-align:center; margin-bottom:12px;'>Portal exclusivo para Líderes de Negocio en Celulares y Tablets</p>", unsafe_allow_html=True)
        
        with st.form("form_login_mobile"):
            input_u = st.text_input("👤 Usuario o Correo", placeholder="Ej. lider7841 o correo@...").strip().lower()
            input_p = st.text_input("🔒 Contraseña", type="password", placeholder="••••••••")
            btn_log = st.form_submit_button("🚀 Entrar al Sistema", type="primary", use_container_width=True)
            
            if btn_log:
                u_auth = autenticar_usuario(input_u, input_p)
                if u_auth:
                    st.session_state['user'] = u_auth
                    st.query_params['session'] = generar_token_sesion(u_auth)
                    registrar_evento_auditoria(
                        u_auth,
                        categoria="🔑 Acceso",
                        accion="Inicio de Sesión",
                        detalle=f"Ingreso exitoso a la App Móvil ({u_auth.get('rol', '')})",
                        dispositivo="📱 Móvil / Tablet"
                    )
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos.")
        st.stop()

    current_u = st.session_state.get('user') or {}
    if current_u.get('debe_cambiar_password', False):
        st.warning("🔒 **Cambio Obligatorio de Contraseña Inicial**")
        st.info("Por tu seguridad, debes definir una contraseña personal para continuar.")
        with st.form("form_pwd_mobile"):
            pn = st.text_input("🔑 Nueva Contraseña Personal", type="password")
            pc = st.text_input("🔑 Confirmar Nueva Contraseña", type="password")
            if st.form_submit_button("💾 Guardar y Continuar", use_container_width=True):
                if len(pn.strip()) < 4 or pn != pc:
                    st.error("❌ Las contraseñas deben coincidir y tener al menos 4 caracteres.")
                else:
                    ok, msg = cambiar_password_usuario(current_u['username'], pn)
                    if ok:
                        current_u['debe_cambiar_password'] = False
                        st.session_state['user'] = current_u
                        st.query_params['session'] = generar_token_sesion(current_u)
                        registrar_evento_auditoria(
                            current_u,
                            categoria="🔑 Seguridad",
                            accion="Cambio de Contraseña",
                            detalle="Actualización de contraseña personal en móvil",
                            dispositivo="📱 Móvil / Tablet"
                        )
                        st.success("✅ Contraseña actualizada.")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")
        if st.button("🚪 Cancelar y Salir", key="btn_cancel_mob_standalone", use_container_width=True):
            st.session_state['user'] = None
            st.query_params.clear()
            st.rerun()
        st.stop()

    render_vista_movil(current_u, mostrar_salir=True)
