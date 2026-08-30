import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import plotly.express as px
import plotly.graph_objects as go
import io
import os
import time
import urllib.parse
import importlib
import contextlib
import procesador

# Recargamos dinámicamente procesador para garantizar que cualquier cambio en procesador.py se aplique sin reiniciar el servidor
importlib.reload(procesador)
from procesador import (
    calcular_metas_ciclo,
    generar_analisis_como_vamos,
    procesar_base_tableau_manager,
    guardar_comentario_lider,
    guardar_todos_comentarios,
    autocorregir_texto_espanol,
    color_nivel,
    color_situacion,
    color_deuda_mora,
    actualizar_situacion_comercial_desde_mi_grupo,
    actualizar_base_desde_activas,
    autenticar_usuario,
    buscar_cuenta_usuario,
    cargar_usuarios,
    registrar_o_actualizar_usuario,
    cambiar_password_usuario,
    restablecer_password_usuario,
    cargar_configuracion,
    guardar_configuracion,
    DEFAULT_PERMISOS_PESTANAS,
    inicializar_db_sqlite,
    consultar_tableau_sql,
    sincronizar_excel_tableau_a_sqlite,
    sincronizar_excel_metas_a_sqlite,
    sincronizar_excel_geral_a_sqlite,
    consultar_geral_sql,
    procesar_analisis_geral_cobranza,
    generar_mensaje_whatsapp_cobranza,
    MATRIZ_GANANCIA,
    ETIQUETAS_ACTIVAS,
    ETIQUETAS_FACTURACION,
    calcular_matriz_ganancia,
    calcular_bono_lider_mentora,
    calcular_puntos_convencion_ciclo,
    obtener_diagnostico_retencion_grupo,
    obtener_potencializador_saldo,
    obtener_indice_activas,
    obtener_indice_facturacion,
    rotar_y_guardar_nuevo_ciclo,
    eliminar_datos_por_grupo_o_usuario,
    vaciar_base_datos_completa,
    color_cumplimiento,
    color_avance,
    color_saldo,
    exportar_excel_con_colores,
    exportar_tableau_excel_con_colores,
    limpiar_y_ordenar_columnas_tableau,
    obtener_base_tableau_completa_original,
    limpiar_numero,
    validar_sector_archivo,
    validar_archivo_como_vamos,
    auto_crear_usuarios_lideres_desde_bases,
    obtener_mapa_lideres,
    cargar_historico_sectores,
    guardar_historico_sectores,
    verificar_estado_suscripcion,
    registrar_nueva_gerente,
    actualizar_suscripcion_sector,
    obtener_resumen_suscripciones,
    eliminar_usuario_perfil,
    obtener_nombre_sector_usuario,
    procesar_archivo_objetivos_arte,
    cargar_objetivos_arte,
    cargar_catalogo_sectores,
    extraer_catalogo_sectores_desde_arte,
    limpiar_nombre_sector_solo,
    obtener_cumpleanos_equipo,
    MESES_ESPANOL,
    PLANTILLA_CUMPLEANOS_DEFAULT
)

# 1. Configuración de la página
st.set_page_config(
    page_title="Panel de Control - Estado de Ciclo Líderes",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inicializar y verificar tablas relacionales de SQLite
try:
    inicializar_db_sqlite()
except Exception as _e_init:
    print(f"Nota de inicialización SQLite: {_e_init}")

# Helper global para aplicar mapas de estilo con compatibilidad de versiones de pandas
def aplicar_mapa_styler(styler, func, subset=None):
    try:
        if hasattr(styler, 'map'):
            return styler.map(func, subset=subset) if subset is not None else styler.map(func)
        elif hasattr(styler, 'applymap'):
            return styler.applymap(func, subset=subset) if subset is not None else styler.applymap(func)
    except Exception:
        pass
    return styler

# Wrappers con caché de alto rendimiento para descargas instantáneas
@st.cache_data(show_spinner=False)
def cached_export_excel_tableau(df):
    return exportar_tableau_excel_con_colores(df, nombre_hoja="Base_Consultoras")

@st.cache_data(show_spinner=False)
def cached_export_csv(df):
    return df.to_csv(index=False).encode('utf-8-sig')

# Estilos CSS personalizados para mejorar el diseño estético
st.markdown("""
<style>
    /* ----------------------------------------------------
       NATURA & AVON BRAND DESIGN SYSTEM (LIGHT & DARK MODE)
       Natura: Warm Vibrant Orange (#FF6B00 / #F58220)
       Avon: Royal Fuchsia Magenta (#E3007B / #9B0053)
       ---------------------------------------------------- */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Outfit', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Gradient Brand Text Headers */
    /* ========================================================= */
    /* SISTEMA DE DISEÑO ULTRA-RESPONSIVO (LAPTOPS, TABLETS, PC) */
    /* ========================================================= */
    
    .block-container {
        padding-top: clamp(1.2rem, 2.2vw, 2.2rem) !important;
        padding-bottom: clamp(2rem, 3vw, 3.5rem) !important;
        padding-left: clamp(1rem, 2vw, 2.8rem) !important;
        padding-right: clamp(1rem, 2vw, 2.8rem) !important;
        max-width: 100% !important;
    }

    .main-header {
        font-size: clamp(1.5rem, 2.2vw, 2.3rem) !important;
        font-weight: 800;
        background: linear-gradient(135deg, #FF6B00 0%, #F58220 30%, #E3007B 70%, #9B0053 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem;
        letter-spacing: -0.02em;
        line-height: 1.2;
    }
    
    .sub-header {
        font-size: clamp(0.82rem, 1vw, 1.02rem) !important;
        opacity: 0.85;
        margin-bottom: 1.2rem;
        font-weight: 500;
        line-height: 1.4;
    }

    /* Streamlit Metric Cards - Fluidas y Adaptables a Portátiles y Pantallas */
    [data-testid="stMetric"] {
        background: linear-gradient(135deg, rgba(255, 107, 0, 0.07) 0%, rgba(227, 0, 123, 0.07) 100%) !important;
        border: 1px solid rgba(227, 0, 123, 0.25) !important;
        border-radius: 16px !important;
        padding: clamp(10px, 1.1vw, 16px) clamp(10px, 1.2vw, 18px) !important;
        box-shadow: 0 8px 20px -4px rgba(227, 0, 123, 0.12) !important;
        backdrop-filter: blur(14px) !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        min-width: 0 !important;
        width: 100% !important;
        box-sizing: border-box !important;
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
    }

    [data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 24px -4px rgba(255, 107, 0, 0.22) !important;
        border-color: rgba(255, 107, 0, 0.6) !important;
    }

    [data-testid="stMetricLabel"] p {
        font-size: clamp(0.72rem, 0.85vw, 0.88rem) !important;
        font-weight: 700 !important;
        letter-spacing: 0.01em;
        opacity: 0.9;
        white-space: nowrap !important;
        text-overflow: ellipsis !important;
        overflow: hidden !important;
        line-height: 1.2 !important;
        margin: 0 0 2px 0 !important;
    }

    [data-testid="stMetricValue"] div {
        font-size: clamp(1.1rem, 1.5vw, 1.7rem) !important;
        font-weight: 800 !important;
        background: linear-gradient(135deg, #FF6B00 0%, #E3007B 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        white-space: nowrap !important;
        text-overflow: ellipsis !important;
        overflow: hidden !important;
        line-height: 1.2 !important;
    }

    [data-testid="stMetricDelta"] div {
        font-size: clamp(0.66rem, 0.76vw, 0.8rem) !important;
        font-weight: 600 !important;
        white-space: nowrap !important;
        text-overflow: ellipsis !important;
        overflow: hidden !important;
        line-height: 1.2 !important;
    }

    /* Auto-Wrapping inteligente de columnas Streamlit para Portátiles */
    div[data-testid="stHorizontalBlock"] {
        display: flex !important;
        flex-wrap: wrap !important;
        gap: clamp(8px, 1vw, 14px) !important;
        align-items: stretch !important;
    }

    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        min-width: 0 !important;
        box-sizing: border-box !important;
    }

    /* Reglas de Breakpoints para Portátiles / Laptops (1280px - 1440px) */
    @media (max-width: 1366px) {
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
            flex: 1 1 calc(33.333% - 10px) !important;
            min-width: 170px !important;
        }
    }

    /* Tablets y Pantallas Medianas (768px - 1024px) */
    @media (max-width: 1024px) {
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
            flex: 1 1 calc(50% - 10px) !important;
            min-width: 150px !important;
        }
    }

    /* Móviles (< 640px) */
    @media (max-width: 640px) {
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
            flex: 1 1 100% !important;
            min-width: 100% !important;
        }
    }

    /* Expanders & Cards */
    [data-testid="stExpander"] {
        background: rgba(255, 107, 0, 0.03) !important;
        border: 1px solid rgba(227, 0, 123, 0.2) !important;
        border-radius: 16px !important;
        overflow: hidden;
    }

    /* Primary Natura-Avon Gradient Buttons */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #FF6B00 0%, #E3007B 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 12px !important;
        font-weight: 700 !important;
        padding: 0.55rem 1.4rem !important;
        box-shadow: 0 6px 18px rgba(227, 0, 123, 0.25) !important;
        transition: all 0.25s ease !important;
    }

    div.stButton > button[kind="primary"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 24px rgba(255, 107, 0, 0.35) !important;
        background: linear-gradient(135deg, #F58220 0%, #9B0053 100%) !important;
    }

    /* Ocultar icono del ojo para revelar contraseñas / campos sensibles */
    button[aria-label*="password" i],
    button[aria-label*="contraseña" i],
    button[aria-label*="Show" i],
    button[aria-label*="Hide" i],
    [data-testid="stTextInputRootElement"] button,
    [data-testid="stTextInput"] button {
        display: none !important;
        visibility: hidden !important;
        pointer-events: none !important;
    }

    /* Executive Login Card Styling (Natura & Avon Theme) */
    .login-container {
        max-width: 1000px;
        margin: 0.5rem auto 1.5rem auto;
    }
    
    .login-hero-card {
        background: linear-gradient(135deg, #1C0A19 0%, #2D0F25 50%, #3B142F 100%);
        border-radius: 24px;
        padding: 35px 38px;
        color: #FFFFFF;
        box-shadow: 0 20px 40px -10px rgba(227, 0, 123, 0.3);
        border: 1px solid rgba(255, 107, 0, 0.3);
        margin-bottom: 25px;
        position: relative;
        overflow: hidden;
    }
    
    .login-hero-card::before {
        content: "";
        position: absolute;
        top: -30%;
        right: -10%;
        width: 400px;
        height: 400px;
        background: radial-gradient(circle, rgba(255, 107, 0, 0.35) 0%, rgba(227, 0, 123, 0.2) 40%, rgba(0,0,0,0) 70%);
        pointer-events: none;
    }

    .login-badge {
        display: inline-block;
        background: linear-gradient(135deg, rgba(255, 107, 0, 0.25) 0%, rgba(227, 0, 123, 0.25) 100%);
        border: 1px solid rgba(255, 107, 0, 0.4);
        color: #FFAA66;
        font-size: 0.8rem;
        font-weight: 800;
        letter-spacing: 0.06em;
        text-transform: uppercase;
        padding: 7px 18px;
        border-radius: 9999px;
        margin-bottom: 14px;
    }

    .login-title {
        font-size: 2.3rem;
        font-weight: 800;
        color: #FFFFFF;
        margin: 0 0 8px 0;
        letter-spacing: -0.02em;
    }

    .login-subtitle {
        font-size: 1.02rem;
        color: #E2D4DF;
        margin: 0;
        line-height: 1.5;
    }

    .login-form-card {
        background: rgba(45, 15, 37, 0.6);
        border-radius: 20px;
        padding: 30px;
        box-shadow: 0 12px 30px -5px rgba(0, 0, 0, 0.25);
        border: 1px solid rgba(255, 107, 0, 0.2);
        backdrop-filter: blur(12px);
    }

    .demo-credentials-card {
        background: rgba(35, 12, 28, 0.5);
        border-radius: 20px;
        padding: 26px;
        border: 1px solid rgba(227, 0, 123, 0.2);
    }

    .role-item {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: rgba(20, 8, 17, 0.6);
        border: 1px solid rgba(255, 107, 0, 0.15);
        padding: 13px 18px;
        border-radius: 14px;
        margin-bottom: 11px;
        transition: all 0.2s ease;
    }

    .role-item:hover {
        border-color: #FF6B00;
        box-shadow: 0 4px 14px rgba(255, 107, 0, 0.25);
    }

    .role-tag-gerente { 
        background: rgba(255, 107, 0, 0.2); 
        color: #FF8833; 
        border: 1px solid rgba(255, 107, 0, 0.4); 
        font-size: 0.74rem; 
        font-weight: 800; 
        padding: 4px 12px; 
        border-radius: 8px; 
    }
    .role-tag-lider { 
        background: rgba(227, 0, 123, 0.2); 
        color: #FF4D9E; 
        border: 1px solid rgba(227, 0, 123, 0.4); 
        font-size: 0.74rem; 
        font-weight: 800; 
        padding: 4px 12px; 
        border-radius: 8px; 
    }
    .role-tag-asesor { 
        background: rgba(245, 130, 32, 0.2); 
        color: #FFB84D; 
        border: 1px solid rgba(245, 130, 32, 0.4); 
        font-size: 0.74rem; 
        font-weight: 800; 
        padding: 4px 12px; 
        border-radius: 8px; 
    }
</style>
""", unsafe_allow_html=True)

# Funciones de formato
def formato_cop(val):
    num = limpiar_numero(val, 0.0)
    return f"${num:,.0f}".replace(",", ".")

# from frases import obtener_frase_motivacional_diaria

def formato_cop_signo(val):
    num = limpiar_numero(val, 0.0)
    if num == 0:
        return "$0"
    signo = "-" if num < 0 else ""
    return f"{signo}${abs(num):,.0f}".replace(",", ".")

def formato_porcentaje(val):
    if pd.isna(val):
        return "0.0%"
    try:
        if isinstance(val, str) and '%' in val:
            num = float(val.replace('%', '').strip())
        else:
            num = float(limpiar_numero(val, 0.0))
        if 0 < abs(num) <= 2.5:
            num = num * 100.0
        return f"{num:.1f}%"
    except Exception:
        return "0.0%"

def formato_saldo_entero(val):
    num = int(limpiar_numero(val, 0.0))
    return f"{num}"

def estilo_cumplimiento_facturacion(val):
    num = limpiar_numero(val, 0.0)
    if num >= 100.0:
        return 'background-color: #DCFCE7; color: #166534; font-weight: bold;'
    elif num >= 80.0:
        return 'background-color: #FEF9C3; color: #854D0E; font-weight: bold;'
    else:
        return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'

def estilo_cumplimiento_activas(val):
    num = limpiar_numero(val, 0.0)
    if num >= 100.0:
        return 'background-color: #DCFCE7; color: #166534; font-weight: bold;'
    elif num >= 80.0:
        return 'background-color: #FEF9C3; color: #854D0E; font-weight: bold;'
    else:
        return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'

def aplicar_estilo_styler(styler, func, subset):
    if hasattr(styler, 'map'):
        return styler.map(func, subset=subset)
    elif hasattr(styler, 'applymap'):
        return styler.applymap(func, subset=subset)
    return styler

# def renderizar_banner_motivacional(cumplimiento_pct, nombre_lider, codigo_grupo):
#     info = obtener_frase_motivacional_diaria(cumplimiento_pct, nombre_lider, codigo_grupo)
#     frase_txt = info['frase']
#     autor_txt = info['autor']
#     subtitulo = info['subtitulo']
#     cat = info['categoria']
#     
#     if cat == 'excelencia':
#         gradient = "linear-gradient(135deg, rgba(16, 185, 129, 0.22) 0%, rgba(15, 23, 42, 0.85) 100%)"
#         border_color = "rgba(16, 185, 129, 0.4)"
#         badge_bg = "rgba(16, 185, 129, 0.2)"
#         badge_color = "#34D399"
#         badge_txt = f"🌟 LIDERAZGO IMPARABLE • CUMPLIMIENTO {cumplimiento_pct:.1f}%"
#     elif cat == 'aceleracion':
#         gradient = "linear-gradient(135deg, rgba(245, 158, 11, 0.22) 0%, rgba(15, 23, 42, 0.85) 100%)"
#         border_color = "rgba(245, 158, 11, 0.4)"
#         badge_bg = "rgba(245, 158, 11, 0.2)"
#         badge_color = "#FBBF24"
#         badge_txt = f"🎯 ZONA DE ACELERACIÓN • CUMPLIMIENTO {cumplimiento_pct:.1f}%"
#     else:
#         gradient = "linear-gradient(135deg, rgba(139, 92, 246, 0.22) 0%, rgba(15, 23, 42, 0.85) 100%)"
#         border_color = "rgba(139, 92, 246, 0.4)"
#         badge_bg = "rgba(139, 92, 246, 0.2)"
#         badge_color = "#C084FC"
#         badge_txt = f"💪 TRANSFORMANDO CAMPAÑA • CUMPLIMIENTO {cumplimiento_pct:.1f}%"
#         
#     st.markdown(f"""
#     <style>
#     @keyframes desaparecerBanner {{
#         0% {{ opacity: 1; transform: translateY(0); max-height: 300px; margin-bottom: 22px; }}
#         85% {{ opacity: 1; transform: translateY(0); max-height: 300px; margin-bottom: 22px; }}
#         98% {{ opacity: 0; transform: translateY(-15px); max-height: 300px; margin-bottom: 22px; }}
#         100% {{ opacity: 0; transform: translateY(-20px); max-height: 0px; margin-bottom: 0px; padding-top: 0px; padding-bottom: 0px; border: none; overflow: hidden; visibility: hidden; }}
#     }}
#     .banner-animado-10s {{
#         animation: desaparecerBanner 10s cubic-bezier(0.4, 0, 0.2, 1) forwards;
#         overflow: hidden;
#     }}
#     </style>
#     <div class="banner-animado-10s" style="
#         background: {gradient};
#         border: 1px solid {border_color};
#         border-radius: 16px;
#         padding: 20px 24px;
#         margin-bottom: 22px;
#         backdrop-filter: blur(12px);
#         box-shadow: 0 10px 25px -5px rgba(0,0,0,0.3);
#     ">
#         <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px;">
#             <span style="background: {badge_bg}; color: {badge_color}; border: 1px solid {border_color}; font-size: 0.75rem; font-weight: 700; padding: 4px 14px; border-radius: 9999px; letter-spacing: 0.05em;">{badge_txt}</span>
#             <span style="color: #94A3B8; font-size: 0.82rem; font-weight: 500;">✨ Inspiración Diaria para {nombre_lider}</span>
#         </div>
#         <div style="font-size: 1.15rem; font-weight: 600; color: #F8FAFC; line-height: 1.5; font-style: italic; margin-bottom: 8px;">
#             "{frase_txt}"
#         </div>
#         <div style="display: flex; align-items: center; justify-content: space-between;">
#             <div style="color: #CBD5E1; font-size: 0.88rem; font-weight: 700;">— {autor_txt}</div>
#             <div style="color: #94A3B8; font-size: 0.83rem; font-weight: 500;">{subtitulo}</div>
#         </div>
#     </div>
#     """, unsafe_allow_html=True)

def renderizar_banner_cumpleanos(df_tableau, user_rol, user_nombre, user_grupo, user_sector, key_suffix="top"):
    """
    Renderiza el módulo y recordatorio de cumpleaños para las líderes y gerentes.
    Identifica de forma automática las asesoras de cumpleaños HOY, en los próximos 7 días
    y en el mes en curso, con botón de felicitación directa 1-clic por WhatsApp.
    """
    if df_tableau is None or df_tableau.empty:
        return

    plantilla_wa = st.session_state.get('plantilla_wa_cumpleanos', PLANTILLA_CUMPLEANOS_DEFAULT)
    data_cumple = obtener_cumpleanos_equipo(
        df_tableau,
        user_nombre=user_nombre,
        plantilla_wa=plantilla_wa
    )
    
    hoy_list = data_cumple['hoy']
    semana_list = data_cumple['semana']
    mes_list = data_cumple['mes']
    total_mes = data_cumple['total_mes']
    nombre_mes = data_cumple['nombre_mes']
    
    if total_mes == 0:
        return

    # Globos automáticos si hoy hay cumpleañeras (1 sola vez por sesión)
    if len(hoy_list) > 0 and 'balloons_cumple_shown' not in st.session_state:
        st.balloons()
        st.session_state['balloons_cumple_shown'] = True

    # Estilos según el estado de cumpleaños (Fondo de alto contraste y legibilidad total en temas Claro y Oscuro)
    if len(hoy_list) > 0:
        card_gradient = "linear-gradient(135deg, #C2185B 0%, #E3007B 50%, #FF6B00 100%)"
        border_color = "rgba(255, 215, 0, 0.8)"
        badge_bg = "rgba(0, 0, 0, 0.28)"
        badge_color = "#FFFFFF"
        badge_border = "1px solid rgba(255, 255, 255, 0.6)"
        badge_txt = f"🎉 ¡HOY CELEBRAMOS! • {len(hoy_list)} CUMPLEAÑERA{'S' if len(hoy_list) > 1 else ''}"
        icon_main = "🎂"
        expanded_default = False
    elif len(semana_list) > 0:
        card_gradient = "linear-gradient(135deg, #831843 0%, #9D174D 45%, #C026D3 100%)"
        border_color = "rgba(244, 114, 182, 0.7)"
        badge_bg = "linear-gradient(135deg, #FF6B00 0%, #E3007B 100%)"
        badge_color = "#FFFFFF"
        badge_border = "1px solid rgba(255, 255, 255, 0.45)"
        badge_txt = f"📅 PRÓXIMOS 7 DÍAS • {len(semana_list)} ASESORA{'S' if len(semana_list) > 1 else ''}"
        icon_main = "🎁"
        expanded_default = False
    else:
        card_gradient = "linear-gradient(135deg, #1E1B4B 0%, #312E81 50%, #4338CA 100%)"
        border_color = "rgba(129, 140, 248, 0.6)"
        badge_bg = "rgba(255, 255, 255, 0.18)"
        badge_color = "#FFFFFF"
        badge_border = "1px solid rgba(255, 255, 255, 0.35)"
    # Formato unificado del título del expander con todas las opciones y texto de acción
    equipo_info = f" — 🌸 Equipo de {user_nombre}" + (f" (Grupo {user_grupo})" if user_grupo else "")
    if len(hoy_list) > 0:
        titulo_expander = f"{icon_main} {badge_txt}  •  ✨ Ver Listado & Enviar Felicitaciones por WhatsApp ({total_mes} en {nombre_mes}){equipo_info}"
    elif len(semana_list) > 0:
        titulo_expander = f"{icon_main} {badge_txt}  •  ✨ Ver Listado & Enviar Felicitaciones por WhatsApp ({total_mes} en {nombre_mes}){equipo_info}"
    else:
        titulo_expander = f"{icon_main} {badge_txt}  •  ✨ Ver Listado de Cumpleaños ({total_mes} en {nombre_mes}){equipo_info}"

    # Estilización CSS del Expander Unificado de Cumpleaños
    st.markdown(f"""
    <style>
    div.stElementContainer:has(.cumple-banner-unified) + div.stElementContainer [data-testid="stExpander"],
    div:has(.cumple-banner-unified) + div [data-testid="stExpander"] {{
        border: 2px solid {border_color} !important;
        border-radius: 14px !important;
        background: rgba(15, 23, 42, 0.02) !important;
        box-shadow: 0 6px 20px -4px rgba(0,0,0,0.22) !important;
        margin-bottom: 14px !important;
        overflow: hidden !important;
    }}
    div.stElementContainer:has(.cumple-banner-unified) + div.stElementContainer [data-testid="stExpander"] > details,
    div:has(.cumple-banner-unified) + div [data-testid="stExpander"] > details {{
        border: none !important;
        background: transparent !important;
    }}
    div.stElementContainer:has(.cumple-banner-unified) + div.stElementContainer [data-testid="stExpander"] > details > summary,
    div:has(.cumple-banner-unified) + div [data-testid="stExpander"] > details > summary {{
        background: {card_gradient} !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 10px 18px !important;
        color: #FFFFFF !important;
        font-weight: 800 !important;
        cursor: pointer !important;
        transition: all 0.25s ease !important;
    }}
    div.stElementContainer:has(.cumple-banner-unified) + div.stElementContainer [data-testid="stExpander"] > details[open] > summary,
    div:has(.cumple-banner-unified) + div [data-testid="stExpander"] > details[open] > summary {{
        border-bottom-left-radius: 0px !important;
        border-bottom-right-radius: 0px !important;
        border-bottom: 1px solid rgba(255,255,255,0.25) !important;
    }}
    div.stElementContainer:has(.cumple-banner-unified) + div.stElementContainer [data-testid="stExpander"] > details > summary:hover,
    div:has(.cumple-banner-unified) + div [data-testid="stExpander"] > details > summary:hover {{
        filter: brightness(1.08) !important;
        box-shadow: inset 0 0 10px rgba(255,255,255,0.2) !important;
    }}
    div.stElementContainer:has(.cumple-banner-unified) + div.stElementContainer [data-testid="stExpander"] > details > summary p,
    div.stElementContainer:has(.cumple-banner-unified) + div.stElementContainer [data-testid="stExpander"] > details > summary span,
    div:has(.cumple-banner-unified) + div [data-testid="stExpander"] > details > summary p,
    div:has(.cumple-banner-unified) + div [data-testid="stExpander"] > details > summary span {{
        color: #FFFFFF !important;
        font-weight: 800 !important;
        font-size: 0.95rem !important;
        letter-spacing: 0.01em !important;
    }}
    div.stElementContainer:has(.cumple-banner-unified) + div.stElementContainer [data-testid="stExpander"] > details > summary svg,
    div:has(.cumple-banner-unified) + div [data-testid="stExpander"] > details > summary svg {{
        fill: #FFFFFF !important;
        color: #FFFFFF !important;
        width: 1.25rem !important;
        height: 1.25rem !important;
    }}
    div.stElementContainer:has(.cumple-banner-unified) + div.stElementContainer [data-testid="stExpander"] > details > div[data-testid="stExpanderDetails"],
    div:has(.cumple-banner-unified) + div [data-testid="stExpander"] > details > div[data-testid="stExpanderDetails"] {{
        padding: 14px 16px !important;
        background: transparent !important;
    }}
    </style>
    <div class="cumple-banner-unified"></div>
    """, unsafe_allow_html=True)

    with st.expander(titulo_expander, expanded=False):
        tab_c_hoy, tab_c_sem, tab_c_mes, tab_c_edit = st.tabs([
            f"🎈 Hoy ({len(hoy_list)})",
            f"📅 Próximos 7 Días ({len(semana_list)})",
            f"🗓️ Todo el Mes ({total_mes})",
            "✍️ Personalizar Felicitación"
        ])
        
        def _render_cards_cumple(items_list, es_hoy=False):
            if not items_list:
                st.info("🌸 No hay cumpleaños en este rango actualmente.")
                return
                
            cols_grid = st.columns(2 if len(items_list) > 1 else 1)
            for idx, item in enumerate(items_list):
                with cols_grid[idx % 2]:
                    nivel_style = color_nivel(item['nivel'])
                    tag_t = item.get('etiqueta_tiempo', f"Día {item['dia']}")
                    bg_t = "#10B981" if es_hoy else ("#F59E0B" if item.get('dias_falta', 99) <= 2 else "#6366F1")
                    
                    # Tarjeta compacta visual con alto contraste garantizado
                    st.markdown(f"""
<div style="background: #1E293B; border: 1px solid rgba(227, 0, 123, 0.45); border-radius: 12px; padding: 10px 14px 8px 14px; margin-bottom: 6px; box-shadow: 0 4px 12px rgba(0,0,0,0.2);">
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
        <span style="font-size: 1rem; font-weight: 700; color: #FFFFFF;">🌸 {item['nombre']}</span>
        <span style="background: {bg_t}; color: #FFFFFF; font-size: 0.72rem; font-weight: 800; padding: 2px 9px; border-radius: 9999px;">{tag_t}</span>
    </div>
    <div style="display: flex; gap: 8px; align-items: center; flex-wrap: wrap; font-size: 0.78rem; color: #E2E8F0;">
        <span style="{nivel_style} padding: 1px 7px; border-radius: 4px; font-weight: 700;">{item['nivel']}</span>
        <span>CB: <b style='color:#FFFFFF;'>{item['codigo_cb']}</b></span>
        <span>• Grupo: <b style='color:#FFFFFF;'>{item['grupo']}</b></span>
        <span>• <span style='color:#CBD5E1;'>{item['sit_comercial']}</span></span>
    </div>
</div>
""", unsafe_allow_html=True)
                    
                    # Botón nativo de WhatsApp 1-clic directo
                    if item.get('link_wa'):
                        st.link_button(
                            f"📲 Felicitar a {item['primer_nombre']} por WhatsApp",
                            url=item['link_wa'],
                            use_container_width=True
                        )
                    else:
                        st.caption("📵 *Sin número de celular registrado*")
                    st.write("")
                    
        with tab_c_hoy:
            if hoy_list:
                st.success(f"🎂 **¡Hoy tenemos {len(hoy_list)} cumpleañera{'s' if len(hoy_list) > 1 else ''} en tu equipo!** Toca el botón para enviarles el mensaje personalizado por WhatsApp:")
                _render_cards_cumple(hoy_list, es_hoy=True)
                col_b1, col_b2 = st.columns([2, 2])
                with col_b1:
                    if st.button("🎈 Volver a lanzar globos", key=f"btn_globos_tab_{key_suffix}"):
                        st.balloons()
            else:
                st.info("🌸 Hoy no hay cumpleaños en tu equipo. ¡Revisa la pestaña de los **Próximos 7 Días** para prepararte!")
                
        with tab_c_sem:
            if semana_list:
                st.markdown(f"###### 📅 Cumpleaños en los próximos 7 días ({len(semana_list)} asesoras):")
                _render_cards_cumple(semana_list, es_hoy=False)
            else:
                st.info("📅 No hay cumpleaños registrados en los próximos 7 días.")
                
        with tab_c_mes:
            if mes_list:
                st.markdown(f"###### 🗓️ Todas las cumpleañeras del mes de {nombre_mes} ({len(mes_list)} en total):")
                df_mes_vista = pd.DataFrame([
                    {
                        'Día': f"{it['dia']} de {nombre_mes}",
                        'Asesora': it['nombre'],
                        'Nivel': it['nivel'],
                        'Grupo': it['grupo'],
                        'Código CB': it['codigo_cb'],
                        'Celular': it['celular'] if it['celular'] else 'Sin Celular',
                        'Situación': it['sit_comercial'],
                        'Enlace WhatsApp': it['link_wa']
                    }
                    for it in mes_list
                ])
                st.dataframe(
                    df_mes_vista,
                    column_config={
                        "Enlace WhatsApp": st.column_config.LinkColumn(
                            "📲 Chat WhatsApp",
                            help="Clic para abrir WhatsApp con el mensaje de cumpleaños",
                            display_text="📲 Enviar WA"
                        )
                    },
                    use_container_width=True,
                    hide_index=True
                )
                csv_cumple = df_mes_vista.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label=f"📥 Descargar Lista de Cumpleaños de {nombre_mes} (CSV)",
                    data=csv_cumple,
                    file_name=f"Cumpleaños_{nombre_mes}_{user_grupo if user_grupo else 'Sector'}.csv",
                    mime="text/csv",
                    key=f"btn_descargar_cumple_mes_{key_suffix}"
                )
            else:
                st.info(f"No hay registros de cumpleaños para {nombre_mes}.")
                
        with tab_c_edit:
            st.markdown("###### ✍️ Personalizar Mensaje Predeterminado de WhatsApp")
            st.caption("Modifica el mensaje que se enviará automáticamente. Etiquetas disponibles: `{primer_nombre}`, `{nombre}`, `{nivel}`, `{lider}`.")
            
            nueva_plantilla = st.text_area(
                "Plantilla de Felicitación:",
                value=plantilla_wa,
                height=130,
                key=f"txt_plantilla_wa_cumple_{key_suffix}"
            )
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                if st.button("💾 Guardar Plantilla", key=f"btn_guardar_plantilla_cumple_{key_suffix}"):
                    st.session_state['plantilla_wa_cumpleanos'] = nueva_plantilla
                    st.success("✅ Plantilla actualizada para esta sesión.")
                    st.rerun()
            with col_p2:
                if st.button("🔄 Restaurar Predeterminada", key=f"btn_reset_plantilla_cumple_{key_suffix}"):
                    st.session_state['plantilla_wa_cumpleanos'] = PLANTILLA_CUMPLEANOS_DEFAULT
                    st.success("✅ Mensaje restaurado al original.")
                    st.rerun()

# --- COMPONENTES GRÁFICOS INTERACTIVOS (LA JOYA DEL PASTEL) ---
def crear_tacometro_360(titulo, valor_pct, meta_val, real_val):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=min(float(valor_pct), 150.0),
        number={'suffix': '%', 'font': {'color': '#F8FAFC', 'size': 32}},
        title={'text': f"<b>{titulo}</b><br><span style='font-size:0.8em;color:#94A3B8'>Real: {real_val} | Obj: {meta_val}</span>", 'font': {'color': '#F8FAFC', 'size': 14}},
        gauge={
            'axis': {'range': [0, 150], 'tickwidth': 1, 'tickcolor': "#475569"},
            'bar': {'color': "#3B82F6", 'thickness': 0.25},
            'bgcolor': "rgba(15, 23, 42, 0.5)",
            'borderwidth': 1,
            'bordercolor': "rgba(255, 255, 255, 0.1)",
            'steps': [
                {'range': [0, 85], 'color': 'rgba(239, 68, 68, 0.3)'},
                {'range': [85, 99.9], 'color': 'rgba(245, 158, 11, 0.3)'},
                {'range': [99.9, 150], 'color': 'rgba(16, 185, 129, 0.3)'}
            ],
            'threshold': {
                'line': {'color': "#10B981", 'width': 4},
                'thickness': 0.75,
                'value': 100
            }
        }
    ))
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=230,
        margin=dict(l=20, r=20, t=50, b=20)
    )
    return fig

def crear_ranking_lideres_fig(df_metas):
    if df_metas is None or df_metas.empty:
        return None
    col_nom = 'Nombre de consultora' if 'Nombre de consultora' in df_metas.columns else 'Nombre Consultora'
    col_cump = 'Cumplimiento Facturación' if 'Cumplimiento Facturación' in df_metas.columns else 'cump_facturacion'
    
    df_rank = df_metas.copy()
    if col_cump not in df_rank.columns or col_nom not in df_rank.columns:
        return None
    
    df_rank['Cumplimiento_Pct'] = df_rank[col_cump].apply(lambda v: limpiar_numero(v, 0.0))
    
    def _limpiar_nombre_lider_solo(val):
        s = str(val).strip()
        if ' - ' in s:
            s = s.split(' - ', 1)[1].strip()
        if s.lower().startswith('grupo '):
            partes = s.split(' ', 2)
            if len(partes) > 2:
                s = partes[2].strip()
        return s[:30]

    df_rank['Etiqueta_Lider'] = df_rank[col_nom].apply(_limpiar_nombre_lider_solo)
    df_rank = df_rank[~df_rank['Etiqueta_Lider'].str.lower().isin(['nan', 'none', '-', ''])]
    df_rank = df_rank.sort_values(by='Cumplimiento_Pct', ascending=True).tail(12)
    
    fig = px.bar(
        df_rank,
        x='Cumplimiento_Pct',
        y='Etiqueta_Lider',
        orientation='h',
        title="<b>🏆 Ranking de Cumplimiento de Facturación por Líder</b>",
        labels={'Cumplimiento_Pct': '% Cumplimiento', 'Etiqueta_Lider': 'Líder de Negocio'},
        color='Cumplimiento_Pct',
        color_continuous_scale=[
            [0.0, '#EF4444'],
            [0.70, '#F59E0B'],
            [1.0, '#10B981']
        ]
    )
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#F8FAFC', family='Outfit'),
        height=360,
        margin=dict(l=10, r=20, t=40, b=30),
        coloraxis_showscale=False
    )
    fig.update_traces(hovertemplate="<b>%{y}</b><br>Cumplimiento: %{x:.1f}%<extra></extra>")
    return fig

def crear_dona_cartera_fig(df_tableau):
    if df_tableau is None or df_tableau.empty or 'Deuda Mora' not in df_tableau.columns:
        return None
    
    mora_0 = len(df_tableau[df_tableau['Deuda Mora'] <= 0])
    mora_leve = len(df_tableau[(df_tableau['Deuda Mora'] > 0) & (df_tableau['Deuda Mora'] <= 200000)])
    mora_media = len(df_tableau[(df_tableau['Deuda Mora'] > 200000) & (df_tableau['Deuda Mora'] <= 500000)])
    mora_critica = len(df_tableau[df_tableau['Deuda Mora'] > 500000])
    
    df_pie = pd.DataFrame([
        {'Tramo': '🟢 Al Día ($0)', 'Consultoras': mora_0},
        {'Tramo': '🟡 Mora Leve ($1 - $200k)', 'Consultoras': mora_leve},
        {'Tramo': '🟠 Mora Media ($200k - $500k)', 'Consultoras': mora_media},
        {'Tramo': '🔴 Mora Crítica (> $500k)', 'Consultoras': mora_critica}
    ])
    df_pie = df_pie[df_pie['Consultoras'] > 0]
    
    fig = px.pie(
        df_pie,
        values='Consultoras',
        names='Tramo',
        hole=0.55,
        title="<b>⚠️ Distribución de Cartera & Semáforo de Mora</b>",
        color='Tramo',
        color_discrete_map={
            '🟢 Al Día ($0)': '#10B981',
            '🟡 Mora Leve ($1 - $200k)': '#F59E0B',
            '🟠 Mora Media ($200k - $500k)': '#F97316',
            '🔴 Mora Crítica (> $500k)': '#EF4444'
        }
    )
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#F8FAFC', family='Outfit'),
        height=360,
        margin=dict(l=20, r=20, t=40, b=20)
    )
    fig.update_traces(textposition='inside', textinfo='percent+label')
    return fig

def crear_embudo_red_fig(disponibles, inicios, reinicios, activas):
    fig = go.Figure(go.Funnel(
        y=['Disponibles Totales', 'Inicios Nuevos', 'Reinicios', 'Consultoras Activas'],
        x=[disponibles, inicios, reinicios, activas],
        textinfo="value+percent initial",
        marker={"color": ["#3B82F6", "#8B5CF6", "#EC4899", "#10B981"]}
    ))
    fig.update_layout(
        title="<b>🚀 Embudo de Conversión & Retención de Red</b>",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#F8FAFC', family='Outfit'),
        height=360,
        margin=dict(l=20, r=20, t=40, b=20)
    )
    return fig

def crear_treemap_red_fig(df_tableau):
    if df_tableau is None or df_tableau.empty or 'Color' not in df_tableau.columns:
        return None
    
    df_tree = df_tableau.copy()
    col_color = 'Color' if 'Color' in df_tree.columns else 'Nivel / Color'
    col_sit = 'Sit. Comercial' if 'Sit. Comercial' in df_tree.columns else 'Situación'
    
    fig = px.treemap(
        df_tree,
        path=[col_color, col_sit],
        values='Pts Acum' if 'Pts Acum' in df_tree.columns else None,
        title="<b>💎 Mapa Térmico de Red por Nivel de Crecimiento & Estado</b>",
        color=col_color,
        color_discrete_map={
            'Bronce': '#D97706',
            'Plata': '#94A3B8',
            'Oro': '#EAB308',
            'Platino': '#38BDF8',
            'Zafiro': '#3B82F6',
            'Diamante': '#A855F7'
        }
    )
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#F8FAFC', family='Outfit'),
        height=380,
        margin=dict(l=10, r=10, t=40, b=10)
    )
    return fig

def crear_scatter_atencion_fig(df_tableau):
    if df_tableau is None or df_tableau.empty or 'Deuda Mora' not in df_tableau.columns:
        return None
    
    df_scat = df_tableau.copy()
    fig = px.scatter(
        df_scat,
        x='Pts Acum' if 'Pts Acum' in df_scat.columns else 'Pts Natura',
        y='Deuda Mora',
        color='Sit. Comercial' if 'Sit. Comercial' in df_scat.columns else None,
        hover_name='Nombre' if 'Nombre' in df_scat.columns else 'Código CB',
        size='Deuda Total' if 'Deuda Total' in df_scat.columns else None,
        title="<b>📲 Matriz de Atención Prioritaria (Puntos vs Deuda Mora)</b>",
        labels={'Pts Acum': 'Puntos Acumulados', 'Deuda Mora': 'Deuda Mora ($)'}
    )
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#F8FAFC', family='Outfit'),
        height=380,
        margin=dict(l=10, r=10, t=40, b=10)
    )
    return fig

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

# --- CONTROL DE SESIÓN, INACTIVIDAD (15 MIN) Y LOGIN ---
TIEMPO_INACTIVIDAD_SEGUNDOS = 15 * 60  # 15 minutos (900 segundos)

if 'user' not in st.session_state:
    st.session_state['user'] = None

if 'ultimo_acceso' not in st.session_state:
    st.session_state['ultimo_acceso'] = time.time()

# 1. Validar si la sesión activa superó el tiempo máximo de inactividad
if st.session_state['user'] is not None:
    tiempo_inactivo = time.time() - st.session_state.get('ultimo_acceso', time.time())
    if tiempo_inactivo > TIEMPO_INACTIVIDAD_SEGUNDOS:
        st.session_state['user'] = None
        if 'user' in st.query_params:
            del st.query_params['user']
        st.session_state['msg_timeout'] = "⏳ Tu sesión ha expirado automáticamente por inactividad (más de 15 minutos sin interacción). Por tu seguridad y privacidad de los datos, por favor inicia sesión nuevamente."
        st.session_state['ultimo_acceso'] = time.time()
        st.rerun()
    else:
        # Renovar el temporizador en cada interacción activa
        st.session_state['ultimo_acceso'] = time.time()

# 2. Recuperación automática de sesión al refrescar la página (F5) SOLO si no expiró
if st.session_state['user'] is None:
    session_user_param = st.query_params.get('user')
    if session_user_param and not st.session_state.get('msg_timeout'):
        todos_usuarios = cargar_usuarios()
        u_clean_param = str(session_user_param).strip().lower()
        if u_clean_param in todos_usuarios:
            user_info = todos_usuarios[u_clean_param].copy()
            user_info['username'] = u_clean_param
            st.session_state['user'] = user_info
            st.session_state['ultimo_acceso'] = time.time()

if st.session_state['user'] is None:
    st.markdown("""
    <div class="login-container">
        <div class="login-hero-card">
            <div class="login-badge">✨ SISTEMA DE GESTIÓN Y LIDERAZGO EMPRESARIAL</div>
            <h1 class="login-title">Portal de Acceso Corporativo</h1>
            <p class="login-subtitle">Gestión estratégica de metas de ciclo, indicadores de facturación y seguimiento privado por Líder de Negocio.</p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    col_pad1, col_center, col_pad2 = st.columns([0.5, 2, 0.5])
    
    with col_center:
        if st.session_state.get('msg_timeout'):
            st.warning(st.session_state['msg_timeout'])

        tab_login_tab, tab_recuperar_tab, tab_registro_tab = st.tabs(["🔑 Iniciar Sesión", "🆘 ¿Olvidaste tu Usuario o Clave?", "🚀 Probar Gratis (15 Días)"])
        
        with tab_login_tab:
            st.markdown('<div class="login-form-card">', unsafe_allow_html=True)
            st.markdown("#### 🔑 Ingreso al Sistema")
            st.caption("Ingresa con tu correo, usuario institucional, o código de líder.")
            
            with st.form("form_login_modern"):
                input_user = st.text_input("👤 Usuario o Correo", value="", placeholder="Ingresa tu correo, usuario o código...")
                input_pass = st.text_input("🔒 Contraseña", type="password", value="", placeholder="••••••••")
                btn_login = st.form_submit_button("🚀 Entrar al Sistema", type="primary", use_container_width=True)
                
                if btn_login:
                    user_auth = autenticar_usuario(input_user, input_pass)
                    if user_auth:
                        st.session_state['user'] = user_auth
                        st.session_state['ultimo_acceso'] = time.time()
                        if 'msg_timeout' in st.session_state:
                            del st.session_state['msg_timeout']
                        st.query_params['user'] = user_auth.get('username', input_user)
                        st.success(f"¡Bienvenido(a), {user_auth['nombre']}!")
                        st.rerun()
                    else:
                        st.error("❌ Credenciales incorrectas. Si no recuerdas tu clave o usuario, puedes recuperarla en la pestaña '🆘 ¿Olvidaste tu Usuario o Clave?'.")
            st.markdown('</div>', unsafe_allow_html=True)
            
        with tab_recuperar_tab:
            st.markdown('<div class="login-form-card">', unsafe_allow_html=True)
            st.markdown("#### 🆘 Recuperación de Acceso y Contraseñas")
            st.caption("Ingresa tu correo, nombre, código de grupo (Líderes) o código de sector (Gerentes) para verificar tu cuenta y restaurar tu clave al instante.")
            
            with st.form("form_recuperar_acceso_step1"):
                in_recuperar_id = st.text_input(
                    "🔍 Identificador de Cuenta:",
                    placeholder="ej: dolly.parra@natura.net, 9334, 700000466, o lider9640",
                    key="in_recuperar_id_input"
                )
                btn_verificar_cuenta = st.form_submit_button("🔍 Verificar y Buscar mi Cuenta", type="primary", use_container_width=True)
                
                if btn_verificar_cuenta:
                    if not in_recuperar_id.strip():
                        st.warning("⚠️ Por favor escribe tu correo, usuario o código para buscar.")
                    else:
                        ok_find, data_find, msg_find = buscar_cuenta_usuario(in_recuperar_id)
                        if ok_find:
                            st.session_state['cuenta_recuperada_temp'] = data_find
                            import random
                            pin_gen = str(random.randint(100000, 999999))
                            st.session_state['pin_recuperacion'] = pin_gen
                            st.success("✅ ¡Cuenta encontrada exitosamente!")
                            st.rerun()
                        else:
                            st.error(f"❌ {msg_find}")
                            
            cuenta_rec = st.session_state.get('cuenta_recuperada_temp')
            if cuenta_rec:
                u_rec_name = cuenta_rec['username']
                u_rec_nombre = cuenta_rec.get('nombre', '')
                u_rec_rol = cuenta_rec.get('rol', '').capitalize()
                u_rec_sector = cuenta_rec.get('nombre_sector', '') or cuenta_rec.get('codigo_sector', '')
                u_rec_grp = cuenta_rec.get('codigo_grupo', '')
                
                st.markdown("---")
                st.markdown(f"""
                <div style="background: rgba(16, 185, 129, 0.12); border: 1.5px solid #10B981; border-radius: 12px; padding: 14px 18px; margin-bottom: 15px;">
                    <div style="font-size: 1rem; font-weight: 700; color: #10B981; margin-bottom: 6px;">
                        🌸 Verificación Exitosa de Cuenta
                    </div>
                    <div style="font-size: 0.88rem; color: #E2E8F0; line-height: 1.6;">
                        • <b>Titular:</b> {u_rec_nombre}<br>
                        • <b>Rol:</b> {u_rec_rol} {f'• Grupo {u_rec_grp}' if u_rec_grp else ''}<br>
                        • <b>Sector:</b> {u_rec_sector}<br>
                        • 👤 <b>Tu Usuario Oficial de Acceso es:</b> <span style="background: rgba(255,255,255,0.25); padding: 2px 8px; border-radius: 4px; font-weight: 800; color: #FFFFFF;">{u_rec_name}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                tab_rec_wa, tab_rec_pin = st.tabs(["📲 Restablecer & Enviar por WhatsApp", "🔑 Cambiar Contraseña Ahora"])
                
                with tab_rec_wa:
                    st.caption("Restablece tu contraseña a una clave temporal y recíbela directamente en tu WhatsApp:")
                    nueva_clave_sug = "lider123" if cuenta_rec.get('rol') == 'lider' else "admin123"
                    
                    with st.form("form_rec_wa_exec"):
                        tel_rec_in = st.text_input("Ingresa tu número celular (10 dígitos):", placeholder="ej. 3123456789", key="tel_rec_in_wa")
                        btn_reset_wa = st.form_submit_button("🔄 Restablecer a Clave Temporal & Generar WhatsApp", type="primary", use_container_width=True)
                        
                        if btn_reset_wa:
                            ok_rst, msg_rst = restablecer_password_usuario(u_rec_name, nueva_clave_sug, debe_cambiar=True)
                            if ok_rst:
                                st.session_state['ultimo_reset_publico'] = {
                                    'usuario': u_rec_name,
                                    'nombre': u_rec_nombre,
                                    'password': nueva_clave_sug,
                                    'telefono': tel_rec_in
                                }
                                st.success(f"✅ ¡Contraseña restablecida exitosamente a: `{nueva_clave_sug}`!")
                                st.rerun()
                            else:
                                st.error(f"❌ {msg_rst}")
                                
                    ult_pub = st.session_state.get('ultimo_reset_publico')
                    if ult_pub and ult_pub.get('usuario') == u_rec_name:
                        msg_wa_recup = (
                            f"🔐 *RECUPERACIÓN DE ACCESO - SISTEMA NATURA & AVON*\n\n"
                            f"🌸 ¡Hola {ult_pub['nombre'].split()[0].title()}!\n"
                            f"Tus datos de ingreso han sido verificados:\n\n"
                            f"👤 *Usuario:* `{ult_pub['usuario']}`\n"
                            f"🔑 *Contraseña Temporal:* `{ult_pub['password']}`\n\n"
                            f"🌐 *Ingresa aquí:* https://metaseindicadores.up.railway.app\n\n"
                            f"✨ *Nota:* Al iniciar sesión se te pedirá actualizar tu contraseña personal."
                        )
                        
                        st.text_area("Mensaje de recuperación listo:", msg_wa_recup, height=140, key="txt_msg_rec_wa")
                        
                        tel_final = ult_pub.get('telefono', '').strip()
                        if tel_final and len(tel_final) >= 10:
                            link_wa_rec = f"https://api.whatsapp.com/send?phone=57{tel_final}&text={urllib.parse.quote(msg_wa_recup)}"
                            st.link_button("📲 Abrir WhatsApp y Recibir mis Credenciales", url=link_wa_rec, use_container_width=True)
                        else:
                            st.caption("💡 Puedes copiar el mensaje anterior o ingresar tu número celular arriba para abrir WhatsApp directamente.")

                with tab_rec_pin:
                    st.caption("Crea una nueva contraseña inmediatamente con el PIN de verificación:")
                    pin_esperado = st.session_state.get('pin_recuperacion', '123456')
                    
                    with st.form("form_rec_directo_nueva_pass"):
                        st.info(f"💡 Tu código PIN de seguridad de verificación es: **{pin_esperado}**")
                        pin_ingresado = st.text_input("Ingresa el PIN de seguridad:", placeholder="6 dígitos", key="pin_in_rec")
                        nueva_pass_1 = st.text_input("Nueva Contraseña:", type="password", placeholder="Mínimo 6 caracteres", key="nueva_pass_1_in")
                        nueva_pass_2 = st.text_input("Confirma la Nueva Contraseña:", type="password", placeholder="Repite tu contraseña", key="nueva_pass_2_in")
                        
                        btn_cambiar_pass_rec = st.form_submit_button("🔒 Guardar Nueva Contraseña e Iniciar Sesión", type="primary", use_container_width=True)
                        
                        if btn_cambiar_pass_rec:
                            if pin_ingresado.strip() != pin_esperado.strip():
                                st.error("❌ El PIN de seguridad ingresado no es correcto.")
                            elif len(nueva_pass_1) < 6:
                                st.warning("⚠️ La contraseña debe tener al menos 6 caracteres.")
                            elif nueva_pass_1 != nueva_pass_2:
                                st.error("❌ Las contraseñas no coinciden.")
                            else:
                                ok_ch, msg_ch = cambiar_password_usuario(u_rec_name, nueva_pass_1)
                                if ok_ch:
                                    st.success(f"🎉 ¡Contraseña actualizada exitosamente! Iniciando sesión...")
                                    user_authed = autenticar_usuario(u_rec_name, nueva_pass_1)
                                    if user_authed:
                                        st.session_state['user'] = user_authed
                                        st.session_state['ultimo_acceso'] = time.time()
                                        if 'cuenta_recuperada_temp' in st.session_state:
                                            del st.session_state['cuenta_recuperada_temp']
                                        st.query_params['user'] = user_authed.get('username', u_rec_name)
                                        st.rerun()
                                else:
                                    st.error(f"❌ {msg_ch}")
                                    
                st.markdown("---")
                st.markdown("###### 💬 ¿Necesitas asistencia personalizada?")
                msg_soporte = f"Hola Soporte Técnico, soy {u_rec_nombre} del Sector {u_rec_sector}. Necesito ayuda para recuperar mi acceso al usuario {u_rec_name}."
                url_soporte_wa = f"https://api.whatsapp.com/send?phone=573057939537&text={urllib.parse.quote(msg_soporte)}"
                st.link_button("👩‍💻 Contactar a Soporte por WhatsApp (3057939537)", url=url_soporte_wa, use_container_width=True)

            st.markdown('</div>', unsafe_allow_html=True)
            
        with tab_registro_tab:
            st.markdown('<div class="login-form-card">', unsafe_allow_html=True)
            st.markdown("#### 🚀 Registro para Gerentes de Sector")
            st.caption("Activa tu prueba gratuita de 15 días con acceso total para ti y todo tu equipo de líderes.")

            catalogo_sectores = cargar_catalogo_sectores()
            opciones_sector = ["-- Selecciona tu Sector --"]
            mapa_opciones_sector = {}

            if catalogo_sectores:
                for cod, info in sorted(catalogo_sectores.items(), key=lambda x: str(x[1].get('nombre_sector', ''))):
                    nom_sec = limpiar_nombre_sector_solo(info.get('nombre_sector', f'SECTOR {cod}'))
                    label = f"{nom_sec}"
                    if label not in mapa_opciones_sector:
                        opciones_sector.append(label)
                        mapa_opciones_sector[label] = info

            opciones_sector.append("✏️ Otro Sector (Ingresar manualmente)")

            sec_seleccionado_label = st.selectbox(
                "🏢 Sector:",
                options=opciones_sector,
                key="sel_sector_registro"
            )

            info_sec_sel = mapa_opciones_sector.get(sec_seleccionado_label)

            default_cod_sec = info_sec_sel.get('codigo_sector', '') if info_sec_sel else ""
            default_nom_sec = limpiar_nombre_sector_solo(info_sec_sel.get('nombre_sector', '')) if info_sec_sel else ""

            es_manual = (sec_seleccionado_label == "✏️ Otro Sector (Ingresar manualmente)")

            with st.form("form_registro_gerente"):
                reg_nombre = st.text_input("👩‍💼 Nombre Completo", placeholder="Ej: Dolly Parra o Clery Cuellar")
                reg_correo = st.text_input("✉️ Correo Electrónico (Será tu usuario de acceso)", placeholder="ejemplo@gmail.com")
                reg_pass = st.text_input("🔒 Contraseña Segura", type="password", placeholder="Mínimo 6 caracteres")
                reg_tel = st.text_input("📲 Celular / WhatsApp de Contacto", placeholder="Ej: 3113201145")
                
                if es_manual:
                    reg_sec_cod = st.text_input(
                        "🏷️ Código Oficial de Sector",
                        type="password",
                        placeholder="•••••••••",
                        help="Código único de sector"
                    )
                    reg_sec_nom = st.text_input(
                        "🏢 Nombre del Sector",
                        placeholder="Ej: SECTOR COLORES o SECTOR MATICES"
                    )
                else:
                    reg_sec_cod = default_cod_sec
                    reg_sec_nom = default_nom_sec
                
                btn_registro = st.form_submit_button("🎉 Comenzar Mi Prueba Gratis de 15 Días", type="primary", use_container_width=True)
                
                if btn_registro:
                    if not reg_sec_cod or not reg_sec_nom:
                        st.error("Por favor selecciona tu Sector de la lista antes de continuar.")
                    else:
                        ok_reg, msg_reg, u_data = registrar_nueva_gerente(
                            nombre=reg_nombre,
                            correo=reg_correo,
                            password=reg_pass,
                            telefono=reg_tel,
                            cod_sector=reg_sec_cod,
                            nombre_sector=reg_sec_nom
                        )
                        if ok_reg:
                            st.session_state['user'] = u_data
                            st.query_params['user'] = u_data.get('username', reg_correo)
                            st.success("✅ " + msg_reg)
                            st.rerun()
                        else:
                            st.error(msg_reg)
            st.markdown('</div>', unsafe_allow_html=True)
            
    st.stop()

# Usuario logueado activo
current_user = st.session_state.get('user') or {}
user_nombre = current_user.get('nombre', 'Usuario')
user_rol = current_user.get('rol', 'asesor')
user_grupo = str(current_user.get('codigo_grupo', '')).strip() if current_user.get('codigo_grupo') else ""
user_sector = str(current_user.get('codigo_sector', '')).strip() if current_user.get('codigo_sector') else ""
user_sector_nombre = obtener_nombre_sector_usuario(current_user)

# Verificación de Suscripción / Modo Prueba / Bloqueo
info_suscripcion = verificar_estado_suscripcion(current_user)

if not info_suscripcion.get("permitido", True):
    col_bl1, col_bl_center, col_bl2 = st.columns([0.5, 2, 0.5])
    with col_bl_center:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, rgba(239, 68, 68, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%);
                    border: 1px solid rgba(239, 68, 68, 0.4); border-radius: 20px; padding: 32px; text-align: center;
                    box-shadow: 0 20px 40px -10px rgba(0,0,0,0.5);">
            <div style="font-size: 3.5rem; margin-bottom: 12px;">🔒</div>
            <h2 style="color: #F8FAFC; margin-bottom: 8px;">Periodo de Prueba Finalizado</h2>
            <p style="color: #CBD5E1; font-size: 1.05rem; line-height: 1.6; margin-bottom: 20px;">
                {info_suscripcion.get("motivo", "Tu acceso temporal de prueba de 15 días ha concluido.")}
            </p>
            <div style="background: rgba(30, 41, 59, 0.8); border: 1px solid rgba(255,255,255,0.1); border-radius: 12px; padding: 18px; margin-bottom: 24px; text-align: left;">
                <p style="color: #10B981; font-weight: 700; margin-bottom: 6px;">🛡️ Tu información y notas están 100% a salvo</p>
                <p style="color: #94A3B8; font-size: 0.88rem; margin-bottom: 0;">
                    Toda la base de datos de consultoras, metas de ciclo y notas de gestión registradas por tus líderes se encuentran guardadas y protegidas. Al activar tu suscripción, recuperarás acceso inmediato a todo tu historial sin necesidad de volver a cargar ni reconfigurar nada.
                </p>
            </div>
            <p style="color: #F8FAFC; font-weight: 600; margin-bottom: 16px;">Para activar tu suscripción y desbloquear el acceso para ti y tu equipo:</p>
            <a href="https://wa.me/573057939537?text=Hola,%20deseo%20activar%20mi%20suscripción%20para%20el%20Sector%20{user_sector}" target="_blank" style="text-decoration: none;">
                <button style="background: linear-gradient(135deg, #25D366 0%, #128C7E 100%); color: white; border: none; padding: 14px 28px; border-radius: 12px; font-size: 1.05rem; font-weight: 700; cursor: pointer; width: 100%;">
                    💬 Contactar por WhatsApp al 3057939537
                </button>
            </a>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Cerrar Sesión", key="btn_logout_bloqueo", use_container_width=True):
            st.session_state['user'] = None
            if 'user' in st.query_params:
                del st.query_params['user']
            st.rerun()
    st.stop()

# Control de cambio obligatorio de contraseña inicial
if current_user.get('debe_cambiar_password', False):
    col_pwd1, col_pwd2, col_pwd3 = st.columns([1, 2, 1])
    with col_pwd2:
        st.markdown("<br>", unsafe_allow_html=True)
        st.warning("🔒 **Cambio Obligatorio de Contraseña Inicial**")
        st.info("👋 **Bienvenido/a al Sistema.** Estás ingresando con una contraseña inicial o temporal. Por tu seguridad, **debes definir una nueva contraseña personal** para continuar.")
        
        with st.form("form_cambio_password_obligatorio", clear_on_submit=False):
            pass_nueva = st.text_input("🔑 Nueva Contraseña Personal", type="password", help="Ingresa una contraseña segura de al menos 4 caracteres")
            pass_confirm = st.text_input("🔑 Confirmar Nueva Contraseña", type="password", help="Reescribe exactamente la misma contraseña")
            btn_cambiar_pass = st.form_submit_button("💾 Guardar Nueva Contraseña y Continuar", use_container_width=True)
            
            if btn_cambiar_pass:
                if not pass_nueva or len(pass_nueva.strip()) < 4:
                    st.error("❌ La nueva contraseña debe tener al menos 4 caracteres.")
                elif pass_nueva != pass_confirm:
                    st.error("❌ Las contraseñas no coinciden. Por favor reescríbelas.")
                else:
                    ok_p, msg_p = cambiar_password_usuario(current_user['username'], pass_nueva)
                    if ok_p:
                        current_user['debe_cambiar_password'] = False
                        st.session_state['user'] = current_user
                        st.query_params.clear()
                        st.success("✅ " + msg_p)
                        st.rerun()
                    else:
                        st.error("❌ " + msg_p)
                        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Cancelar y Cerrar Sesión", key="btn_cancel_pwd_change", use_container_width=True):
            st.session_state['user'] = None
            st.query_params.clear()
            st.rerun()
    st.stop()

# 3. BARRA LATERAL (Perfil de Usuario, Logout y Opciones según Rol)
st.sidebar.markdown(f"### 👤 {user_nombre}")
if user_rol == 'superadmin':
    st.sidebar.caption("🛠️ **Rol**: Super Administrador del Sistema (Control Total & Roles)")
elif user_rol == 'gerente':
    st.sidebar.caption(f"👑 **Gerencia**: {user_sector_nombre} (Sector `{user_sector}`)")
elif user_rol == 'lider':
    st.sidebar.caption(f"👩‍💼 **Rol**: Líder de Negocio (Grupo `{user_grupo}` • {user_sector_nombre})")
else:
    st.sidebar.caption(f"👤 **Rol**: Asesora / Consulta de Facturación ({user_sector_nombre})")

if info_suscripcion.get("estado") == "prueba":
    st.sidebar.info(f"⏳ **Modo Prueba**: Te quedan **{info_suscripcion['dias_restantes']} días** de prueba gratuita (Vence el {info_suscripcion['fecha_vencimiento_str']}).")

if st.sidebar.button("🚪 Cerrar Sesión", type="secondary"):
    st.session_state['user'] = None
    if 'user' in st.query_params:
        del st.query_params['user']
    if 'msg_timeout' in st.session_state:
        del st.session_state['msg_timeout']
    st.rerun()

modo_tema = st.sidebar.radio(
    "🎨 Tema Visual",
    options=["🌙 Oscuro Neón", "☀️ Modo Claro"],
    index=0,
    horizontal=True,
    key="app_theme_mode_selector"
)
is_dark_theme = (modo_tema == "🌙 Oscuro Neón")

# Inactivador automático en cliente tras 15 minutos sin interacción y activador de corrector nativo
st.markdown("""
<script>
    (function() {
        const TIEMPO_LIMITE_MS = 15 * 60 * 1000; // 15 minutos
        let timeoutInactividad;

        function reiniciarReloj() {
            clearTimeout(timeoutInactividad);
            timeoutInactividad = setTimeout(function() {
                window.location.reload();
            }, TIEMPO_LIMITE_MS);
        }

        ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll', 'click'].forEach(function(evt) {
            window.addEventListener(evt, reiniciarReloj, { passive: true });
        });

        // Habilitar corrector ortográfico nativo del explorador en celdas y campos editables
        function activarCorrectorExplorador(el) {
            if (el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable)) {
                el.setAttribute('spellcheck', 'true');
                el.setAttribute('lang', 'es');
                el.spellcheck = true;
            }
        }
        document.addEventListener('focusin', function(e) {
            activarCorrectorExplorador(e.target);
        }, true);

        reiniciarReloj();
    })();
</script>
""", unsafe_allow_html=True)

st.sidebar.markdown("---")

# Cargar configuración global de permisos
app_config = cargar_configuracion()
puede_subir_archivos = (user_rol in ['gerente', 'superadmin']) or app_config.get('permitir_carga_lideres', False)

# Opciones de subida y administración de archivos
if puede_subir_archivos:
    if user_rol in ['gerente', 'superadmin']:
        st.sidebar.header("🔄 Rotación de Ciclo (Nuevo)")
        st.sidebar.caption("Sube el Excel del nuevo ciclo para convertir el ciclo actual en el 'Como vamos anterior' automáticamente.")
        
        nuevo_ciclo_file = st.sidebar.file_uploader("Cargar Nuevo Ciclo ('Cómo Vamos')", type=["xlsx"], key="uploader_nuevo_ciclo")
        if nuevo_ciclo_file is not None:
            if st.sidebar.button("🚀 Rotar Ciclo y Actualizar Histórico"):
                try:
                    valido, sec_enc, nom_sec, msg_val = validar_archivo_como_vamos(nuevo_ciclo_file, user_sector)
                    if not valido:
                        st.sidebar.error(msg_val)
                    else:
                        with st.spinner("Rotando hojas y guardando nuevo ciclo..."):
                            rotar_y_guardar_nuevo_ciclo(nuevo_ciclo_file)
                            st.cache_data.clear()
                            st.sidebar.success("✅ ¡Ciclo rotado con éxito! El nuevo ciclo ya es el activo.")
                            
                            lideres_creadas = auto_crear_usuarios_lideres_desde_bases()
                            if lideres_creadas:
                                st.session_state['lideres_creadas_log'] = lideres_creadas
                            st.rerun()
                except PermissionError as pe:
                    st.error("⚠️ **Archivo en uso**: El archivo `Base para el como vamos.xlsx` está actualmente abierto en Excel.")
                    st.info("💡 **Solución**: Por favor, **cierra el archivo en Microsoft Excel** y vuelve a presionar el botón '🚀 Rotar Ciclo y Actualizar Histórico'.")
                except Exception as ex:
                    st.error(f"❌ Ocurrió un error al rotar el ciclo: {ex}")

        st.sidebar.markdown("---")
        st.sidebar.subheader("🎯 Metas 'Objetivos Arte'")
        st.sidebar.caption("Sube el archivo `Objetivos Arte.xlsx` (Hoja *Desafíos LNN*) para actualizar las metas de Inicios + Reinicios y Recuperos.")
        
        obj_arte_file = st.sidebar.file_uploader(
            "Cargar 'Objetivos Arte.xlsx'",
            type=["xlsx", "xls"],
            key="uploader_objetivos_arte_sidebar"
        )
        if obj_arte_file is not None:
            if st.sidebar.button("⚡ Actualizar Metas de Inicios, Reinicios y Recuperos", type="primary"):
                try:
                    with st.spinner("Procesando hoja Desafíos LNN de Objetivos Arte..."):
                        res_oa = procesar_archivo_objetivos_arte(obj_arte_file)
                        if res_oa.get('exito'):
                            st.cache_data.clear()
                            st.sidebar.success(f"✅ ¡Metas actualizadas con éxito! ({res_oa.get('total_mapeados', 0)} líderes mapeadas)")
                            st.rerun()
                        else:
                            st.sidebar.error(f"❌ Error al procesar: {res_oa.get('error')}")
                except Exception as ex_oa:
                    st.sidebar.error(f"❌ Ocurrió un error: {ex_oa}")

        st.sidebar.markdown("---")
else:
    st.sidebar.info("🔒 **Carga restringida**: La opción de subida de archivos está desactivada por la Gerencia General para tu perfil.")

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Recargar Datos Actuales"):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")

# Carga de datos
with st.spinner("Cargando y procesando la base de datos..."):
    df_raw = load_and_process_data('Base para el como vamos.xlsx')

if df_raw is None:
    df = pd.DataFrame()
else:
    df = df_raw.copy()
    # Omitir filas 'None', 'NaN' o vacías que dañan la presentación visual de las tablas
    col_lider_check = 'Nombre de consultora' if 'Nombre de consultora' in df.columns else df.columns[0]
    if col_lider_check in df.columns:
        mask_valida_df = df[col_lider_check].notna() & \
                          (~df[col_lider_check].astype(str).str.strip().str.lower().isin(['none', 'nan', '', 'null', '0']))
        df = df[mask_valida_df]

    # Aislamiento Multitenant de Gerencias: Filtrar df por el sector asignado a la Gerente
    if user_rol == 'gerente':
        if user_sector:
            col_sec_found = None
            for c in df.columns:
                c_low = str(c).lower().replace('ó', 'o')
                if 'setor' in c_low or 'sector' in c_low:
                    col_sec_found = c
                    break
                    
            if col_sec_found:
                s_vals = df[col_sec_found].astype(str).str.strip().str.replace('.0', '', regex=False)
                df = df[s_vals == str(user_sector).strip()]
            else:
                df = df.iloc[0:0]
        else:
            # Si la gerente NO tiene sector asignado en usuarios.json, mostrar vista limpia de 0 filas
            df = df.iloc[0:0]

# Header Principal Dinámico según el Sector del Usuario
st.markdown(f"<div class='main-header'>📈 Panel de Control - Estado de Ciclo {user_sector_nombre}</div>", unsafe_allow_html=True)
st.markdown(f"<div class='sub-header'>Gestión de Líderes, Seguimiento de Metas e Indicadores de Crecimiento • {user_sector_nombre}</div>", unsafe_allow_html=True)

# Diagnóstico informativo si no hay datos de metas en "Cómo Vamos" para el sector o rol activo
if df.empty:
    if user_rol == 'gerente':
        info_sector_actual = f"**{user_sector_nombre}** (Cód: `{user_sector}`)"
        
        st.info(
            f"ℹ️ **Bienvenida a tu Panel de Control — {info_sector_actual}**\n\n"
            f"Aún no se han cargado las metas del ciclo actual en **'Cómo Vamos'** para tu sector.\n\n"
            f"✅ **Tus otros módulos se encuentran 100% operativos:**\n"
            f"- Puedes gestionar tu red de consultoras en la pestaña **'📊 Informe Tableau Cam'**.\n"
            f"- Puedes consultar los reportes de cartera y cobranza en **'💳 Geral_Credito&Cobranza'**.\n\n"
            f"📥 **Para activar los tacómetros y gráficas de metas de tu sector:**\n"
            f"Sube el archivo Excel de metas desde la barra lateral izquierda en **'🔄 Rotación de Ciclo (Nuevo)'**."
        )
    elif user_rol == 'lider':
        st.info(
            f"ℹ️ **Información de Metas de Ciclo (Grupo {user_grupo}):**\n\n"
            f"Aún no se encuentran cargadas las metas del ciclo actual en el archivo 'Cómo Vamos' para tu grupo. "
            f"Puedes seguir consultando a tus consultoras y estados de cartera en las pestañas de **Tableau** y **Geral**."
        )
else:
    # Alerta visual si se cargó un archivo sin metas financieras (ej. Reporte de Niveles en vez de Cómo Vamos)
    if 'Objetivo Facturación' in df.columns and 'Real Facturación' in df.columns:
        tot_obj_v = pd.to_numeric(df['Objetivo Facturación'], errors='coerce').fillna(0).sum()
        tot_real_v = pd.to_numeric(df['Real Facturación'], errors='coerce').fillna(0).sum()
        if tot_obj_v == 0 and tot_real_v == 0:
            st.warning(
                "⚠️ **Aviso Importante — Datos de Metas en $0 (Reporte Incompleto o de Niveles)**\n\n"
                "Los datos cargados actualmente para tu sector muestran **todas las metas y ventas en `$0`** "
                "(esto ocurre cuando se sube un **'Reporte de Niveles / Puntos'** en lugar del archivo oficial de **'Cómo Vamos'**).\n\n"
                "💡 **¿Cómo ver tus ventas y metas reales?**:\n"
                "1. Descarga el reporte oficial **'Cómo Vamos'** desde el portal de Natura.\n"
                "2. Súbelo en la barra lateral izquierda en **'🔄 Rotación de Ciclo (Nuevo)'**.\n"
                "3. El sistema actualizará al instante los tacómetros 360°, avances % y ganancias reales de cada líder."
            )

# Notificación de Cuentas de Nuevas Líderes Auto-Generadas
if 'lideres_creadas_log' in st.session_state and st.session_state['lideres_creadas_log']:
    with st.expander("✨ Cuentas de Nuevas Líderes Creadas Automáticamente", expanded=True):
        st.success("🎉 **¡Automatización Completada!** El sistema detectó **nuevas líderes** en las bases de datos y generó sus accesos iniciales:")
        df_l_log = pd.DataFrame(st.session_state['lideres_creadas_log'])
        st.dataframe(df_l_log, use_container_width=True)
        st.info("💡 **Nota para la Gerente**: Comparte con cada nueva líder su usuario (correo/código) y la contraseña temporal generada.")
        if st.button("Entendido / Cerrar Notificación"):
            del st.session_state['lideres_creadas_log']
            st.rerun()

# 3. BARRA LATERAL (Filtros dinámicos)
st.sidebar.header("🔐 Filtros de Control")

# Filtro por Gerencia
col_gerencia = 'Nombre Gerencia' if 'Nombre Gerencia' in df.columns else (df.columns[0] if len(df.columns) > 0 else '')
gerencias_disponibles = sorted([str(g) for g in df[col_gerencia].dropna().unique()]) if (col_gerencia and col_gerencia in df.columns) else []
gerencia_seleccionada = st.sidebar.selectbox(
    "🏢 Selecciona la Gerencia",
    options=["Todas"] + gerencias_disponibles,
    index=0
)

df_filtrado = df.copy()

# Segmentación Privada por Rol (Preservando el código madre intacto)
# Si ingresa una Líder de Negocio, sus tarjetas superiores, tacómetros y reportes se restringen automáticamente a su Grupo
if user_rol == 'lider' and user_grupo and not df_filtrado.empty:
    col_grp_ref = 'Código de grupo' if 'Código de grupo' in df_filtrado.columns else ''
    if col_grp_ref and col_grp_ref in df_filtrado.columns:
        df_filtrado = df_filtrado[df_filtrado[col_grp_ref].astype(str).str.strip() == str(user_grupo).strip()]

if gerencia_seleccionada != "Todas" and col_gerencia and col_gerencia in df_filtrado.columns:
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
else:
    sector_seleccionado = "Todos"

# Filtro Global por Líder / Grupo (Habilitado para Gerencia y SuperAdmin)
if user_rol in ['gerente', 'superadmin'] and not df_filtrado.empty:
    col_grp_ref = 'Código de grupo' if 'Código de grupo' in df_filtrado.columns else ''
    if col_grp_ref and col_grp_ref in df_filtrado.columns:
        grupos_unicos = sorted([str(g).strip() for g in df_filtrado[col_grp_ref].dropna().unique()])
        mapa_lideres_sb = obtener_mapa_lideres()
        
        def format_lider_sb(g_val):
            if g_val == "Todas las Líderes":
                return "🌟 Todas las Líderes"
            nom = mapa_lideres_sb.get(str(g_val).strip())
            if nom:
                return f"👩‍💼 Grupo {g_val} — {nom}"
            return f"👥 Grupo {g_val}"

        lider_seleccionada_sb = st.sidebar.selectbox(
            "👤 Seleccionar Líder / Grupo",
            options=["Todas las Líderes"] + grupos_unicos,
            format_func=format_lider_sb,
            index=0
        )
        if lider_seleccionada_sb != "Todas las Líderes":
            df_filtrado = df_filtrado[df_filtrado[col_grp_ref].astype(str).str.strip() == str(lider_seleccionada_sb).strip()]

# Filtro por Color / Clasificación
if 'Color' in df_filtrado.columns:
    colores_disponibles = sorted([str(c) for c in df_filtrado['Color'].dropna().unique()])
    colores_seleccionados = st.sidebar.multiselect(
        "🎨 Clasificación / Color",
        options=colores_disponibles,
        default=[]
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

def renderizar_modo_app(df_filtrado, user_rol, user_nombre, user_grupo, user_sector, is_dark_theme):
    if is_dark_theme:
        css_theme = """
        <style>
        .stApp { background-color: #0b0f19 !important; color: #f8fafc !important; }
        .app-card-hero {
            background: linear-gradient(135deg, rgba(15, 23, 42, 0.9), rgba(30, 41, 59, 0.9));
            border: 1px solid rgba(56, 189, 248, 0.3);
            border-radius: 16px;
            padding: clamp(10px, 1.2vw, 16px);
            text-align: center;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
            min-width: 0 !important;
            box-sizing: border-box !important;
        }
        .hero-title { font-size: clamp(10px, 0.8vw, 12px); color: #94a3b8; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; white-space: nowrap; text-overflow: ellipsis; overflow: hidden; }
        .hero-val { font-size: clamp(16px, 1.6vw, 24px); font-weight: 800; color: #38bdf8; margin: 4px 0; white-space: nowrap; text-overflow: ellipsis; overflow: hidden; }
        .hero-sub { font-size: clamp(10px, 0.8vw, 12px); color: #10b981; font-weight: 600; white-space: nowrap; }
        </style>
        """
    else:
        css_theme = """
        <style>
        .stApp { background-color: #f8fafc !important; color: #0f172a !important; }
        .app-card-hero {
            background: linear-gradient(135deg, #ffffff, #f1f5f9);
            border: 1px solid #cbd5e1;
            border-radius: 16px;
            padding: clamp(10px, 1.2vw, 16px);
            text-align: center;
            box-shadow: 0 4px 20px 0 rgba(0, 0, 0, 0.06);
            min-width: 0 !important;
            box-sizing: border-box !important;
        }
        .hero-title { font-size: clamp(10px, 0.8vw, 12px); color: #64748b; text-transform: uppercase; letter-spacing: 1px; font-weight: 600; white-space: nowrap; text-overflow: ellipsis; overflow: hidden; }
        .hero-val { font-size: clamp(16px, 1.6vw, 24px); font-weight: 800; color: #0284c7; margin: 4px 0; white-space: nowrap; text-overflow: ellipsis; overflow: hidden; }
        .hero-sub { font-size: clamp(10px, 0.8vw, 12px); color: #059669; font-weight: 600; white-space: nowrap; }
        </style>
        """
    st.markdown(css_theme, unsafe_allow_html=True)

    df_tab_app = consultar_tableau_sql(
        grupo=(user_grupo if user_rol == 'lider' else None),
        sector=(user_sector if (user_rol == 'gerente' and user_sector) else ('__INVALID_SECTOR__' if user_rol == 'gerente' else None))
    )

    r_act = float(df_filtrado['Real Activas'].sum()) if 'Real Activas' in df_filtrado.columns else 0.0
    o_act = float(df_filtrado['Objetivo Activas'].sum()) if 'Objetivo Activas' in df_filtrado.columns else 0.0
    c_act = (r_act / o_act * 100.0) if o_act > 0 else 0.0

    r_fact = float(df_filtrado['Real Facturación'].sum()) if 'Real Facturación' in df_filtrado.columns else 0.0
    o_fact = float(df_filtrado['Objetivo Facturación'].sum()) if 'Objetivo Facturación' in df_filtrado.columns else 0.0
    c_fact = (r_fact / o_fact * 100.0) if o_fact > 0 else 0.0

    gan_tot = float(df_filtrado['Ganancia estimada'].sum()) if 'Ganancia estimada' in df_filtrado.columns else 0.0

    pts_tot = 0
    deuda_tot = 0.0
    if not df_tab_app.empty:
        if 'Pts Acum' in df_tab_app.columns:
            pts_tot = int(df_tab_app['Pts Acum'].apply(lambda x: limpiar_numero(x, 0)).sum())
        if 'Deuda Mora' in df_tab_app.columns:
            deuda_tot = float(df_tab_app['Deuda Mora'].apply(lambda x: limpiar_numero(x, 0)).sum())

    h1, h2, h3, h4 = st.columns(4)
    with h1:
        st.markdown(f"""
        <div class="app-card-hero">
            <div class="hero-title">💰 Facturación Total</div>
            <div class="hero-val">${r_fact/1e6:.1f}M COP</div>
            <div class="hero-sub">↑ {c_fact:.1f}% Cumplimiento</div>
        </div>
        """, unsafe_allow_html=True)
    with h2:
        st.markdown(f"""
        <div class="app-card-hero">
            <div class="hero-title">👥 Activas Reales</div>
            <div class="hero-val">{int(r_act)}</div>
            <div class="hero-sub">↑ {c_act:.1f}% Cumplimiento</div>
        </div>
        """, unsafe_allow_html=True)
    with h3:
        st.markdown(f"""
        <div class="app-card-hero">
            <div class="hero-title">💵 Ganancia Estimada</div>
            <div class="hero-val">${gan_tot:,.0f}</div>
            <div class="hero-sub">✨ Proyección LN</div>
        </div>
        """.replace(",", "."), unsafe_allow_html=True)
    with h4:
        st.markdown(f"""
        <div class="app-card-hero">
            <div class="hero-title">⭐ Puntos / Deuda Mora</div>
            <div class="hero-val">{pts_tot:,} pts</div>
            <div class="hero-sub" style="color: #ef4444;">⚠️ Mora: ${deuda_tot/1e6:.1f}M</div>
        </div>
        """.replace(",", "."), unsafe_allow_html=True)

    # Recordatorio y Banner de Cumpleaños en Modo Minimalista
    renderizar_banner_cumpleanos(df_tab_app, user_rol, user_nombre, user_grupo, user_sector, key_suffix="app_mini")

    st.markdown("<br>", unsafe_allow_html=True)

    tab_app1, tab_app2, tab_app3 = st.tabs([
        "📱 1. Puntos & Listado Consultoras",
        "⚡ 2. KPIs & Tacómetros 360°",
        "👑 3. Mis Líderes (Metas & Facturación)"
    ])

    with tab_app1:
        st.subheader("📊 Listado Minimalista de Consultoras & Puntos")
        if df_tab_app.empty:
            st.info("No se encontraron registros de consultoras para esta vista.")
        else:
            cols_show = [c for c in ['Código CB', 'Líder / Grupo', 'Asesora / Consultora', 'Nivel / Color', 'Sit. Comercial', 'Pts Acum', 'Deuda Mora', 'Ped. Pendientes', 'Notas / Comentarios'] if c in df_tab_app.columns]
            st.dataframe(
                df_tab_app[cols_show] if cols_show else df_tab_app,
                use_container_width=True,
                height=450
            )

    with tab_app2:
        st.subheader("⏱️ Tacómetros de Cumplimiento Global 360°")
        tcol1, tcol2 = st.columns(2)
        with tcol1:
            fig_g_fact = go.Figure(go.Indicator(
                mode="gauge+number",
                value=c_fact,
                number={'suffix': '%'},
                title={'text': "Cumplimiento Facturación ($ COP)"},
                gauge={
                    'axis': {'range': [0, 150]},
                    'bar': {'color': "#38bdf8"},
                    'steps': [
                        {'range': [0, 95], 'color': "#fee2e2"},
                        {'range': [95, 100], 'color': "#fef3c7"},
                        {'range': [100, 150], 'color': "#dcfce7"}
                    ]
                }
            ))
            fig_g_fact.update_layout(height=280, margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig_g_fact, use_container_width=True)
            
        with tcol2:
            fig_g_act = go.Figure(go.Indicator(
                mode="gauge+number",
                value=c_act,
                number={'suffix': '%'},
                title={'text': "Cumplimiento Activas"},
                gauge={
                    'axis': {'range': [0, 150]},
                    'bar': {'color': "#10b981"},
                    'steps': [
                        {'range': [0, 95], 'color': "#fee2e2"},
                        {'range': [95, 100], 'color': "#fef3c7"},
                        {'range': [100, 150], 'color': "#dcfce7"}
                    ]
                }
            ))
            fig_g_act.update_layout(height=280, margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig_g_act, use_container_width=True)

    with tab_app3:
        st.subheader("📈 Tabla de Facturación y Cumplimiento (Mayor a Menor Desempeño)")
        if df_filtrado.empty:
            st.info("No hay datos disponibles para mostrar la tabla de facturación.")
        else:
            df_fact_ordenado = df_filtrado.sort_values(by='Cumplimiento Facturación', ascending=False) if 'Cumplimiento Facturación' in df_filtrado.columns else df_filtrado
            cols_fact = [c for c in ['Nombre de consultora', 'Nombre Setor', 'Objetivo Facturación', 'Real Facturación', 'Cumplimiento Facturación', 'Falta para el 100%', 'Ganancia estimada'] if c in df_fact_ordenado.columns]
            
            df_fact_view = df_fact_ordenado[cols_fact].copy()
            if 'Cumplimiento Facturación' in df_fact_view.columns:
                df_fact_view['Cumplimiento Facturación'] = df_fact_view['Cumplimiento Facturación'].apply(lambda x: f"{x:.1f}%")
            if 'Real Facturación' in df_fact_view.columns:
                df_fact_view['Real Facturación'] = df_fact_view['Real Facturación'].apply(formato_cop)
            if 'Objetivo Facturación' in df_fact_view.columns:
                df_fact_view['Objetivo Facturación'] = df_fact_view['Objetivo Facturación'].apply(formato_cop)
            if 'Ganancia estimada' in df_fact_view.columns:
                df_fact_view['Ganancia estimada'] = df_fact_view['Ganancia estimada'].apply(formato_cop)

            st.dataframe(df_fact_view, use_container_width=True, height=450)

# Cálculo general de métricas clave del dataset filtrado
total_consultoras = len(df_filtrado)

obj_activas = float(df_filtrado['Objetivo Activas'].sum()) if 'Objetivo Activas' in df_filtrado.columns else 0.0
real_activas = float(df_filtrado['Real Activas'].sum()) if 'Real Activas' in df_filtrado.columns else 0.0
cump_activas = (real_activas / obj_activas * 100.0) if obj_activas > 0 else 0.0

obj_fact = float(df_filtrado['Objetivo Facturación'].sum()) if 'Objetivo Facturación' in df_filtrado.columns else 0.0
real_fact = float(df_filtrado['Real Facturación'].sum()) if 'Real Facturación' in df_filtrado.columns else 0.0
cump_fact = (real_fact / obj_fact * 100.0) if obj_fact > 0 else 0.0

ganancia_total = float(df_filtrado['Ganancia estimada'].sum()) if 'Ganancia estimada' in df_filtrado.columns else 0.0

inicios_totales = float(df_filtrado['Inicios'].sum()) if 'Inicios' in df_filtrado.columns else 0.0
reinicios_totales = float(df_filtrado['Reinicios'].sum()) if 'Reinicios' in df_filtrado.columns else 0.0

# 4. TARJETAS DE KPIS SUPERIORES (VISTA COMPLETA)
# Recordatorio y Banner de Cumpleaños para Líderes y Gerentes
grupo_cumple_filtro = user_grupo if user_rol == 'lider' else (lider_seleccionada_sb if ('lider_seleccionada_sb' in locals() and lider_seleccionada_sb != "Todas las Líderes") else None)
sector_cumple_filtro = user_sector if (user_rol == 'gerente' and user_sector) else ('__INVALID_SECTOR__' if user_rol == 'gerente' else None)
df_tableau_cumple = consultar_tableau_sql(grupo=grupo_cumple_filtro, sector=sector_cumple_filtro)
renderizar_banner_cumpleanos(df_tableau_cumple, user_rol, user_nombre, user_grupo, user_sector, key_suffix="full_top")

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
    gan_txt = f"${ganancia_total/1e6:.2f}M COP" if abs(ganancia_total) >= 1_000_000 else f"${ganancia_total:,.0f}".replace(",", ".")
    st.metric(
        "💵 Ganancia Estimada LN",
        gan_txt
    )

with kpi5:
    st.metric(
        "🚀 Inicios / Reinicios",
        f"{int(inicios_totales + reinicios_totales)}",
        f"↑ {int(inicios_totales)} Inic. | {int(reinicios_totales)} Rein."
    )

st.markdown("---")

# 5. CONTENIDO CON PESTAÑAS (TABS)
permisos_tab_config = app_config.get("permisos_pestanas", DEFAULT_PERMISOS_PESTANAS)

tabs_definidas = [
    ("tab_tableau", "📊 Informe Tableau Cam"),
    ("tab_geral", "💳 Geral_Credito&Cobranza"),
    ("tab_resumen", "📊 Resumen & KPIs"),
    ("tab_ganancia", "💵 Simulador de Ganancia"),
    ("tab_diagnostico", "👑 Mis Líderes"),
    ("tab_metas", "🎯 Metas de Crecimiento (Procesador)"),
    ("tab_detalle", "👥 Detalle Completo"),
    ("tab_exportar", "📤 Exportar Datos")
]

if user_rol == 'superadmin':
    tabs_definidas.append(("tab_usuarios", "🔑 Gestión de Usuarios, Roles & Permisos"))
elif user_rol == 'gerente':
    tabs_definidas.append(("tab_lideres_gerente", "🔑 Directorio & Accesos de Líderes"))

tabs_permitidas = []
for key_tab, label_tab in tabs_definidas:
    if user_rol == 'superadmin':
        tabs_permitidas.append((key_tab, label_tab))
    else:
        perm_rol = permisos_tab_config.get(key_tab, {}).get(user_rol, True)
        if perm_rol:
            tabs_permitidas.append((key_tab, label_tab))

if not tabs_permitidas:
    st.warning("🔒 No tienes acceso habilitado a ningún apartado actualmente. Contacta al administrador del sistema.")
else:
    list_tab_objects = st.tabs([label for _, label in tabs_permitidas])
    class HiddenTab:
        def __enter__(self):
            self.ph = st.empty()
            self.container = self.ph.container()
            return self.container.__enter__()
        def __exit__(self, exc_type, exc_val, exc_tb):
            res = self.container.__exit__(exc_type, exc_val, exc_tb)
            self.ph.empty()
            return res

    tab_objs = {key: obj for (key, _), obj in zip(tabs_permitidas, list_tab_objects)}
    
    tab_tableau = tab_objs.get("tab_tableau") or HiddenTab()
    tab_geral = tab_objs.get("tab_geral") or HiddenTab()
    tab_resumen = tab_objs.get("tab_resumen") or HiddenTab()
    tab_ganancia = tab_objs.get("tab_ganancia") or HiddenTab()
    tab_diagnostico = tab_objs.get("tab_diagnostico") or HiddenTab()
    tab_metas = tab_objs.get("tab_metas") or HiddenTab()
    tab_detalle = tab_objs.get("tab_detalle") or HiddenTab()
    tab_exportar = tab_objs.get("tab_exportar") or HiddenTab()
    tab_usuarios = tab_objs.get("tab_usuarios") or HiddenTab()
    tab_lideres_gerente = tab_objs.get("tab_lideres_gerente") or HiddenTab()

# --- TAB 0: INFORME TABLEAU MANAGER ("INFORME TABLEAU CAM") ---
with tab_tableau:
    # 1. Cargar la base desde SQLite (Consulta SQL ultrarrápida indexada aislada por sector/grupo)
    df_tableau = consultar_tableau_sql(
        grupo=(user_grupo if user_rol == 'lider' else None),
        sector=(user_sector if (user_rol == 'gerente' and user_sector) else ('__INVALID_SECTOR__' if user_rol == 'gerente' else None))
    )
    
    if df_tableau is None or df_tableau.empty:
        st.warning("⚠️ No se encontró la base de datos `Base de Datos.xlsx`. Por favor, sube el archivo desde la barra lateral o el panel de administración.")
    else:
        # Filtros de navegación rápida para Tableau Manager
        col_t1, col_t2, col_t3, col_t4 = st.columns([1, 1, 1.2, 1])
        
        # 1. Filtro de Gerencia
        gerencias_t = sorted([str(g) for g in df_tableau['Gerencia'].dropna().unique()]) if 'Gerencia' in df_tableau.columns else []
        with col_t1:
            ger_sel_t = st.selectbox("🏢 Gerencia (Tableau)", options=["Todas"] + gerencias_t, key="tab_ger_sel")
        
        df_tab_filt = df_tableau.copy()
        if ger_sel_t != "Todas" and 'Gerencia' in df_tab_filt.columns:
            df_tab_filt = df_tab_filt[df_tab_filt['Gerencia'] == ger_sel_t]
            
        # 2. Filtro de Sector
        sectores_t = sorted([str(s) for s in df_tab_filt['Sector'].dropna().unique()]) if 'Sector' in df_tab_filt.columns else []
        with col_t2:
            sec_sel_t = st.selectbox("📍 Sector (Tableau)", options=["Todos"] + sectores_t, key="tab_sec_sel")
            
        if sec_sel_t != "Todos" and 'Sector' in df_tab_filt.columns:
            df_tab_filt = df_tab_filt[df_tab_filt['Sector'] == sec_sel_t]

        # 3. Filtro Directo por Líder / Grupo [NUEVO PARA GERENCIA]
        lista_grupos_t = sorted([str(g).strip() for g in df_tab_filt['Grupo'].dropna().unique()]) if 'Grupo' in df_tab_filt.columns else []
        mapa_lideres_tab = obtener_mapa_lideres()

        def format_lider_tab(g_val):
            if g_val == "Todas las Líderes (Consolidado Zona)":
                return "🌟 Todas las Líderes (Consolidado Zona)"
            nom = mapa_lideres_tab.get(str(g_val).strip())
            if nom:
                return f"👩‍💼 Grupo {g_val} — {nom}"
            return f"👥 Grupo {g_val}"

        with col_t3:
            lider_sel_t = st.selectbox(
                "👤 Selección de Líder / Grupo",
                options=["Todas las Líderes (Consolidado Zona)"] + lista_grupos_t,
                format_func=format_lider_tab,
                key="tab_lider_grp_sel"
            )

        if lider_sel_t != "Todas las Líderes (Consolidado Zona)" and 'Grupo' in df_tab_filt.columns:
            df_tab_filt = df_tab_filt[df_tab_filt['Grupo'].astype(str).str.strip() == str(lider_sel_t).strip()]

        # 4. Buscador de Asesora / Consultora
        with col_t4:
            busq_t = st.text_input("🔍 Buscar Asesora (Nombre / Código)", "", key="tab_busq")
            
        if busq_t.strip():
            mask_t = pd.Series(False, index=df_tab_filt.index)
            if 'Nombre' in df_tab_filt.columns:
                mask_t = mask_t | df_tab_filt['Nombre'].astype(str).str.contains(busq_t, case=False, na=False)
            if 'Codigo CB' in df_tab_filt.columns:
                mask_t = mask_t | df_tab_filt['Codigo CB'].astype(str).str.contains(busq_t, case=False, na=False)
            df_tab_filt = df_tab_filt[mask_t]

        # Tarjetas de resumen rápido Tableau (Visibles en la cabecera superior de Informe Tableau)
        tc1, tc2, tc3, tc4, tc5 = st.columns(5)
        with tc1:
            st.metric("👥 Total cadastro", f"{len(df_tab_filt)}")
        with tc2:
            col_sit_check = 'Sit. Comercial' if 'Sit. Comercial' in df_tab_filt.columns else ('Situación' if 'Situación' in df_tab_filt.columns else None)
            if col_sit_check and not df_tab_filt.empty:
                s_vals_lower = df_tab_filt[col_sit_check].astype(str).str.strip().str.lower()
                mask_disp = s_vals_lower.apply(
                    lambda s: any(k in s for k in ['activa', 'activas', 'inactiva 1', 'inactiva 2', 'inactiva 3', 'i1', 'i2', 'i3']) and not any(k in s for k in ['inactiva 4', 'inactiva 5', 'inactiva 6', 'i4', 'i5', 'i6'])
                )
                tot_disponibles = int(mask_disp.sum())
            else:
                tot_disponibles = 0
            st.metric("🎯 Total disponibles", f"{tot_disponibles}")
        with tc3:
            tot_mora = df_tab_filt['Deuda Mora'].sum() if 'Deuda Mora' in df_tab_filt.columns else 0
            st.metric("⚠️ Deuda Mora Total", f"${tot_mora/1e6:.2f}M COP")
        with tc4:
            tot_cred = df_tab_filt['Credito Disponible'].sum() if 'Credito Disponible' in df_tab_filt.columns else 0
            st.metric("💳 Crédito Dispon.", f"${tot_cred/1e6:.2f}M COP")
        with tc5:
            tot_pago = len(df_tab_filt[df_tab_filt['Ped. Pendientes'] > 0]) if 'Ped. Pendientes' in df_tab_filt.columns else 0
            st.metric("⌛ Aguardando Pago", f"{tot_pago} pers.")

        # Filtros adicionales por columna en un expansor dedicado (colapsado por defecto)
        with st.expander("🔍 Filtros Avanzados por Columna (Sit. Comercial, Nivel, Mora, Pedidos y Notas)", expanded=False):
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

        # Opciones de administración (visible solo para Gerencia y SuperAdmin en expander colapsado)
        if user_rol in ['gerente', 'superadmin']:
            with st.expander("⚙️ Opciones de Administración (Actualizar Base Tableau, mi_grupo & Activas)", expanded=False):
                col_adm1, col_adm2, col_adm3 = st.columns(3)
                with col_adm1:
                    st.markdown("###### 📁 1. Base Tableau Completa (`Base de Datos.xlsx`)")
                    archivo_tableau = st.file_uploader("Selecciona `Base de Datos.xlsx`", type=["xlsx"], key="tableau_uploader")
                    if archivo_tableau is not None:
                        file_id = f"{archivo_tableau.name}_{archivo_tableau.size}"
                        if st.session_state.get('last_processed_tableau') != file_id:
                            valido, sec_enc, nom_sec, msg_val = validar_sector_archivo(archivo_tableau, user_sector)
                            if not valido:
                                st.error(msg_val)
                            else:
                                try:
                                    with open("Base de Datos.xlsx", "wb") as f:
                                        f.write(archivo_tableau.getbuffer())
                                    ok_sync = sincronizar_excel_tableau_a_sqlite("Base de Datos.xlsx")
                                    
                                    if ok_sync:
                                        st.cache_data.clear()
                                        st.session_state['last_processed_tableau'] = file_id
                                        
                                        # Actualizar DataFrame de Tableau en vivo para la vista actual
                                        df_tableau = consultar_tableau_sql(
                                            grupo=(user_grupo if user_rol == 'lider' else None),
                                            sector=(user_sector if user_rol == 'gerente' else None)
                                        )
                                        
                                        lideres_creadas = auto_crear_usuarios_lideres_desde_bases()
                                        if lideres_creadas:
                                            st.session_state['lideres_creadas_log'] = lideres_creadas
                                        st.success("✅ ¡Base de Datos.xlsx actualizada y convertida a SQL exitosamente!")
                                        st.rerun()
                                    else:
                                        st.error("⚠️ El archivo subido no corresponde a la Base Maestra de Tableau ('Base de Datos.xlsx'). Si deseas actualizar las metas de ciclo, súbelo en la barra lateral izquierda en '🔄 Rotación de Ciclo (Nuevo)'.")
                                except Exception as e_up:
                                    st.error(f"Error al actualizar la base: {e_up}")
                        else:
                            st.success("✅ ¡Base de datos activa y cargada exitosamente!")

                with col_adm2:
                    st.markdown("###### 🔄 2. Opción A: Cargar `mi_grupo`")
                    st.caption("Actualiza la Situación Comercial de cada consultora vinculando por Código CB.")
                    file_mg = st.file_uploader("Selecciona `mi_grupo.xls`", type=["xls", "xlsx"], key="mi_grupo_uploader_tab")
                    
                    if os.path.exists("mi_grupo.xls"):
                        if st.button("⚡ Cruzar desde 'mi_grupo.xls' local", type="secondary", key="btn_mg_local"):
                            res_mg = actualizar_situacion_comercial_desde_mi_grupo("mi_grupo.xls")
                            if res_mg.get('exito'):
                                st.cache_data.clear()
                                st.session_state['res_mg_log'] = {
                                    'msg': f"✅ ¡Actualización exitosa! {res_mg['coincidencias']} coincidencia(s), {res_mg['cambios']} cambio(s) de estado.",
                                    'detalles': res_mg.get('detalles', [])
                                }
                                st.rerun()
                            else:
                                st.error(f"Error: {res_mg.get('error')}")

                    if file_mg is not None:
                        if st.button("🚀 Actualizar desde 'mi_grupo'", type="primary", key="btn_mg_subido"):
                            res_mg = actualizar_situacion_comercial_desde_mi_grupo(file_mg)
                            if res_mg.get('exito'):
                                st.cache_data.clear()
                                st.session_state['res_mg_log'] = {
                                    'msg': f"✅ ¡Actualización exitosa! {res_mg['coincidencias']} coincidencia(s), {res_mg['cambios']} cambio(s) de estado.",
                                    'detalles': res_mg.get('detalles', [])
                                }
                                st.rerun()
                            else:
                                st.error(f"Error: {res_mg.get('error')}")

                with col_adm3:
                    st.markdown("###### ⚡ 3. Opción B: Cargar `activas`")
                    st.caption("Actualiza estados a 'Activa', pedidos, facturación y puntos vinculando por Código CB.")
                    file_act = st.file_uploader("Selecciona archivo `activas`", type=["xlsx", "xls", "csv"], key="activas_uploader_tab")
                    
                    local_act_path = next((p for p in ["activas.xlsx", "activas.xls", "activas.csv", "Activas.xlsx", "Activas.xls"] if os.path.exists(p)), None)
                    if local_act_path:
                        if st.button(f"⚡ Cruzar desde '{local_act_path}' local", type="secondary", key="btn_act_local"):
                            res_act = actualizar_base_desde_activas(local_act_path)
                            if res_act.get('exito'):
                                st.cache_data.clear()
                                st.session_state['res_act_log'] = {
                                    'msg': f"✅ ¡Cruce de Activas exitoso! {res_act['coincidencias']} coincidencia(s), {res_act['cambios_totales']} consultora(s) actualizada(s).",
                                    'detalles': res_act.get('detalles', [])
                                }
                                st.rerun()
                            else:
                                st.error(f"Error: {res_act.get('error')}")

                    if file_act is not None:
                        if st.button("🚀 Cruzar y Actualizar Activas", type="primary", key="btn_act_subido"):
                            res_act = actualizar_base_desde_activas(file_act)
                            if res_act.get('exito'):
                                st.cache_data.clear()
                                st.session_state['res_act_log'] = {
                                    'msg': f"✅ ¡Cruce de Activas exitoso! {res_act['coincidencias']} coincidencia(s), {res_act['cambios_totales']} consultora(s) actualizada(s).",
                                    'detalles': res_act.get('detalles', [])
                                }
                                st.rerun()
                            else:
                                st.error(f"Error: {res_act.get('error')}")

                if st.session_state.get('res_mg_log'):
                    log_mg = st.session_state['res_mg_log']
                    st.success(log_mg['msg'])
                    if log_mg.get('detalles'):
                        with st.expander("📋 Ver detalle de Consultoras Actualizadas con mi_grupo", expanded=True):
                            st.dataframe(pd.DataFrame(log_mg['detalles']), use_container_width=True)
                    if st.button("Cerrar notificación (mi_grupo)", key="btn_close_mg_log"):
                        del st.session_state['res_mg_log']
                        st.rerun()

                if st.session_state.get('res_act_log'):
                    log_data = st.session_state['res_act_log']
                    st.success(log_data['msg'])
                    if log_data.get('detalles'):
                        with st.expander("📋 Ver detalle de Consultoras Actualizadas con Activas", expanded=True):
                            st.dataframe(pd.DataFrame(log_data['detalles']), use_container_width=True)
                    if st.button("Cerrar notificación (activas)", key="btn_close_act_log"):
                        del st.session_state['res_act_log']
                        st.rerun()

        st.markdown("---")

        # Subpestañas internas dentro de Informe Tableau Cam
        tab_tab_main, tab_tab_pago, tab_tab_niveles, tab_tab_whatsapp, tab_tab_cumple = st.tabs([
            "📋 Base Maestra Gestionable",
            "⌛ Aguardando Pago / Pendientes",
            "🎨 Análisis por Nivel & Estado",
            "📲 Asistente & Campañas WhatsApp",
            "🎂 Cumpleaños & Reconocimiento"
        ])

        # --- SUBPESTAÑA 1: BASE MAESTRA GESTIONABLE ---
        with tab_tab_main:
            # Editor de Comentarios en Masa / Guardar Comentarios
            st.markdown("##### 📝 Comentarios y Notas Persistentes de la Líder")
            st.caption("Escribe las notas de gestión por cada asesora. Se guardarán de forma permanente por `Codigo CB`. Puedes usar el corrector del explorador (subrayado rojo y clic derecho) para sugerencias ortográficas directas.")

            # Limpiar, ordenar y estandarizar columnas para que coincidan exactamente con la base canónica (16 columnas)
            df_edit_view = limpiar_y_ordenar_columnas_tableau(df_tab_filt, mapa_lideres_tab)

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

            edited_df = st.data_editor(
                df_edit_styled,
                column_config=col_config,
                use_container_width=True,
                hide_index=True,
                key="editor_tabla_tableau"
            )

            # Auto-guardado inteligente en segundo plano al modificar cualquier celda
            editor_state = st.session_state.get("editor_tabla_tableau", {})
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

            # Barra de control y respaldo manual
            col_save1, col_save2 = st.columns([1.5, 2.5])
            with col_save1:
                if st.button("💾 Guardar Manualmente", type="primary", use_container_width=True):
                    dict_guardar = {}
                    for idx, row in edited_df.iterrows():
                        codigo_key = str(row.get('Código CB', '')).strip()
                        nota_val = str(row.get('Notas / Comentarios Líder', '')).strip()
                        if codigo_key:
                            dict_guardar[codigo_key] = nota_val
                    
                    if guardar_todos_comentarios(dict_guardar):
                        st.success("✅ ¡Todas las notas han sido guardadas exitosamente!")
                        st.rerun()
            with col_save2:
                st.caption("🟢 **Guardado en segundo plano activo**: Al escribir una nota y presionar `Enter` o cambiar de fila, se guarda de forma instantánea y liviana.")

            # --- BARRA DE DESCARGAS DINÁMICAS (XLSX Y CSV CON FILTROS Y ORDEN EXACTO) ---
            st.markdown("---")
            cant_export = len(df_edit_view)
            st.markdown(f"##### 📥 Opciones de Descarga Dinámica ({cant_export:,} consultoras filtradas)".replace(",", "."))
            st.caption("Exporta la base conservando exactamente los filtros aplicados y el orden visible en pantalla, con semáforos de color:")

            # Obtener lista ordenada de Códigos CB visibles en pantalla
            cbs_visibles = [str(x).strip() for x in df_edit_view['Código CB'] if pd.notna(x)] if 'Código CB' in df_edit_view.columns else None
            df_export_tab = obtener_base_tableau_completa_original(cbs_filtrados=cbs_visibles)

            # Construir nombre dinámico y descriptivo del archivo según los filtros activos
            partes_nombre = ["Base_Consultoras"]
            if lider_sel_t != "Todas las Líderes (Consolidado Zona)":
                g_slug = str(lider_sel_t).strip().replace(" ", "_")
                lider_nombre = mapa_lideres_tab.get(str(lider_sel_t).strip(), "")
                if lider_nombre:
                    lider_clean = "".join(c for c in lider_nombre if c.isalnum() or c == ' ').strip().replace(" ", "_")
                    partes_nombre.append(f"Gpo_{g_slug}_{lider_clean[:15]}")
                else:
                    partes_nombre.append(f"Gpo_{g_slug}")
            if sits_sel:
                sits_slug = "_".join([str(s).replace(" ", "") for s in sits_sel[:2]])
                partes_nombre.append(sits_slug)
            if colores_tab_sel:
                col_slug = "_".join([str(c) for c in colores_tab_sel[:2]])
                partes_nombre.append(col_slug)
            if f_mora_opt != "Todas":
                partes_nombre.append("ConMora" if "Con" in f_mora_opt else "SinMora")
            if f_ped_opt != "Todos":
                partes_nombre.append("ConPedidos")
            if busq_t.strip():
                partes_nombre.append("Busqueda")

            nombre_base_archivo = "_".join(partes_nombre)

            cdown1, cdown2, cdown3 = st.columns(3)

            with cdown1:
                excel_colores_bytes = exportar_tableau_excel_con_colores(df_export_tab, nombre_hoja="Base_Completa")
                st.download_button(
                    label="📗 Excel Completo a Colores (60+ Cols)",
                    data=excel_colores_bytes,
                    file_name=f"{nombre_base_archivo}_Completa.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with cdown2:
                excel_gestion_bytes = exportar_tableau_excel_con_colores(df_edit_view, nombre_hoja="Vista_Gestion")
                st.download_button(
                    label="📑 Excel Vista Gestión (16 Cols)",
                    data=excel_gestion_bytes,
                    file_name=f"{nombre_base_archivo}_Gestion.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with cdown3:
                csv_tab_bytes = df_export_tab.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📊 Archivo CSV (.csv)",
                    data=csv_tab_bytes,
                    file_name=f"{nombre_base_archivo}.csv",
                    mime="text/csv",
                    use_container_width=True
                )

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

                    # Botón de enlace directo a WhatsApp
                    tel_p = str(row_p.get('celular', '')).strip().replace(' ', '').replace('-', '').replace('+', '')
                    if tel_p and len(tel_p) >= 10:
                        url_wa_p = f"https://api.whatsapp.com/send?phone=57{tel_p}&text={urllib.parse.quote(msg_pago)}"
                        st.markdown(f"""
                        <div style="margin-top: 8px; margin-bottom: 12px;">
                            <a href="{url_wa_p}" target="_blank" style="display: inline-block; background-color: #25D366; color: white; padding: 9px 18px; border-radius: 8px; text-decoration: none; font-weight: bold; font-size: 14px; box-shadow: 0 4px 6px rgba(0,0,0,0.12);">
                                📲 Abrir WhatsApp y Enviar Mensaje a {nombre_p}
                            </a>
                        </div>
                        """, unsafe_allow_html=True)

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

        # --- SUBPESTAÑA 4: ASISTENTE & CAMPAÑAS WHATSAPP ---
        with tab_tab_whatsapp:
            st.markdown("##### 📲 Centro de Gestión & Campañas de WhatsApp")
            st.caption("Genera mensajes comerciales hiper-personalizados y envíalos de inmediato a través de WhatsApp Web / Móvil o expórtalos para plataformas masivas.")

            col_camp1, col_camp2 = st.columns([1.5, 2.5])
            with col_camp1:
                tipo_camp = st.selectbox(
                    "🎯 Tipo de Campaña / Objetivo:",
                    options=[
                        "💳 1. Cobro de Cartera en Mora",
                        "⌛ 2. Liberación de Pedidos Retenidos",
                        "🌟 3. Impulso de Puntos & Ascenso de Nivel",
                        "🌸 4. Reactivación de Consultoras Inactivas",
                        "✍️ 5. Mensaje Libre / Personalizado"
                    ],
                    key="tipo_camp_sel_tab"
                )

                # Segmentación automática según tipo de campaña
                if "1. Cobro" in tipo_camp:
                    df_wa_target = df_tab_filt[(df_tab_filt['Deuda Mora'] > 0) | (df_tab_filt['Deuda Total'] > 0)].copy()
                    plantilla_def = (
                        "Hola *{primer_nombre}* 🌸, te saluda tu Líder de *Natura & Avon*.\n\n"
                        "Queremos recordarte que tienes un saldo pendiente de *{deuda_mora}* (Total: {deuda_total}).\n\n"
                        "Te invitamos a realizar tu pago hoy para liberar tu cupo de crédito y seguir disfrutando de tus beneficios. ¡Escríbeme si necesitas apoyo con tu código de pago! 💳✨"
                    )
                elif "2. Liberación" in tipo_camp:
                    df_wa_target = df_tab_filt[(df_tab_filt['Ped. Pendientes'] > 0) | (df_tab_filt['Ped. Mora'] > 0)].copy()
                    plantilla_def = (
                        "Hola *{primer_nombre}* 🛍️, te saluda tu Líder de *Natura & Avon*.\n\n"
                        "Tienes *{pedidos} pedido(s)* en espera de liberación por saldo de *{deuda_mora}*.\n\n"
                        "Al poner al día tu pago hoy, tu pedido saldrá de inmediato para despacho. ¡Quedo atenta para ayudarte! 📦✨"
                    )
                elif "3. Impulso" in tipo_camp:
                    df_wa_target = df_tab_filt[(df_tab_filt['Pts Acum'] > 0) | (df_tab_filt['Pts Asc'] > 0)].copy()
                    plantilla_def = (
                        "¡Hola *{primer_nombre}*! 🌟 Felicitaciones por tus *{pts_acum} puntos* acumulados en tu nivel *{nivel}*.\n\n"
                        "Estás muy cerca de tu siguiente meta de premios y beneficios exclusivos de este ciclo. ¡Pasa tu pedido y gana más con Natura & Avon! 🎁✨"
                    )
                elif "4. Reactivación" in tipo_camp:
                    df_wa_target = df_tab_filt[df_tab_filt['Sit. Comercial'].astype(str).str.contains('Inactiva', case=False, na=False)].copy()
                    plantilla_def = (
                        "¡Hola *{primer_nombre}*! 🌸 Te extrañamos mucho en nuestro grupo de *Natura & Avon*.\n\n"
                        "En este ciclo tenemos promociones exclusivas, descuentos especiales y kits de reinicio pensados para ti.\n\n"
                        "¿Te gustaría que te comparta el catálogo virtual interactivo de este ciclo? 📖✨"
                    )
                else:
                    df_wa_target = df_tab_filt.copy()
                    plantilla_def = (
                        "Hola *{primer_nombre}* 🌸, te saluda tu Líder de *Natura & Avon*.\n\n"
                        "Queremos desearte muchos éxitos en este ciclo. ¡Cuenta con nosotras para tus pedidos y metas comerciales! ✨"
                    )

                st.metric("👥 Consultoras en este Segmento", f"{len(df_wa_target):,} personas".replace(",", "."))

            with col_camp2:
                plantilla_txt = st.text_area(
                    "✏️ Plantilla del Mensaje (Variables: `{primer_nombre}`, `{nombre}`, `{deuda_mora}`, `{deuda_total}`, `{pedidos}`, `{nivel}`, `{pts_acum}`, `{credito_disp}`):",
                    value=plantilla_def,
                    height=160,
                    key=f"plantilla_txt_{tipo_camp[:2]}"
                )

            st.markdown("---")

            if df_wa_target.empty:
                st.info("ℹ️ No hay consultoras que cumplan con el criterio del segmento seleccionado.")
            else:
                st.markdown(f"###### 📋 Listado de Contacto para Campaña ({len(df_wa_target)} Asesoras)")
                st.caption("Haz clic en el enlace verde de WhatsApp de cada fila para abrir el chat instantáneo con el mensaje ya escrito:")

                # Construir tabla con enlaces directos de WhatsApp
                filas_wa = []
                for idx, r in df_wa_target.iterrows():
                    nom_full = str(r.get('Nombre', r.get('Asesora / Consultora', ''))).strip()
                    primer_n = nom_full.split()[0].title() if nom_full else "Consultora"
                    cel = str(r.get('celular', '')).strip().replace(' ', '').replace('-', '').replace('+', '')
                    
                    deuda_m = formato_cop(r.get('Deuda Mora', 0))
                    deuda_t = formato_cop(r.get('Deuda Total', 0))
                    cred_d = formato_cop(r.get('Credito Disponible', 0))
                    ped_val = int(limpiar_numero(r.get('Ped. Pendientes', 0)))
                    pts_val = int(limpiar_numero(r.get('Pts Acum', 0)))
                    col_nivel = str(r.get('Color', r.get('Nivel / Color', 'Consultora')))

                    # Reemplazar variables en plantilla
                    msg_personalizado = (
                        plantilla_txt
                        .replace("{primer_nombre}", primer_n)
                        .replace("{nombre}", nom_full.title())
                        .replace("{deuda_mora}", deuda_m)
                        .replace("{deuda_total}", deuda_t)
                        .replace("{credito_disp}", cred_d)
                        .replace("{pedidos}", str(ped_val))
                        .replace("{pts_acum}", str(pts_val))
                        .replace("{nivel}", col_nivel)
                    )

                    link_wa = f"https://api.whatsapp.com/send?phone=57{cel}&text={urllib.parse.quote(msg_personalizado)}" if cel and len(cel) >= 10 else ""

                    filas_wa.append({
                        'Código CB': str(r.get('Codigo CB', r.get('Código CB', ''))),
                        'Asesora': nom_full,
                        'Grupo': str(r.get('Grupo', '')),
                        'Celular': cel if cel else "Sin registrar",
                        'Sit. Comercial': str(r.get('Sit. Comercial', '')),
                        'Deuda Mora': deuda_m,
                        'Ped. Pendientes': ped_val,
                        'Mensaje Generado': msg_personalizado,
                        'Enlace WhatsApp': link_wa
                    })

                df_wa_table = pd.DataFrame(filas_wa)

                # Mostrar con column_config LinkColumn
                st.dataframe(
                    df_wa_table[['Código CB', 'Asesora', 'Grupo', 'Celular', 'Sit. Comercial', 'Deuda Mora', 'Ped. Pendientes', 'Enlace WhatsApp']],
                    column_config={
                        "Enlace WhatsApp": st.column_config.LinkColumn(
                            "📲 Chat WhatsApp",
                            help="Haz clic para abrir WhatsApp Web o App con el mensaje listo",
                            display_text="📲 Enviar WA"
                        )
                    },
                    use_container_width=True,
                    hide_index=True
                )

                # Botón de Descarga Masiva para plataformas (UltraMsg / Evolution API / Python)
                col_exp1, col_exp2 = st.columns([1.5, 2.5])
                with col_exp1:
                    csv_wa_bytes = df_wa_table.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        label="📥 Exportar Base para Envíos Masivos (CSV)",
                        data=csv_wa_bytes,
                        file_name=f"Campana_WhatsApp_{tipo_camp[:6].strip().replace(' ', '_')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                with col_exp2:
                    st.caption("💡 **Tip de Productividad**: Puedes usar este archivo CSV con herramientas como UltraMsg, Evolution API o Meta Cloud API para despachar cientos de mensajes en segundos sin riesgo de baneo.")

        # --- SUBPESTAÑA 5: CUMPLEAÑOS Y RECONOCIMIENTO ---
        with tab_tab_cumple:
            st.subheader("🎂 Calendario & Reconocimiento de Cumpleaños")
            st.markdown("Seguimiento de fechas especiales para fortalecer el vínculo comercial y humano con las asesoras de tu red.")
            renderizar_banner_cumpleanos(df_tab_filt, user_rol, user_nombre, user_grupo, user_sector, key_suffix="tableau_tab")

st.markdown("---")

# --- TAB GERAL: CRÉDITO & COBRANZA PREVENTIVA Y CARTERA ("Geral_Credito&Cobranza") ---
with tab_geral:
    st.subheader("💳 Geral: Crédito & Cobranza Inteligente")
    st.markdown("Control dinámico de cartera Natura & Avon, alertas de vencimiento preventivo (*Mañana*, *Pasado Mañana*), semáforo de mora y despachador de WhatsApp con 1 clic.")

    # 1. EXPANDER DE CARGA DINÁMICA DE ARCHIVO Geral.xlsx
    with st.expander("📤 Cargar / Actualizar Archivo 'Geral.xlsx'", expanded=False):
        st.markdown("##### 📂 Subida de Archivo Maestro de Crédito & Cobranza")
        st.caption("Carga el archivo Excel descargado del sistema oficial para actualizar las deudas, facturas y fechas de vencimiento de las asesoras.")
        
        file_geral_subido = st.file_uploader(
            "Selecciona o arrastra el archivo Geral.xlsx:",
            type=["xlsx", "xls"],
            key="uploader_geral_excel"
        )
        
        if file_geral_subido is not None:
            col_u_g1, col_u_g2 = st.columns([2, 1])
            with col_u_g1:
                st.info(f"📄 Archivo seleccionado: **{file_geral_subido.name}** ({file_geral_subido.size / 1024:.1f} KB)")
            with col_u_g2:
                btn_procesar_geral = st.button("🚀 Procesar & Sincronizar Base Geral", type="primary", use_container_width=True, key="btn_proc_geral_up")
                
            if btn_procesar_geral:
                with st.spinner("⏳ Analizando estructura y sincronizando registros en SQLite..."):
                    sec_para_validar = user_sector if user_rol == 'gerente' else None
                    ok_g, num_g, msg_g = sincronizar_excel_geral_a_sqlite(file_geral_subido, sector_esperado=sec_para_validar)
                    if ok_g:
                        st.success(f"✅ {msg_g}")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error(f"❌ Error al procesar: {msg_g}")

    # 2. CONSULTA DE DATOS DESDE SQLITE CON AUTO-RECUPERACIÓN
    sec_filtro_g = user_sector if user_rol == 'gerente' else None
    grp_filtro_g = user_grupo if user_rol == 'lider' else None
    
    df_geral_raw = consultar_geral_sql(grupo=grp_filtro_g, sector=sec_filtro_g)
    
    # Si SQLite está vacío pero existe Geral.xlsx local en disco, sincronizar automáticamente
    if (df_geral_raw is None or df_geral_raw.empty) and os.path.exists("Geral.xlsx"):
        sincronizar_excel_geral_a_sqlite("Geral.xlsx")
        df_geral_raw = consultar_geral_sql(grupo=grp_filtro_g, sector=sec_filtro_g)
        
    if df_geral_raw is None or df_geral_raw.empty:
        st.info("ℹ️ No hay registros de crédito y cobranza en el sistema. Por favor sube el archivo **Geral.xlsx** en el botón de arriba.")
    else:
        # Procesar análisis financiero
        analisis_g = procesar_analisis_geral_cobranza(df_geral_raw)
        kpis_g = analisis_g['kpis']
        df_pendientes = analisis_g['df_pendientes']
        df_manana = analisis_g['df_vence_manana']
        df_pasado = analisis_g['df_pasado_manana']
        df_mora = analisis_g['df_en_mora']
        df_7d = analisis_g['df_proximos_7d']
        heatmap_df = analisis_g['heatmap_data']
        
        # 3. FILA DE KPIS EJECUTIVOS FINANCIEROS
        kg1, kg2, kg3, kg4, kg5 = st.columns(5)
        with kg1:
            st.metric(
                "💰 Cartera Total Viva",
                f"${kpis_g['total_cartera']/1e6:.2f}M COP" if kpis_g['total_cartera'] >= 1e6 else f"${kpis_g['total_cartera']:,.0f}".replace(",", "."),
                f"{kpis_g['total_facturas_pendientes']} Facturas"
            )
        with kg2:
            st.metric(
                "🟡 Vence MAÑANA",
                f"${kpis_g['total_manana']/1e6:.2f}M COP" if kpis_g['total_manana'] >= 1e6 else f"${kpis_g['total_manana']:,.0f}".replace(",", "."),
                f"{kpis_g['facturas_manana']} Facturas (Prioridad)",
                delta_color="normal"
            )
        with kg3:
            st.metric(
                "🟢 Próximos 7 Días",
                f"${kpis_g['total_7d']/1e6:.2f}M COP" if kpis_g['total_7d'] >= 1e6 else f"${kpis_g['total_7d']:,.0f}".replace(",", "."),
                f"{kpis_g['facturas_7d']} Facturas"
            )
        with kg4:
            st.metric(
                "🚨 Cartera en Mora",
                f"${kpis_g['total_mora']/1e6:.2f}M COP" if kpis_g['total_mora'] >= 1e6 else f"${kpis_g['total_mora']:,.0f}".replace(",", "."),
                f"{kpis_g['facturas_mora']} Vencidas",
                delta_color="inverse"
            )
        with kg5:
            st.metric(
                "👥 Consultoras con Deuda",
                f"{kpis_g['consultoras_unicas']}",
                f"{len(df_geral_raw[df_geral_raw['situacion'] == 'Pagado'])} Pagadas / Al Día"
            )

        st.markdown("---")

        # 4. CRONOGRAMA DE PROYECCIÓN Y VISTAS RÁPIDAS (ORDENADO DE MAYOR A MENOR DEUDA Y CON COLORES ARMÓNICOS)
        st.markdown("##### 📋 Cronograma de Cartera & Gestión por Tramo de Vencimiento")
        st.caption(r"Ordenado de mayor a menor deuda con semáforo armónico: 🔴 **Deuda Alta / Mora** (>= \$300.000 COP) | 🟠 **Deuda Media** (\$150.000 - \$300.000 COP) | 🟢 **Deuda Controlada** (< \$150.000 COP)")

        tab_v_manana, tab_v_pasado, tab_v_mora, tab_v_7d, tab_v_todas, tab_v_pagadas = st.tabs([
            f"🟡 Vencen Mañana ({len(df_manana)})",
            f"🟢 Pasado Mañana (+2 a +3d) ({len(df_pasado)})",
            f"🚨 En Mora ({len(df_mora)})",
            f"📅 Próximos 7 Días ({len(df_7d)})",
            f"🗓️ Todas las Pendientes ({len(df_pendientes)})",
            f"✅ Historial Pagados ({len(df_geral_raw[df_geral_raw['situacion'] == 'Pagado'])})"
        ])

        def _limpiar_texto_plan_pago(val):
            val_str = str(val or '').strip()
            if 'plan_recibimiento' in val_str:
                parts = val_str.split('plan_recibimiento')
                if len(parts) > 1:
                    clean = parts[1].strip()
                    for tok in ['26', '36', '47', 'Name:']:
                        clean = clean.replace(tok, '').strip()
                    return clean[:35].strip()
            return val_str[:35].strip() if val_str else "Estándar"

        def _color_saldo_total_armonico(val):
            try:
                num = float(limpiar_numero(val))
            except Exception:
                num = 0.0
            if num >= 300000:
                return 'background-color: rgba(239, 68, 68, 0.18); color: #DC2626; font-weight: 700;'
            elif num >= 150000:
                return 'background-color: rgba(245, 158, 11, 0.18); color: #D97706; font-weight: 700;'
            else:
                return 'background-color: rgba(16, 185, 129, 0.15); color: #059669; font-weight: 700;'

        def _color_nivel_deuda_badge(val):
            val_str = str(val)
            if 'Alta' in val_str or 'Mora' in val_str:
                return 'background-color: rgba(239, 68, 68, 0.20); color: #DC2626; font-weight: 700; border-radius: 4px;'
            elif 'Media' in val_str:
                return 'background-color: rgba(245, 158, 11, 0.20); color: #D97706; font-weight: 700; border-radius: 4px;'
            elif 'Pagado' in val_str or 'Al Día' in val_str:
                return 'background-color: rgba(59, 130, 246, 0.18); color: #2563EB; font-weight: 700; border-radius: 4px;'
            else:
                return 'background-color: rgba(16, 185, 129, 0.18); color: #059669; font-weight: 700; border-radius: 4px;'

        def _color_dias_restantes_badge(val):
            val_str = str(val)
            if 'Mora' in val_str or '-' in val_str or '🔴' in val_str:
                return 'background-color: rgba(239, 68, 68, 0.20); color: #DC2626; font-weight: 700;'
            elif 'Hoy' in val_str or '🚨' in val_str:
                return 'background-color: rgba(249, 115, 22, 0.20); color: #EA580C; font-weight: 700;'
            elif 'Mañana' in val_str or '🟡' in val_str:
                return 'background-color: rgba(234, 179, 8, 0.20); color: #D97706; font-weight: 700;'
            elif 'Pagado' in val_str or '✅' in val_str:
                return 'background-color: rgba(16, 185, 129, 0.18); color: #059669; font-weight: 700;'
            else:
                return 'background-color: rgba(16, 185, 129, 0.15); color: #059669; font-weight: 600;'

        def _formatear_tabla_geral(df_in):
            if df_in is None or df_in.empty:
                st.info("🎉 No hay facturas en esta categoría actualmente.")
                return None

            # 1. Ordenar de mayor a menor deuda de forma segura
            cols_sort = [c for c in ['saldo_total', 'dias_para_vencer'] if c in df_in.columns]
            if cols_sort:
                df_ordenado = df_in.sort_values(by=cols_sort, ascending=[False] * len(cols_sort)).copy()
            else:
                df_ordenado = df_in.copy()

            # 2. Asignar Nivel de Deuda
            def calcular_etiqueta_nivel(r):
                sit = str(r.get('situacion', '')).strip()
                if sit.lower() == 'pagado':
                    return "✅ Pagado / Al Día"
                s = float(limpiar_numero(r.get('saldo_total', 0)))
                d = float(limpiar_numero(r.get('dias_para_vencer', 0))) if 'dias_para_vencer' in r else 0.0
                if d < 0:
                    return "🔴 En Mora"
                elif s >= 300000:
                    return "🔴 Alta"
                elif s >= 150000:
                    return "🟠 Media"
                else:
                    return "🟢 Controlada"

            df_ordenado['Nivel Deuda'] = df_ordenado.apply(calcular_etiqueta_nivel, axis=1)

            # 3. Limpiar Días Restantes
            def formatear_dias_legibles(r):
                sit = str(r.get('situacion', '')).strip()
                if sit.lower() == 'pagado':
                    return "✅ Al Día"
                if 'dias_para_vencer' in r:
                    d = int(limpiar_numero(r.get('dias_para_vencer', 0)))
                    if d < 0:
                        return f"🔴 {abs(d)} d. mora"
                    elif d == 0:
                        return "🚨 Vence Hoy"
                    elif d == 1:
                        return "🟡 Vence Mañana"
                    elif 2 <= d <= 3:
                        return f"🟢 En {d} días"
                    else:
                        return f"📅 En {d} días"
                else:
                    return "📅 Programado"

            df_ordenado['Estado Vencimiento'] = df_ordenado.apply(formatear_dias_legibles, axis=1)

            # 4. Limpiar Plan de Pago
            if 'plan_recibimiento' in df_ordenado.columns:
                df_ordenado['plan_recibimiento'] = df_ordenado['plan_recibimiento'].apply(_limpiar_texto_plan_pago)

            cols_mostrar = [
                'nombre', 'codigo_cb', 'grupo', 'numero_factura', 'fecha_vencimiento',
                'Estado Vencimiento', 'Nivel Deuda', 'saldo_principal', 'saldo_financiero',
                'saldo_total', 'plan_recibimiento', 'telefono_movil', 'telefono_movil_2'
            ]
            cols_existentes = [c for c in cols_mostrar if c in df_ordenado.columns]
            df_disp = df_ordenado[cols_existentes].copy()

            rename_dict = {
                'nombre': 'Consultora',
                'codigo_cb': 'Código CB',
                'grupo': 'Grupo',
                'numero_factura': 'Factura',
                'fecha_vencimiento': 'Vencimiento',
                'saldo_principal': 'Saldo Capital',
                'saldo_financiero': 'Saldo Financiero',
                'saldo_total': 'Saldo Total',
                'plan_recibimiento': 'Plan de Pago',
                'telefono_movil': 'Celular',
                'telefono_movil_2': 'Movil 2'
            }
            df_disp = df_disp.rename(columns=rename_dict)

            format_dict = {}
            if 'Saldo Capital' in df_disp.columns:
                format_dict['Saldo Capital'] = lambda v: f"${v:,.0f} COP".replace(",", ".")
            if 'Saldo Financiero' in df_disp.columns:
                format_dict['Saldo Financiero'] = lambda v: f"${v:,.0f} COP".replace(",", ".")
            if 'Saldo Total' in df_disp.columns:
                format_dict['Saldo Total'] = lambda v: f"${v:,.0f} COP".replace(",", ".")

            styler = df_disp.style.format(format_dict)

            if 'Saldo Total' in df_disp.columns:
                styler = styler.map(_color_saldo_total_armonico, subset=['Saldo Total'])
            if 'Nivel Deuda' in df_disp.columns:
                styler = styler.map(_color_nivel_deuda_badge, subset=['Nivel Deuda'])
            if 'Estado Vencimiento' in df_disp.columns:
                styler = styler.map(_color_dias_restantes_badge, subset=['Estado Vencimiento'])

            return styler

        with tab_v_manana:
            st.markdown("###### 🟡 Facturas que Vencen Mañana (Recordatorio Preventivo)")
            st.caption("Envía el recordatorio preventivo con mucho cariño para que la consultora pague a tiempo y mantenga su crédito activo.")
            df_m_disp = _formatear_tabla_geral(df_manana)
            if df_m_disp is not None:
                st.dataframe(df_m_disp, use_container_width=True, hide_index=True)

        with tab_v_pasado:
            st.markdown("###### 🟢 Facturas que Vencen Pasado Mañana (+2 y +3 Días)")
            df_p_disp = _formatear_tabla_geral(df_pasado)
            if df_p_disp is not None:
                st.dataframe(df_p_disp, use_container_width=True, hide_index=True)

        with tab_v_mora:
            st.markdown("###### 🚨 Cartera Vencida (En Mora)")
            st.caption("Asesoras con días de retraso vencidos. Incluye cobro de recargos financieros acumulados.")
            df_mo_disp = _formatear_tabla_geral(df_mora)
            if df_mo_disp is not None:
                st.dataframe(df_mo_disp, use_container_width=True, hide_index=True)

        with tab_v_7d:
            st.markdown("###### 📅 Proyección de Vencimientos en los Próximos 7 Días")
            df_7d_disp = _formatear_tabla_geral(df_7d)
            if df_7d_disp is not None:
                st.dataframe(df_7d_disp, use_container_width=True, hide_index=True)

        with tab_v_todas:
            st.markdown("###### 🗓️ Todas las Facturas Pendientes de Cobro")
            df_t_disp = _formatear_tabla_geral(df_pendientes)
            if df_t_disp is not None:
                st.dataframe(df_t_disp, use_container_width=True, hide_index=True)

        with tab_v_pagadas:
            st.markdown("###### ✅ Títulos Pagados y Conciliados (Excluidos de Cartera)")
            df_pagados = df_geral_raw[df_geral_raw['situacion'] == 'Pagado']
            df_pag_disp = _formatear_tabla_geral(df_pagados)
            if df_pag_disp is not None:
                st.dataframe(df_pag_disp, use_container_width=True, hide_index=True)

        st.markdown("---")

        # 5. SELECCIÓN MÚLTIPLE & DESPACHADOR MASIVO DE WHATSAPP
        st.markdown("##### 📢 Envío Masivo & Recordatorios de Cobranza por WhatsApp")
        st.caption("Selecciona una, varias o todas las consultoras de un tramo para generar sus mensajes personalizados en lote y enviarlos con 1 clic o conectarte con un API.")
        
        # Botones de Carga Rápida de Lotes
        col_btn_m1, col_btn_m2, col_btn_m3, col_btn_m4, col_btn_m5 = st.columns(5)
        
        if 'titulos_seleccionados_masivo' not in st.session_state:
            st.session_state['titulos_seleccionados_masivo'] = df_manana['titulo'].tolist() if not df_manana.empty else []
            
        with col_btn_m1:
            if st.button("🟡 Vencen Mañana", use_container_width=True):
                st.session_state['titulos_seleccionados_masivo'] = df_manana['titulo'].tolist()
                st.rerun()
        with col_btn_m2:
            if st.button("🚨 En Mora", use_container_width=True):
                st.session_state['titulos_seleccionados_masivo'] = df_mora['titulo'].tolist()
                st.rerun()
        with col_btn_m3:
            if st.button("🟢 Pasado Mañana", use_container_width=True):
                st.session_state['titulos_seleccionados_masivo'] = df_pasado['titulo'].tolist()
                st.rerun()
        with col_btn_m4:
            if st.button("📅 Próximos 7 Días", use_container_width=True):
                st.session_state['titulos_seleccionados_masivo'] = df_7d['titulo'].tolist()
                st.rerun()
        with col_btn_m5:
            if st.button("🧹 Limpiar Todo", use_container_width=True):
                st.session_state['titulos_seleccionados_masivo'] = []
                st.rerun()

        # Opciones para el multiselect
        mapa_titulos_dict = {
            row['titulo']: f"{row['nombre']} — Fact. {row['numero_factura']} (${row['saldo_total']:,.0f}) [Días: {row['dias_para_vencer']}]"
            for _, row in df_pendientes.iterrows()
        }
        
        # Multiselect de asesoras
        sel_titulos_activos = st.multiselect(
            "👥 Asesoras Seleccionadas para la Campaña:",
            options=list(mapa_titulos_dict.keys()),
            default=[t for t in st.session_state['titulos_seleccionados_masivo'] if t in mapa_titulos_dict],
            format_func=lambda t: mapa_titulos_dict.get(t, t),
            key="multiselect_geral_masivo"
        )
        
        st.session_state['titulos_seleccionados_masivo'] = sel_titulos_activos
        
        if sel_titulos_activos:
            df_target_masivo = df_pendientes[df_pendientes['titulo'].isin(sel_titulos_activos)].copy()
            st.info(f"🎯 **{len(df_target_masivo)} asesora(s) seleccionada(s)** — Monto total de campaña: **${df_target_masivo['saldo_total'].sum():,.0f} COP**".replace(",", "."))
            
            col_cfg1, col_cfg2 = st.columns([1.2, 1.8])
            with col_cfg1:
                tipo_camp_sel = st.selectbox(
                    "Tipo de Plantilla:",
                    options=['auto', 'manana', 'hoy', 'mora', 'general'],
                    format_func=lambda x: {
                        'auto': '⚡ Automático (Detecta si vence mañana, hoy o mora)',
                        'manana': '🎁 Vence Mañana (Preventivo Cordial)',
                        'hoy': '🚨 Vence Hoy (Urgente sin recargos)',
                        'mora': '⚠️ En Mora (Cobranza con Recargos)',
                        'general': '🌸 Recordatorio General'
                    }.get(x, x),
                    key="sel_tipo_camp_masivo"
                )
                nombre_remit_masivo = st.text_input("Nombre de la Líder / Remitente:", value=user_nombre if user_nombre else "Tu Líder", key="in_remit_masivo")

            with col_cfg2:
                st.caption("Variables que se reemplazan en cada mensaje: `{primer_nombre}`, `{nombre}`, `{factura}`, `{saldo_total}`, `{vencimiento}`")
                
            # Generar tabla de mensajes
            filas_campana = []
            for _, r in df_target_masivo.iterrows():
                t_msg = tipo_camp_sel
                if t_msg == 'auto':
                    d_r = r['dias_para_vencer']
                    t_msg = 'manana' if d_r == 1 else ('hoy' if d_r == 0 else ('mora' if d_r < 0 else 'general'))
                    
                msg_ind = generar_mensaje_whatsapp_cobranza(r, tipo=t_msg, nombre_remitente=nombre_remit_masivo)
                cel = str(r.get('telefono_movil', '')).strip()
                cel2 = str(r.get('telefono_movil_2', '')).strip()
                
                link_w1 = f"https://api.whatsapp.com/send?phone=57{cel}&text={urllib.parse.quote(msg_ind)}" if cel and len(cel) >= 10 else ""
                link_w2 = f"https://api.whatsapp.com/send?phone=57{cel2}&text={urllib.parse.quote(msg_ind)}" if cel2 and len(cel2) >= 10 else ""
                
                filas_campana.append({
                    'Asesora': r.get('nombre'),
                    'Código CB': r.get('codigo_cb'),
                    'Grupo': r.get('grupo'),
                    'Celular': cel if cel else "Sin celular",
                    'Movil 2': cel2 if cel2 else "",
                    'Factura': r.get('numero_factura'),
                    'Vencimiento': r.get('fecha_vencimiento'),
                    'Días Restantes': r.get('dias_para_vencer'),
                    'Saldo Total': f"${r.get('saldo_total'):,.0f} COP".replace(",", "."),
                    'Mensaje Personalizado': msg_ind,
                    'Enlace Directo': link_w1,
                    'Enlace Directo WhatsApp': link_w1,
                    'Enlace Móvil 2': link_w2
                })
                
            df_campana_out = pd.DataFrame(filas_campana)
            
            # Vista previa del lote
            st.dataframe(
                df_campana_out[['Asesora', 'Grupo', 'Celular', 'Factura', 'Vencimiento', 'Días Restantes', 'Saldo Total', 'Mensaje Personalizado']],
                use_container_width=True,
                hide_index=True
            )
            
            col_d1, col_d2 = st.columns([1.5, 1.5])
            with col_d1:
                # Selector individual rápido dentro del lote
                st.markdown("###### 📲 Despachar Asesora Individual:")
                nom_sel_rapido = st.selectbox("Elige la asesora para enviar de inmediato:", options=df_campana_out['Asesora'].tolist(), key="sel_rapido_camp")
                row_sel_rap = df_campana_out[df_campana_out['Asesora'] == nom_sel_rapido].iloc[0]
                link_wa_enviar = row_sel_rap.get('Enlace Directo') or row_sel_rap.get('Enlace Directo WhatsApp')
                if link_wa_enviar:
                    st.link_button(f"📲 Abrir WhatsApp y Enviar a {str(nom_sel_rapido).split()[0].title()}", url=link_wa_enviar, use_container_width=True)
                else:
                    st.warning("⚠️ Esta asesora no tiene un número celular válido de 10 dígitos registrado.")
                    
            with col_d2:
                st.markdown("###### 📥 Descargar Base de Campaña:")
                csv_camp = df_campana_out.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label=f"📥 Descargar Campaña CSV ({len(df_campana_out)} Mensajes)",
                    data=csv_camp,
                    file_name=f"Campana_Cobranza_WA_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key="btn_descargar_camp_wa"
                )

            # 6. INTEGRACIÓN Y PASARELAS PARA ENVÍOS AUTOMÁTICOS
            with st.expander("🔌 Integración & Conexión con Pasarelas de WhatsApp (Envío Automático)", expanded=False):
                st.markdown("##### 🚀 Pasarela de Envíos Masivos Automáticos")
                st.markdown("""
                Para enviar mensajes masivos a cientos de asesoras sin tocar tu teléfono 1 a 1, te recomendamos conectar una **API de WhatsApp**:
                
                * 👑 **Evolution API (Recomendada - 100% Gratuita & Open Source)**:
                  * Puedes montarla en Railway o Docker en 2 minutos.
                  * Escaneas el código QR con el WhatsApp de la Líder o Gerente.
                  * Envía mensajes con variables, texto en negrita y emojis sin coste por mensaje.
                * 🌐 **UltraMsg / Wassenger / Z-API (Cloud SaaS)**:
                  * Plataforma en la nube lista para usar con API Key.
                * 🏢 **WhatsApp Cloud API Oficial (Meta Graph API)**:
                  * La solución corporativa de Meta para grandes volúmenes.
                """)
                
                st.markdown("###### ⚙️ Despachador API en Vivo:")
                col_api1, col_api2 = st.columns(2)
                with col_api1:
                    api_url_in = st.text_input("Endpoint / URL de la API:", placeholder="ej. https://mi-evolution-api.up.railway.app/message/sendText/mi_instancia", key="in_api_url_geral")
                with col_api2:
                    api_token_in = st.text_input("API Key / Bearer Token:", type="password", placeholder="ej. B6D711FCDE4D4FD5936544120E713976", key="in_api_token_geral")
                    
                btn_disparar_api = st.button(f"🚀 Iniciar Envío Automático a las {len(df_campana_out)} Asesoras", type="primary", use_container_width=True, key="btn_disparar_api_geral")
                
                if btn_disparar_api:
                    if not api_url_in.strip() or not api_token_in.strip():
                        st.warning("⚠️ Ingresa la URL del Endpoint y el Token de tu API para iniciar el envío automático.")
                    else:
                        import requests
                        progress_bar = st.progress(0.0)
                        status_txt = st.empty()
                        enviados_ok = 0
                        errores_cnt = 0
                        
                        for i, r_c in enumerate(df_campana_out.iterrows()):
                            r_c = r_c[1]
                            cel_num = str(r_c['Celular']).strip()
                            if cel_num and len(cel_num) >= 10:
                                payload = {
                                    "number": f"57{cel_num}",
                                    "text": r_c['Mensaje Personalizado'],
                                    "body": r_c['Mensaje Personalizado']
                                }
                                headers = {
                                    "apikey": api_token_in.strip(),
                                    "Authorization": f"Bearer {api_token_in.strip()}",
                                    "Content-Type": "application/json"
                                }
                                try:
                                    res = requests.post(api_url_in.strip(), json=payload, headers=headers, timeout=10)
                                    if res.status_code in [200, 201]:
                                        enviados_ok += 1
                                    else:
                                        errores_cnt += 1
                                except Exception:
                                    errores_cnt += 1
                                    
                            progress_bar.progress((i + 1) / len(df_campana_out))
                            status_txt.caption(f"Despachando {i+1} de {len(df_campana_out)}: {r_c['Asesora']}...")
                            
                        st.success(f"✅ ¡Proceso finalizado! Enviados con éxito: {enviados_ok} | Fallidos: {errores_cnt}")
        else:
            st.info("👆 Selecciona al menos una asesora arriba o usa los botones de carga rápida para armar la campaña.")

        st.markdown("---")

        # 7. BOTÓN DE DESCARGA EXCEL / CSV DE CARTERA COMPLETA
        csv_geral_exp = df_pendientes.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 Descargar Base Completa de Cartera Pendiente (CSV / Excel)",
            data=csv_geral_exp,
            file_name=f"Cartera_Geral_Pendiente_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True,
            key="btn_descargar_cartera_geral_csv"
        )

st.markdown("---")

# --- TAB 1: RESUMEN Y KPIS ---
with tab_resumen:
    st.subheader("💎 Tablero Estadístico Interactivo & Métricas de Rendimiento")
    st.markdown("Visualizaciones ejecutivas dinámicas para el seguimiento del ciclo en tiempo real.")

    # 1. Tacómetros de Cumplimiento Global (Gauge Charts 360°)
    st.markdown("##### ⏱️ Tacómetros de Cumplimiento Global del Ciclo")
    cg1, cg2 = st.columns(2)
    with cg1:
        fig_g1 = crear_tacometro_360(
            "Cumplimiento Facturación",
            cump_fact,
            formato_cop(obj_fact),
            formato_cop(real_fact)
        )
        st.plotly_chart(fig_g1, use_container_width=True)
        
    with cg2:
        fig_g2 = crear_tacometro_360(
            "Cumplimiento Consultoras Activas",
            cump_activas,
            f"{int(obj_activas)} pers.",
            f"{int(real_activas)} pers."
        )
        st.plotly_chart(fig_g2, use_container_width=True)

    st.markdown("---")

    # 2. Fila de Gráficos Principales (Diferenciados por Perfil)
    if user_rol in ['superadmin', 'gerente']:
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            fig_rank = crear_ranking_lideres_fig(df_filtrado)
            if fig_rank:
                st.plotly_chart(fig_rank, use_container_width=True)
            else:
                st.info("No hay suficiente información para generar el ranking.")
                
        with col_g2:
            fig_dona = crear_dona_cartera_fig(df_tableau)
            if fig_dona:
                st.plotly_chart(fig_dona, use_container_width=True)
            else:
                st.info("No hay datos de cartera disponibles.")

        st.markdown("---")
        col_fun1, col_fun2 = st.columns(2)
        with col_fun1:
            disp_tot = int(df_filtrado['Disponibles'].sum()) if 'Disponibles' in df_filtrado.columns else 0
            fig_fun = crear_embudo_red_fig(disp_tot, inicios_totales, reinicios_totales, real_activas)
            st.plotly_chart(fig_fun, use_container_width=True)
            
        with col_fun2:
            fig_tree = crear_treemap_red_fig(df_tableau)
            if fig_tree:
                st.plotly_chart(fig_tree, use_container_width=True)

    else: # Vista para Líderes de Negocio
        col_l1, col_l2 = st.columns(2)
        with col_l1:
            fig_tree = crear_treemap_red_fig(df_tableau)
            if fig_tree:
                st.plotly_chart(fig_tree, use_container_width=True)
            else:
                st.info("No hay datos de niveles disponibles para tu equipo.")
                
        with col_l2:
            fig_scat = crear_scatter_atencion_fig(df_tableau)
            if fig_scat:
                st.plotly_chart(fig_scat, use_container_width=True)
            else:
                st.info("No hay datos de cartera para la matriz de atención.")

        st.markdown("---")
        col_lf1, col_lf2 = st.columns(2)
        with col_lf1:
            disp_tot = int(df_filtrado['Disponibles'].sum()) if 'Disponibles' in df_filtrado.columns else 0
            fig_fun = crear_embudo_red_fig(disp_tot, inicios_totales, reinicios_totales, real_activas)
            st.plotly_chart(fig_fun, use_container_width=True)
            
        with col_lf2:
            fig_dona = crear_dona_cartera_fig(df_tableau)
            if fig_dona:
                st.plotly_chart(fig_dona, use_container_width=True)

    st.markdown("---")
    st.markdown("##### 📍 Desempeño Resumido por Sector y Clasificación")
    col_left, col_right = st.columns([1, 1])
    with col_left:
        st.markdown("###### 📍 Desempeño por Sector")
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
        st.markdown("###### 🎨 Distribución por Clasificación / Color")
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
        val_f_real, val_f_obj, val_a_real, val_a_obj, val_inicios, val_saldo = 0.0, 0.0, 0, 0, 0, 0

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

    st.markdown("---")
    st.markdown("### 🌟 Simuladores Avanzados de Negocio & Proyección")
    st.caption("Módulos oficiales del Simulador LN dinamizados en tiempo real con la data de tu grupo comercial:")

    tab_sub_mentora, tab_sub_retencion, tab_sub_convencion = st.tabs([
        "🎓 Bono Líder Mentora",
        "🔄 Proyección de Retención & Actividad",
        "🏆 Puntos a Convención Natura"
    ])

    # 1. BONO LÍDER MENTORA
    with tab_sub_mentora:
        st.markdown("##### 🎓 Simulador Bono Líder Mentora (Período 202612 - 202618)")
        st.caption(r"Premia el crecimiento y consolidación de tu grupo con un bono de hasta **\$600.000 COP** sujeto al cumplimiento de dos habilitadores clave.")
        
        col_bm_in1, col_bm_in2, col_bm_in3 = st.columns(3)
        with col_bm_in1:
            bm_act_real = st.number_input("Activas Reales Logradas:", value=int(val_a_real), min_value=0, step=1, key=f"bm_act_r_{leader_key}")
            bm_act_obj = st.number_input("Activas Meta / Desafío:", value=int(val_a_obj if val_a_obj > 0 else 100), min_value=1, step=1, key=f"bm_act_o_{leader_key}")
        with col_bm_in2:
            bm_saldo_real = st.number_input("Saldo Comercial Real:", value=int(val_saldo), min_value=-50, max_value=100, step=1, key=f"bm_sal_r_{leader_key}")
            st.caption("💡 **Meta Saldo**: Debe ser mayor o igual a **2**.")
        with col_bm_in3:
            nom_lider_display = df_filtrado.iloc[0].get('Nombre de consultora', user_nombre) if not df_filtrado.empty else user_nombre
            st.info(f"👤 **Líder:** {nom_lider_display}\n\n🏷️ **Grupo:** {user_grupo if user_grupo else 'Seleccionado'}\n\n📅 **Período:** Ciclos 202612 - 202618")

        res_bm = calcular_bono_lider_mentora(bm_act_real, bm_act_obj, bm_saldo_real)
        
        col_bm_k1, col_bm_k2, col_bm_k3 = st.columns(3)
        with col_bm_k1:
            st.metric(
                "Habilitador 1: % Activas (Meta ≥ 95%)",
                f"{res_bm['pct_alcanzado_activas']*100:.1f}%",
                "✅ Cumple Habilitador" if res_bm['cumple_activas'] else "❌ No Cumple (<95%)",
                delta_color="normal" if res_bm['cumple_activas'] else "inverse"
            )
        with col_bm_k2:
            st.metric(
                "Habilitador 2: Saldo (Meta ≥ 2)",
                f"{bm_saldo_real}",
                "✅ Cumple Habilitador" if res_bm['cumple_saldo'] else "❌ No Cumple (<2)",
                delta_color="normal" if res_bm['cumple_saldo'] else "inverse"
            )
        with col_bm_k3:
            st.metric(
                "Bono Líder Mentora Estimado",
                f"${res_bm['bono_cop']:,.0f} COP".replace(",", "."),
                f"Escala: {res_bm['rango_activas']}"
            )
            
        if res_bm['cumple_ambos']:
            st.success(res_bm['mensaje_estado'])
        else:
            st.warning(res_bm['mensaje_estado'])

        with st.expander("📋 Ver Tabla Oficial de Escalas de Pago - Bono Mentora"):
            df_escalas_bm = pd.DataFrame([
                {"Rango de Activas": "Menos de 40", "Habilitador % Activas": "≥ 95%", "Habilitador Saldo": "≥ 2", "Bono Oficial": "$0 COP"},
                {"Rango de Activas": "40 a 59", "Habilitador % Activas": "≥ 95%", "Habilitador Saldo": "≥ 2", "Bono Oficial": "$300.000 COP"},
                {"Rango de Activas": "60 a 79", "Habilitador % Activas": "≥ 95%", "Habilitador Saldo": "≥ 2", "Bono Oficial": "$400.000 COP"},
                {"Rango de Activas": "80 a 99", "Habilitador % Activas": "≥ 95%", "Habilitador Saldo": "≥ 2", "Bono Oficial": "$500.000 COP"},
                {"Rango de Activas": "100 o más", "Habilitador % Activas": "≥ 95%", "Habilitador Saldo": "≥ 2", "Bono Oficial": "$600.000 COP"}
            ])
            st.table(df_escalas_bm)

    # 2. PROYECCIÓN DE RETENCIÓN & ACTIVIDAD
    with tab_sub_retencion:
        st.markdown("##### 🔄 Proyección de Retención & Movimiento de Base (Ciclos 202612 - 202618)")
        st.caption("Analiza el comportamiento de tu base de consultoras (Activas, Inactivas 1 a 6) y simula el impacto en Actividad y Saldo al reactivarlas.")
        
        grp_ret_target = str(df_filtrado.iloc[0].get('Codigo de grupo', user_grupo)).strip().split('.')[0] if not df_filtrado.empty else user_grupo
        diag_ret = obtener_diagnostico_retencion_grupo(grupo=grp_ret_target, sector=user_sector)
        conteos_ret = diag_ret['conteos']
        
        c_act_ini = conteos_ret['Activa'] if conteos_ret['Activa'] > 0 else int(val_a_real)
        c_i1_ini = conteos_ret['Inactiva 1'] if conteos_ret['Inactiva 1'] > 0 else max(10, int(val_a_real * 0.35))
        c_i2_ini = conteos_ret['Inactiva 2'] if conteos_ret['Inactiva 2'] > 0 else max(5, int(val_a_real * 0.15))
        c_i3_ini = conteos_ret['Inactiva 3'] if conteos_ret['Inactiva 3'] > 0 else max(3, int(val_a_real * 0.08))
        c_i4_ini = conteos_ret['Inactiva 4'] if conteos_ret['Inactiva 4'] > 0 else max(2, int(val_a_real * 0.05))
        c_i5_ini = conteos_ret['Inactiva 5'] if conteos_ret['Inactiva 5'] > 0 else max(1, int(val_a_real * 0.03))
        c_i6_ini = conteos_ret['Inactiva 6'] if conteos_ret['Inactiva 6'] > 0 else max(1, int(val_a_real * 0.02))
        
        col_ret_diag, col_ret_sim = st.columns([1.1, 1.9])
        
        with col_ret_diag:
            st.markdown("###### 📊 Situación Actual de la Red")
            df_base_sit = pd.DataFrame({
                "Estado Comercial": ["Activas", "Inactivas 1", "Inactivas 2", "Inactivas 3", "Inactivas 4", "Inactivas 5", "Inactivas 6"],
                "Consultoras": [c_act_ini, c_i1_ini, c_i2_ini, c_i3_ini, c_i4_ini, c_i5_ini, c_i6_ini]
            })
            st.dataframe(df_base_sit, use_container_width=True, hide_index=True)
            
            tot_disp_ini = c_act_ini + c_i1_ini + c_i2_ini + c_i3_ini
            pct_act_ini = (c_act_ini / tot_disp_ini * 100.0) if tot_disp_ini > 0 else 0.0
            
            st.metric("Disponibles Base (Act + I1..I3)", f"{tot_disp_ini}")
            st.metric("% Actividad Base Actual", f"{pct_act_ini:.1f}%")

        with col_ret_sim:
            st.markdown("###### 🎯 Simular Ciclo a Proyectar")
            col_ciclo_sel, col_inicios_rei = st.columns(2)
            with col_ciclo_sel:
                ciclo_proy_sel = st.selectbox("¿Qué ciclo vas a proyectar?", options=["202613", "202614", "202615", "202616", "202617", "202618"], key=f"sel_ciclo_ret_{leader_key}")
            with col_inicios_rei:
                sim_ret_inicios = st.number_input("Inicios Nuevos Proyectados:", value=int(val_inicios), min_value=0, step=1, key=f"ret_ini_{leader_key}")
                sim_ret_reinicios = st.number_input("Reinicios Proyectados:", value=2, min_value=0, step=1, key=f"ret_rei_{leader_key}")
            
            st.markdown("###### Proyección de Activación por Tramo:")
            col_sl1, col_sl2 = st.columns(2)
            with col_sl1:
                p_act_retenidas = st.slider("Activas que repetirán pedido:", min_value=0, max_value=max(1, c_act_ini), value=min(c_act_ini, int(c_act_ini * 0.70)), key=f"sl_ret_act_{leader_key}")
                p_i1_recup = st.slider("Inactivas 1 a reactivar:", min_value=0, max_value=max(1, c_i1_ini), value=min(c_i1_ini, int(c_i1_ini * 0.50)), key=f"sl_ret_i1_{leader_key}")
            with col_sl2:
                p_i2_recup = st.slider("Inactivas 2 a reactivar:", min_value=0, max_value=max(1, c_i2_ini), value=min(c_i2_ini, int(c_i2_ini * 0.30)), key=f"sl_ret_i2_{leader_key}")
                p_i3_recup = st.slider("Inactivas 3 a reactivar:", min_value=0, max_value=max(1, c_i3_ini), value=min(c_i3_ini, int(c_i3_ini * 0.20)), key=f"sl_ret_i3_{leader_key}")

            tot_act_proyectadas = p_act_retenidas + p_i1_recup + p_i2_recup + p_i3_recup + sim_ret_inicios + sim_ret_reinicios
            sin_activar_total = (c_act_ini - p_act_retenidas) + (c_i1_ini - p_i1_recup) + (c_i2_ini - p_i2_recup) + (c_i3_ini - p_i3_recup)
            tot_disp_proy = tot_act_proyectadas + sin_activar_total
            pct_act_proy = (tot_act_proyectadas / tot_disp_proy * 100.0) if tot_disp_proy > 0 else 0.0
            
            saldo_proy_estimado = sim_ret_inicios + sim_ret_reinicios - (c_i3_ini - p_i3_recup)

            st.markdown("---")
            kr1, kr2, kr3 = st.columns(3)
            with kr1:
                st.metric("Total Activas Proyectadas", f"{tot_act_proyectadas}", f"+{sim_ret_inicios} Inicios / +{sim_ret_reinicios} Reinicios")
            with kr2:
                st.metric("% Actividad Proyectada", f"{pct_act_proy:.1f}%", f"{pct_act_proy - pct_act_ini:+.1f}% vs Actual")
            with kr3:
                st.metric("Saldo Proyectado Estimado", f"{saldo_proy_estimado:+d}", "Crecimiento de Red" if saldo_proy_estimado >= 0 else "Riesgo de decrecimiento", delta_color="normal" if saldo_proy_estimado >= 2 else "inverse")

    # 3. SIMULADOR DE CONVENCIÓN NATURA
    with tab_sub_convencion:
        st.markdown("##### 🏆 Simulador de Puntos a Convención Natura (7 Ciclos: 202612 - 202618)")
        st.caption("Calcula y proyecta tu puntaje oficial acumulado para clasificar a la Convención Nacional/Internacional Natura según el saldo de cada ciclo.")
        
        ciclos_conv = ["202612", "202613", "202614", "202615", "202616", "202617", "202618"]
        
        col_pres1, col_pres2, col_pres3, col_pres4 = st.columns(4)
        if f'saldos_conv_{leader_key}' not in st.session_state:
            st.session_state[f'saldos_conv_{leader_key}'] = [int(val_saldo)] + [3]*6

        with col_pres1:
            if st.button("🌱 Preset Conservador (Saldo 2)", use_container_width=True, key=f"p_cons_{leader_key}"):
                st.session_state[f'saldos_conv_{leader_key}'] = [int(val_saldo)] + [2]*6
                st.rerun()
        with col_pres2:
            if st.button("⭐ Preset Destacado (Saldo 4)", use_container_width=True, key=f"p_med_{leader_key}"):
                st.session_state[f'saldos_conv_{leader_key}'] = [int(val_saldo)] + [4]*6
                st.rerun()
        with col_pres3:
            if st.button("💎 Preset Diamante (Saldo 6+)", use_container_width=True, key=f"p_dia_{leader_key}"):
                st.session_state[f'saldos_conv_{leader_key}'] = [int(val_saldo)] + [6]*6
                st.rerun()
        with col_pres4:
            if st.button("🔄 Restablecer a Actual", use_container_width=True, key=f"p_rst_{leader_key}"):
                st.session_state[f'saldos_conv_{leader_key}'] = [int(val_saldo)] * 7
                st.rerun()

        cols_ciclos = st.columns(7)
        saldos_ingresados = []
        puntos_ciclos = []
        
        for i, c_name in enumerate(ciclos_conv):
            with cols_ciclos[i]:
                st.markdown(f"**Ciclo {c_name}**" + (" 📍 *(Actual)*" if i == 0 else ""))
                s_def = st.session_state[f'saldos_conv_{leader_key}'][i] if i < len(st.session_state[f'saldos_conv_{leader_key}']) else 3
                s_val = st.number_input(f"Saldo {c_name}:", value=int(s_def), min_value=-20, max_value=50, step=1, key=f"conv_s_{c_name}_{leader_key}")
                saldos_ingresados.append(s_val)
                pts, r_txt = calcular_puntos_convencion_ciclo(s_val)
                puntos_ciclos.append(pts)
                
                if pts >= 100:
                    st.success(f"💎 **{pts} Pts**\n\n*(Saldo: {r_txt})*")
                elif pts >= 60:
                    st.info(f"⭐ **{pts} Pts**\n\n*(Saldo: {r_txt})*")
                elif pts > 0:
                    st.warning(f"🟡 **{pts} Pts**\n\n*(Saldo: {r_txt})*")
                else:
                    st.error(f"❌ **0 Pts**\n\n*(Saldo: <2)*")

        st.session_state[f'saldos_conv_{leader_key}'] = saldos_ingresados
        puntos_totales_conv = sum(puntos_ciclos)
        
        st.markdown("---")
        kc1, kc2, kc3 = st.columns([1.5, 2, 1.5])
        with kc1:
            st.metric(
                "Puntaje Total Proyectado",
                f"{puntos_totales_conv} Pts",
                f"Promedio: {puntos_totales_conv/7:.1f} pts/ciclo"
            )
        with kc2:
            pct_meta_conv = min(1.0, puntos_totales_conv / 560.0)
            st.markdown(f"**Termómetro de Clasificación a Convención:** ({puntos_totales_conv} / 560 Pts)")
            st.progress(pct_meta_conv)
            if puntos_totales_conv >= 700:
                st.balloons()
                st.success("🌟 **¡Nivel Excepcional!** Estás en posición privilegiada para ganar la Convención Nacional e Internacional.")
            elif puntos_totales_conv >= 420:
                st.info("✈️ **¡En Carrera de Clasificación!** Mantén este ritmo de saldo para asegurar tu cupo a la Convención.")
            else:
                st.warning("⚠️ **Atención:** Necesitas promediar al menos saldo 3 (60 pts) o saldo 4 (80 pts) en los ciclos restantes para clasificar.")
        with kc3:
            st.caption("📋 **Tabla Oficial de Puntos por Saldo:**\n* Saldo < 2: **0 Pts**\n* Saldo = 2: **40 Pts**\n* Saldo = 3: **60 Pts**\n* Saldo = 4: **80 Pts**\n* Saldo = 5: **100 Pts**\n* Saldo ≥ 6: **120 Pts**")

# --- TAB 3: MIS LÍDERES ---
with tab_diagnostico:
    st.subheader("👑 Mis Líderes")
    st.markdown("Generación de tablas dinámicas automatizadas para medición y seguimiento comparativo entre todas las Líderes de Negocio.")
    
    # Utilizar el conjunto de datos completo (df) para permitir la medición comparativa entre Líderes
    df_diag = df.copy()
    if gerencia_seleccionada != "Todas" and col_gerencia and col_gerencia in df_diag.columns:
        df_diag = df_diag[df_diag[col_gerencia] == gerencia_seleccionada]
        
    col_lider = 'Nombre de consultora' if 'Nombre de consultora' in df_diag.columns else (df_diag.columns[0] if len(df_diag.columns) > 0 else '')
    diag = generar_analisis_como_vamos(df_diag) if not df_diag.empty else None

    if df_diag.empty:
        st.info("ℹ️ No hay datos de 'Cómo Vamos' cargados para mostrar las tablas dinámicas de diagnóstico. Sube un archivo desde 'Rotación de Ciclo' para comenzar.")
    else:
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

        col_f1, col_f2 = st.columns([3.2, 1.8])
        with col_f1:
            filtro_segmento = st.radio(
                "🎯 **Filtrar por Tipo de Red:**",
                options=[
                    f"🌟 Todas ({count_tot})",
                    f"👑 Líderes ({count_lideres})",
                    f"🌱 Semillas ({count_semillas})"
                ],
                horizontal=True,
                key="filtro_segmento_red_diagnostico"
            )
        with col_f2:
            st.caption("💡 **Criterio de Clasificación:**\n* **👑 Líderes:** Desafío/Meta Facturación > $0\n* **🌱 Semillas:** Desafío/Meta Facturación = $0 o menor")

        if "👑 Líderes" in filtro_segmento:
            df_diag = df_diag[df_diag['Tipo_Red'] == '👑 Líder'].copy()
        elif "🌱 Semillas" in filtro_segmento:
            df_diag = df_diag[df_diag['Tipo_Red'] == '🌱 Semilla'].copy()

        # --- 1. TABLA DE FACTURACIÓN (Formato exacto Clery Cuellar + Ganancia Estimada Total) ---
        st.markdown("#### 💰 1. Tabla de Facturación y Cumplimiento (Ordenadas de Mayor a Menor Cumplimiento)")
        
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
            
        # Aplicar paleta de colores condicionales a Tabla 1
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

        st.markdown("---")

        # --- 2. TABLA DE ACTIVAS / PEDIDOS ---
        st.markdown("#### 👥 2. Tabla de Activas / Pedidos (Ordenadas de Mayor a Menor Cumplimiento)")
        
        # Asegurar cálculo dinámico de Cumplimiento Activas
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

        # Aplicar paleta de colores condicionales a Tabla 2
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

        # --- 3. CUADRO RESUMEN DE DISPONIBLES (Disponibles Proyectadas, Día XX, % Cump LN, falta) ---
        st.markdown("---")
        col_hdr_disp, col_num_dia = st.columns([3.2, 1.8])
        with col_hdr_disp:
            st.markdown("#### 📋 3. Cuadro Resumen de Disponibles (Desafío vs. Avance por Día)")
            st.caption("Fórmulas del modelo: Disponibles Proyectadas extraídas de `Objetivos Arte.xlsx` (Desafíos LNN), `% Cump LN = (Día / Proyectadas) * 100` y `falta = Proyectadas - Día`.")
        with col_num_dia:
            dia_corte = st.number_input("📅 Día de Avance (Editable):", min_value=1, max_value=21, value=14, step=1, key="dia_avance_corte_14_key")

        nombre_col_dia = f"Dia {dia_corte}"

        if col_lider and col_lider in df_diag.columns:
            df_disp_prep = df_diag.copy()
            col_grp_diag = next((c for c in df_disp_prep.columns if any(k in str(c).lower() for k in ['código de grupo', 'codigo de grupo', 'cód. grupo', 'cod grupo', 'grupo'])), None)
            
            mapa_arte_disp = cargar_objetivos_arte()
            mapa_grp_disp = mapa_arte_disp.get('por_grupo', {})
            mapa_nom_disp = mapa_arte_disp.get('por_nombre', {})

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
                    # Fallback a disponibles actuales si no existe meta cargada
                    return int(limpiar_numero(row.get(col_disp_actual, 0), 0))

                df_disp_calc['Disponibles Proyectadas'] = df_disp_prep.apply(_obtener_desafio_disp_row, axis=1)
                df_disp_calc[nombre_col_dia] = df_disp_prep[col_disp_actual].apply(lambda v: int(limpiar_numero(v, 0)))
                
                # Fórmulas del Cuadro
                df_disp_calc['% Cump LN'] = df_disp_calc.apply(
                    lambda r: (r[nombre_col_dia] / r['Disponibles Proyectadas'] * 100.0) if r['Disponibles Proyectadas'] > 0 else 0.0,
                    axis=1
                )
                df_disp_calc['falta'] = df_disp_calc.apply(
                    lambda r: max(0, r['Disponibles Proyectadas'] - r[nombre_col_dia]),
                    axis=1
                )

                # Ordenar por Cumplimiento Descendente
                df_disp_calc = df_disp_calc.sort_values(by='% Cump LN', ascending=False).reset_index(drop=True)

                # Fila de Totales
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

                # Formatear valores para visualización limpiando ceros e incluyendo %
                df_disp_formatted = df_disp_final.copy()
                df_disp_formatted['Disponibles Proyectadas'] = df_disp_formatted['Disponibles Proyectadas'].apply(lambda v: f"{int(v):,}".replace(",", "."))
                df_disp_formatted[nombre_col_dia] = df_disp_formatted[nombre_col_dia].apply(lambda v: f"{int(v):,}".replace(",", "."))
                df_disp_formatted['% Cump LN'] = df_disp_formatted['% Cump LN'].apply(lambda v: f"{v:.1f}%")
                df_disp_formatted['falta'] = df_disp_formatted['falta'].apply(lambda v: f"{int(v):,}".replace(",", "."))

                # Funciones de estilo condicional de semáforo armónico
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

                # Aplicar Styler con compatibilidad
                styler_disp = df_disp_formatted.style
                if hasattr(styler_disp, 'map'):
                    styler_disp = styler_disp.map(_estilo_cump_val, subset=['% Cump LN']).map(_estilo_falta_val, subset=['falta'])
                elif hasattr(styler_disp, 'applymap'):
                    styler_disp = styler_disp.applymap(_estilo_cump_val, subset=['% Cump LN']).applymap(_estilo_falta_val, subset=['falta'])

                # Validar si el día possède información de avance
                if tot_dia > 0:
                    st.dataframe(styler_disp, use_container_width=True)
                elif tot_desafios > 0:
                    st.markdown(f"""
                    <div style="background: linear-gradient(135deg, rgba(255, 107, 0, 0.1), rgba(227, 0, 123, 0.1)); border: 1px solid rgba(255, 107, 0, 0.3); border-radius: 14px; padding: 20px; text-align: center; margin: 12px 0;">
                        <div style="font-size: 2rem; margin-bottom: 6px;">📅</div>
                        <h5 style="margin: 0 0 6px 0; color: #FF8833; font-weight: 800; font-size: 1.1rem;">Sin Información Registrada para el {nombre_col_dia}</h5>
                        <p style="margin: 0; font-size: 0.93rem; opacity: 0.88; line-height: 1.4;">
                            Actualmente no se registran pedidos o activas acumuladas para el <b>Día {dia_corte}</b> en el corte seleccionado.<br>
                            <span style="font-size: 0.85rem; opacity: 0.75;">(Prueba ajustando el número del día o subiendo un archivo de corte actualizado).</span>
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.info(f"ℹ️ No hay datos cargados para generar el resumen de disponibles para el {nombre_col_dia}.")

        # --- 4. CUADRO RESUMEN DE INICIOS + REINICIOS (Meta, Hoy, Avance, para activar!) ---
        st.markdown("---")
        st.markdown("#### 🚀 4. Cuadro Resumen de Inicios + Reinicios (Meta vs. Ingresos del Ciclo)")
        st.caption("Fórmulas del modelo: `Hoy = Inicios + Reinicios`, `Avance = (Hoy / Meta) * 100`, `para activar! = Meta - Hoy`.")

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

            # Ordenar por Porcentaje de Avance Descendente (de mayor a menor)
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

        # --- 5. CUADRO RESUMEN DE RECUPEROS (Meta vs. Recuperos del Ciclo) ---
        st.markdown("---")
        st.markdown("#### 🎯 5. Cuadro Resumen de Recuperos (Meta vs. Recuperos del Ciclo)")
        st.caption("Fórmulas del modelo: `Hoy = Recuperos Reales`, `Avance = (Hoy / Meta) * 100`, `para activar! = Meta - Hoy`.")

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

        # --- 6. CUADRO RESUMEN DE RETENCIÓN I2 (Meta 8% Fuga / Retención I2) ---
        st.markdown("---")
        st.markdown("#### 🔄 6. Cuadro Resumen de Retención I2 (Meta 8% Máx. Fuga I2)")
        st.caption("Fórmulas del modelo: `Meta Retención I2 = Disponibles * 8%`, `Falta I2 Activarse = Inactiva 2 - Meta I2`, `% Retención I2 = (Inactiva 2 / Disponibles) * 100`, `Avance = Inactiva 2 Anterior - Inactiva 2 Actual`.")

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

        # --- 7. CUADRO RESUMEN DE RETENCIÓN I3 (Meta 6% Fuga / Retención I3) ---
        st.markdown("---")
        st.markdown("#### 🔄 7. Cuadro Resumen de Retención I3 (Meta 6% Máx. Fuga I3)")
        st.caption("Fórmulas del modelo: `Meta Retención I3 = Disponibles * 6%`, `Falta I3 Activarse = Inactiva 3 - Meta I3`, `% Retención I3 = (Inactiva 3 / Disponibles) * 100`, `Avance = Inactiva 3 Anterior - Inactiva 3 Actual`.")

        if col_lider and col_lider in df_diag.columns:
            df_i3_prep = df_diag.copy()
            col_disp_i3 = 'Disponibles' if 'Disponibles' in df_i3_prep.columns else None
            col_i3 = next((c for c in df_i3_prep.columns if str(c).lower().strip() in ['inactiva 3', 'inactiva_3', 'inactivas 3', 'inactivas_3', 'i3']), None)
            col_i3_ant = next((c for c in df_i3_prep.columns if 'inactiva 3_anterior' in str(c).lower() or 'inactivas 3_anterior' in str(c).lower()), None)

            if col_disp_i3 and col_i3:
                df_i3_calc = pd.DataFrame()
                df_i3_calc['LÍDER DE NEGOCIOS'] = df_i3_prep[col_lider].astype(str)

                val_disp3 = df_i3_prep[col_disp_i3].apply(lambda v: limpiar_numero(v, 0.0))
                val_i3 = df_i3_prep[col_i3].apply(lambda v: limpiar_numero(v, 0.0))

                df_i3_calc['META RETENCIÓN I3'] = (val_disp3 * 0.06).round().astype(int)
                df_i3_calc['FALTA I3 ACTIVARSE'] = (val_i3 - df_i3_calc['META RETENCIÓN I3']).round().astype(int)
                df_i3_calc['% RETENCIÓN META 6%'] = (val_i3 / val_disp3.replace(0, pd.NA) * 100.0).fillna(0.0)

                if col_i3_ant:
                    val_i3_ant = df_i3_prep[col_i3_ant].apply(lambda v: limpiar_numero(v, 0.0))
                    df_i3_calc['AVANCE RETENCION I3'] = (val_i3_ant - val_i3).fillna(0).astype(int)
                else:
                    df_i3_calc['AVANCE RETENCION I3'] = 0

                df_i3_calc = df_i3_calc.sort_values(by='% RETENCIÓN META 6%', ascending=True).reset_index(drop=True)

                tot_disp_i3 = float(val_disp3.sum())
                tot_i3 = float(val_i3.sum())
                tot_meta_i3 = int(df_i3_calc['META RETENCIÓN I3'].sum())
                tot_falta_i3 = int(df_i3_calc['FALTA I3 ACTIVARSE'].sum())
                tot_pct_i3 = (tot_i3 / tot_disp_i3 * 100.0) if tot_disp_i3 > 0 else 0.0
                tot_av_i3 = int(df_i3_calc['AVANCE RETENCION I3'].sum())

                row_tot_i3 = pd.DataFrame([{
                    'LÍDER DE NEGOCIOS': 'TOTAL GENERAL',
                    'META RETENCIÓN I3': tot_meta_i3,
                    'FALTA I3 ACTIVARSE': tot_falta_i3,
                    '% RETENCIÓN META 6%': tot_pct_i3,
                    'AVANCE RETENCION I3': tot_av_i3
                }])
                df_i3_final = pd.concat([df_i3_calc, row_tot_i3], ignore_index=True)

                df_i3_formatted = df_i3_final[['LÍDER DE NEGOCIOS', 'META RETENCIÓN I3', 'FALTA I3 ACTIVARSE', '% RETENCIÓN META 6%', 'AVANCE RETENCION I3']].copy()
                df_i3_formatted['META RETENCIÓN I3'] = df_i3_formatted['META RETENCIÓN I3'].apply(lambda v: f"{int(v):,}".replace(",", "."))
                df_i3_formatted['FALTA I3 ACTIVARSE'] = df_i3_formatted['FALTA I3 ACTIVARSE'].apply(lambda v: f"{int(v):,}".replace(",", "."))
                df_i3_formatted['% RETENCIÓN META 6%'] = df_i3_formatted['% RETENCIÓN META 6%'].apply(lambda v: f"{v:.1f}%")
                df_i3_formatted['AVANCE RETENCION I3'] = df_i3_formatted['AVANCE RETENCION I3'].apply(lambda v: f"{int(v):,}".replace(",", "."))

                def _estilo_pct_retencion_6(val_str):
                    try:
                        num = float(str(val_str).replace('%', '').strip())
                        if num <= 6.0:
                            return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                        elif num <= 8.0:
                            return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                        else:
                            return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
                    except Exception:
                        return ''

                styler_i3 = df_i3_formatted.style
                if hasattr(styler_i3, 'map'):
                    styler_i3 = styler_i3.map(_estilo_falta_retencion, subset=['FALTA I3 ACTIVARSE']).map(_estilo_pct_retencion_6, subset=['% RETENCIÓN META 6%'])
                elif hasattr(styler_i3, 'applymap'):
                    styler_i3 = styler_i3.applymap(_estilo_falta_retencion, subset=['FALTA I3 ACTIVARSE']).applymap(_estilo_pct_retencion_6, subset=['% RETENCIÓN META 6%'])

                st.dataframe(styler_i3, use_container_width=True)

        # --- 8. CUADRO DE ACTIVAS Y ACTIVIDAD FRECUENTE ---
        st.markdown("---")
        st.markdown("#### 💎 8. Cuadro de Activas y Actividad Frecuente (Base Estable)")
        st.caption("Fórmulas del modelo: `Activas Frecuentes = Real Activas - Recuperos - Inicios - Reinicios`, `Actividad Frecuente = (Activas Frecuentes / Disponibles) * 100`.")

        if col_lider and col_lider in df_diag.columns:
            df_af_prep = df_diag.copy()
            col_disp_af = 'Disponibles' if 'Disponibles' in df_af_prep.columns else None
            col_real_act = 'Real Activas' if 'Real Activas' in df_af_prep.columns else None
            col_rec_af = 'Recuperos' if 'Recuperos' in df_af_prep.columns else None
            col_ini_af = 'Inicios' if 'Inicios' in df_af_prep.columns else None
            col_rei_af = 'Reinicios' if 'Reinicios' in df_af_prep.columns else None
            col_af_directa = next((c for c in df_af_prep.columns if 'activas frecuentes' in str(c).lower() or 'activas_frecuentes' in str(c).lower()), None)

            if col_disp_af:
                df_af_calc = pd.DataFrame()
                df_af_calc['LÍDER DE NEGOCIOS'] = df_af_prep[col_lider].astype(str)

                val_disp_af = df_af_prep[col_disp_af].apply(lambda v: limpiar_numero(v, 0.0))

                if col_real_act:
                    r_act = df_af_prep[col_real_act].apply(lambda v: limpiar_numero(v, 0.0))
                    r_rec = df_af_prep[col_rec_af].apply(lambda v: limpiar_numero(v, 0.0)) if col_rec_af else 0
                    r_ini = df_af_prep[col_ini_af].apply(lambda v: limpiar_numero(v, 0.0)) if col_ini_af else 0
                    r_rei = df_af_prep[col_rei_af].apply(lambda v: limpiar_numero(v, 0.0)) if col_rei_af else 0
                    val_act_frec = (r_act - r_rec - r_ini - r_rei).apply(lambda v: max(0, v))
                elif col_af_directa:
                    val_act_frec = df_af_prep[col_af_directa].apply(lambda v: limpiar_numero(v, 0.0))
                else:
                    val_act_frec = pd.Series(0, index=df_af_prep.index)

                df_af_calc['ACTIVAS FRECUENTES'] = val_act_frec.round().astype(int)
                # Cálculo de Actividad Frecuente (%) según fórmula oficial: (Activas Frecuentes / Disponibles) * 100
                df_af_calc['ACTIVIDAD FRECUENTE'] = (val_act_frec / val_disp_af.replace(0, pd.NA) * 100.0).fillna(0.0)

                df_af_calc = df_af_calc.sort_values(by='ACTIVIDAD FRECUENTE', ascending=False).reset_index(drop=True)

                tot_disp_af = float(val_disp_af.sum())
                tot_af = int(df_af_calc['ACTIVAS FRECUENTES'].sum())
                tot_pct_af = (tot_af / tot_disp_af * 100.0) if tot_disp_af > 0 else 0.0

                row_tot_af = pd.DataFrame([{
                    'LÍDER DE NEGOCIOS': 'TOTAL GENERAL',
                    'ACTIVAS FRECUENTES': tot_af,
                    'ACTIVIDAD FRECUENTE': tot_pct_af
                }])
                df_af_final = pd.concat([df_af_calc, row_tot_af], ignore_index=True)

                df_af_formatted = df_af_final[['LÍDER DE NEGOCIOS', 'ACTIVAS FRECUENTES', 'ACTIVIDAD FRECUENTE']].copy()
                df_af_formatted['ACTIVAS FRECUENTES'] = df_af_formatted['ACTIVAS FRECUENTES'].apply(lambda v: f"{int(v):,}".replace(",", "."))
                df_af_formatted['ACTIVIDAD FRECUENTE'] = df_af_formatted['ACTIVIDAD FRECUENTE'].apply(lambda v: f"{v:.1f}%")

                def _estilo_actividad_frecuente(val_str):
                    try:
                        num = float(str(val_str).replace('%', '').strip())
                        if num >= 55.0:
                            return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
                        elif num >= 50.0:
                            return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
                        else:
                            return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
                    except Exception:
                        return ''

                styler_af = df_af_formatted.style
                if hasattr(styler_af, 'map'):
                    styler_af = styler_af.map(_estilo_actividad_frecuente, subset=['ACTIVIDAD FRECUENTE'])
                elif hasattr(styler_af, 'applymap'):
                    styler_af = styler_af.applymap(_estilo_actividad_frecuente, subset=['ACTIVIDAD FRECUENTE'])

                st.dataframe(styler_af, use_container_width=True)

    st.markdown("---")

    # --- 9. MÓDULO DE COMPARTIR POR WHATSAPP ---
    st.markdown("#### 📲 9. Módulo para Compartir Resumen por WhatsApp")
    st.caption("Selecciona una Líder para generar su reporte en formato texto listo para copiar o enviar directamente por WhatsApp Web / Móvil.")
    
    lider_sel = None
    if col_lider and col_lider in df_filtrado.columns and not df_filtrado.empty:
        lista_lideres = sorted(df_filtrado[col_lider].dropna().astype(str).unique())
        lider_sel = st.sidebar.selectbox("👤 Selecciona la Líder para enviar reporte:", options=lista_lideres) if False else st.selectbox("👤 Selecciona la Líder para enviar reporte:", options=lista_lideres)
    
    if lider_sel and col_lider and col_lider in df_filtrado.columns and not df_filtrado.empty:
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
    
    # Ordenar de mayor a menor por Real Activas para mantener consistencia
    if 'Real Activas' in df_filtrado.columns:
        df_metas_sorted = df_filtrado.sort_values(by='Real Activas', ascending=False)
    else:
        df_metas_sorted = df_filtrado
        
    df_metas_view = df_metas_sorted[cols_existentes_metas].copy()

    # Formatear números enteros (Real Activas y Tramos de Meta) sin decimales (.000000)
    cols_enteras = ['Real Activas', 'Meta_Crecer_1plus_150k', 'Meta_Crecer_3plus_200k', 'Meta_Crecer_5plus_300k', 'Meta_Crecer_7plus_500k', 'Meta_Crecer_9plus_750k']
    for col_e in cols_enteras:
        if col_e in df_metas_view.columns:
            df_metas_view[col_e] = df_metas_view[col_e].apply(lambda v: f"{int(limpiar_numero(v, 0))}")

    # Formatear avance % limpiamente
    def _formato_avance_clean(v):
        if pd.isna(v):
            return "N/A"
        try:
            num = float(v)
            if abs(num) > 1.5:
                return f"{num:+.2f}%"
            else:
                return f"{num * 100.0:+.2f}%"
        except Exception:
            return str(v)

    if 'Avance % Facturación' in df_metas_view.columns:
        df_metas_view['Avance % Facturación'] = df_metas_view['Avance % Facturación'].apply(_formato_avance_clean)

    if 'Falta para el 100%' in df_metas_view.columns:
        df_metas_view['Falta para el 100%'] = df_metas_view['Falta para el 100%'].apply(formato_cop_signo)

    df_metas_renamed = df_metas_view.rename(columns={
        'Meta_Crecer_1plus_150k': 'Meta 1+ (+150k)',
        'Meta_Crecer_3plus_200k': 'Meta 3+ (+200k)',
        'Meta_Crecer_5plus_300k': 'Meta 5+ (+300k)',
        'Meta_Crecer_7plus_500k': 'Meta 7+ (+500k)',
        'Meta_Crecer_9plus_750k': 'Meta 9+ (+750k)',
        'Avance % Facturación': 'Avance % vs Ant.'
    })

    # Funciones de estilo condicional para Metas de Crecimiento
    def _estilo_avance_vs_ant(val_str):
        try:
            val_clean = str(val_str).replace('%', '').replace('+', '').strip()
            num = float(val_clean)
            if num > 0:
                return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
            elif num == 0:
                return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
            else:
                return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
        except Exception:
            return ''

    def _estilo_falta_100_metas(val_str):
        try:
            s = str(val_str)
            if '-' in s or '$0' in s:
                return 'background-color: #d1fae5; color: #065f46; font-weight: bold;'
            else:
                return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
        except Exception:
            return ''

    def _estilo_tramos_meta(val):
        return 'background-color: #e0f2fe; color: #0369a1; font-weight: bold;'

    styler_metas = df_metas_renamed.style
    if 'Avance % vs Ant.' in df_metas_renamed.columns:
        if hasattr(styler_metas, 'map'):
            styler_metas = styler_metas.map(_estilo_avance_vs_ant, subset=['Avance % vs Ant.'])
        elif hasattr(styler_metas, 'applymap'):
            styler_metas = styler_metas.applymap(_estilo_avance_vs_ant, subset=['Avance % vs Ant.'])

    if 'Falta para el 100%' in df_metas_renamed.columns:
        if hasattr(styler_metas, 'map'):
            styler_metas = styler_metas.map(_estilo_falta_100_metas, subset=['Falta para el 100%'])
        elif hasattr(styler_metas, 'applymap'):
            styler_metas = styler_metas.applymap(_estilo_falta_100_metas, subset=['Falta para el 100%'])

    tramos_presentes = [c for c in ['Meta 1+ (+150k)', 'Meta 3+ (+200k)', 'Meta 5+ (+300k)', 'Meta 7+ (+500k)', 'Meta 9+ (+750k)'] if c in df_metas_renamed.columns]
    if tramos_presentes:
        if hasattr(styler_metas, 'map'):
            styler_metas = styler_metas.map(_estilo_tramos_meta, subset=tramos_presentes)
        elif hasattr(styler_metas, 'applymap'):
            styler_metas = styler_metas.applymap(_estilo_tramos_meta, subset=tramos_presentes)

    st.dataframe(styler_metas, use_container_width=True)

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
    st.subheader("📥 Centro de Descargas & Exportación")
    st.markdown("Genera y descarga reportes en **Excel (.xlsx)** con colores reales de semáforo o **CSV (.csv)**.")

    st.markdown("#### 📊 1. Base Maestra de Tableau (Consultoras, Puntos & Cartera)")
    df_tab_exp = obtener_base_tableau_completa_original(
        grupo=(user_grupo if user_rol == 'lider' else None),
        sector=(user_sector if (user_rol == 'gerente' and user_sector) else None)
    )
    if df_tab_exp is not None and not df_tab_exp.empty:
        c_te1, c_te2 = st.columns(2)
        with c_te1:
            xl_t_bytes = cached_export_excel_tableau(df_tab_exp)
            st.download_button(
                label="📗 Descargar Excel a Colores (.xlsx)",
                data=xl_t_bytes,
                file_name="Base_Consultoras_Tableau_Colores.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with c_te2:
            csv_t_bytes = cached_export_csv(df_tab_exp)
            st.download_button(
                label="📊 Descargar Archivo CSV (.csv)",
                data=csv_t_bytes,
                file_name="Base_Consultoras_Tableau.csv",
                mime="text/csv",
                use_container_width=True
            )
    else:
        st.info("ℹ️ Sube la base de Tableau en la primera pestaña para habilitar las descargas de este módulo.")

    st.markdown("---")
    st.markdown("#### 🎯 2. Metas de Ciclo 'Cómo Vamos' (Facturación, Activas & Saldos)")
    if not df_filtrado.empty:
        col_exp1, col_exp2 = st.columns(2)
        
        with col_exp1:
            excel_colores_bytes = exportar_excel_con_colores({
                'Activas': df_filtrado[[c for c in ['Nombre de consultora', 'Nombre Setor', 'Color', 'Real Activas', 'Objetivo Activas', 'Cumplimiento Activas'] if c in df_filtrado.columns]],
                'Facturacion': df_filtrado[[c for c in ['Nombre de consultora', 'Nombre Setor', 'Real Facturación', 'Objetivo Facturación', 'Cumplimiento Facturación', 'Falta para el 100%'] if c in df_filtrado.columns]],
                'Saldos': df_filtrado[[c for c in ['Nombre de consultora', 'Nombre Setor', 'Saldo', 'Potencializador_Pct', 'Ganancia estimada'] if c in df_filtrado.columns]],
                'Disponibles': df_filtrado[[c for c in ['Nombre de consultora', 'Nombre Setor', 'Disponibles', 'Real Activas', 'Inicios', 'Reinicios', 'Recuperos'] if c in df_filtrado.columns]],
                'Retención': df_filtrado[[c for c in ['Nombre de consultora', 'Nombre Setor', 'Disponibles', 'Inactiva 2', 'Inactiva 3'] if c in df_filtrado.columns]],
                'Actividad Frecuente': df_filtrado[[c for c in ['Nombre de consultora', 'Nombre Setor', 'Disponibles', 'Real Activas', 'Recuperos', 'Inicios', 'Reinicios', 'Activas Frecuentes', '%Actividad Frecuente'] if c in df_filtrado.columns]]
            })
            st.download_button(
                label="📗 Metas a Colores (.xlsx)",
                data=excel_colores_bytes,
                file_name="Reporte_Metas_Lideres_Colores.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        with col_exp2:
            csv_data = df_filtrado.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📊 Metas en CSV (.csv)",
                data=csv_data,
                file_name="Resultado_Metas_Procesadas_Filtrado.csv",
                mime="text/csv",
                use_container_width=True
            )
    else:
        st.info("ℹ️ Sube el archivo 'Cómo Vamos' en la barra lateral para habilitar las descargas de este módulo.")

# --- TAB 6: GESTIÓN DE USUARIOS & ROLES (EXCLUSIVO SUPER ADMIN) ---
with tab_usuarios:
        st.subheader("🔑 Gestión de Usuarios, Roles & Permisos (Super Admin)")
        st.markdown("Administra credenciales de acceso, asigna roles (*superadmin, gerente, lider, asesor*), activa o desbloquea suscripciones en 1 clic y controla los periodos de prueba de 15 días.")

        # --- SECCIÓN 1: CONTROL DE SUSCRIPCIONES Y DESBLOQUEO EN 1 CLIC ---
        st.markdown("---")
        st.subheader("💳 Control de Suscripciones, Pruebas Gratuitas y Desbloqueos en 1 Clic")
        st.caption("Visualiza el estado de cada Sector registrado y desbloquea o activa planes pagados en tiempo real:")

        df_res_sub = obtener_resumen_suscripciones()
        if not df_res_sub.empty:
            cols_show_sub = [c for c in ["Código Sector", "Nombre Sector", "Gerente Responsable", "Contacto (WhatsApp)", "Líderes Activas", "Estado", "Vence el", "Tiempo Restante"] if c in df_res_sub.columns]
            st.dataframe(df_res_sub[cols_show_sub], use_container_width=True)
            
            col_acc1, col_acc2 = st.columns([1.2, 1])
            with col_acc1:
                st.markdown("##### ⚡ Gestión de Suscripción por Sector")
                lista_sec_opciones = df_res_sub["Código Sector"].tolist()
                
                with st.form("form_gestion_suscripcion_sector"):
                    sel_sec_id = st.selectbox(
                        "Selecciona el Sector a gestionar:",
                        options=lista_sec_opciones,
                        format_func=lambda s: f"Sector {s} — {df_res_sub[df_res_sub['Código Sector'] == s]['Nombre Sector'].iloc[0]} ({df_res_sub[df_res_sub['Código Sector'] == s]['Gerente Responsable'].iloc[0]})"
                    )
                    
                    accion_sub = st.selectbox(
                        "Acción a Realizar:",
                        options=[
                            "🟢 Activar Plan Pagado (+30 Días / 1 Mes)",
                            "🟢 Activar Plan Pagado (+90 Días / Trimestral)",
                            "🟢 Activar Plan Pagado (+365 Días / Anual)",
                            "👑 Activar Suscripción Permanente (Sin Vencimiento)",
                            "⏳ Dar Prórroga de Prueba (+5 Días de Cortesía)",
                            "⛔ Suspender / Bloquear Acceso a este Sector",
                            "🔓 Desbloquear Acceso al Sector"
                        ]
                    )
                    
                    btn_aplicar_sub = st.form_submit_button("🚀 Aplicar Cambio al Sector", type="primary", use_container_width=True)
                    
                    if btn_aplicar_sub:
                        if "30 Días" in accion_sub:
                            ok_s, msg_s = actualizar_suscripcion_sector(sel_sec_id, nuevo_estado="activo", dias_extension=30, es_pago=True)
                        elif "90 Días" in accion_sub:
                            ok_s, msg_s = actualizar_suscripcion_sector(sel_sec_id, nuevo_estado="activo", dias_extension=90, es_pago=True)
                        elif "365 Días" in accion_sub:
                            ok_s, msg_s = actualizar_suscripcion_sector(sel_sec_id, nuevo_estado="activo", dias_extension=365, es_pago=True)
                        elif "Permanente" in accion_sub:
                            ok_s, msg_s = actualizar_suscripcion_sector(sel_sec_id, nuevo_estado="activo", dias_extension=-1, es_pago=True)
                        elif "Prórroga" in accion_sub:
                            ok_s, msg_s = actualizar_suscripcion_sector(sel_sec_id, nuevo_estado="prueba", dias_extension=5, es_pago=False)
                        elif "Suspender" in accion_sub:
                            ok_s, msg_s = actualizar_suscripcion_sector(sel_sec_id, nuevo_estado="bloqueado", dias_extension=0, es_pago=False)
                        elif "Desbloquear" in accion_sub:
                            ok_s, msg_s = actualizar_suscripcion_sector(sel_sec_id, nuevo_estado="activo", dias_extension=30, es_pago=True)
                        
                        if ok_s:
                            st.success(f"✅ ¡Éxito! {msg_s}")
                            st.rerun()
                        else:
                            st.error(f"❌ {msg_s}")
                            
            with col_acc2:
                st.markdown("##### ℹ️ Información de Ayuda")
                st.info(
                    "💡 **¿Cómo funciona el desbloqueo?**\n\n"
                    "- Al seleccionar **Activar Plan Pagado**, la Gerente y **todas sus líderes** recuperan acceso inmediato.\n"
                    "- Toda la data previa, comentarios y notas de consultoras quedan disponibles al instante.\n"
                    "- El sistema audita y actualiza las cuentas en cascada."
                )

        st.markdown("---")
        st.subheader("👥 Gestión de Cuentas, Directorio & Restablecimiento de Claves")
        col_u1, col_u2 = st.columns([1.2, 1])

        with col_u1:
            st.markdown("##### 📋 Listado General de Cuentas Registradas")
            users_dict = cargar_usuarios()
            list_u = []
            for uname, udata in users_dict.items():
                list_u.append({
                    "Usuario": uname,
                    "Nombre": udata.get("nombre", ""),
                    "Rol": udata.get("rol", ""),
                    "Código de Grupo": udata.get("codigo_grupo") or "N/A",
                    "Código de Sector": udata.get("codigo_sector") or "N/A",
                    "Nombre Sector": udata.get("nombre_sector") or "N/A",
                    "Estado": udata.get("estado_suscripcion") or "activo"
                })
            df_u_all = pd.DataFrame(list_u)
            st.dataframe(df_u_all, use_container_width=True)

            csv_all_users = df_u_all.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Descargar Directorio de Cuentas Completo (CSV / Excel)",
                data=csv_all_users,
                file_name="Directorio_General_Usuarios_App.csv",
                mime="text/csv",
                use_container_width=True,
                key="btn_descarga_csv_admin_users"
            )

        with col_u2:
            tab_reseteo_u, tab_crear_u, tab_eliminar_u = st.tabs(["⚡ Reseteo Rápido & WA", "➕ Crear / Editar", "🗑️ Eliminar"])
            
            with tab_reseteo_u:
                st.markdown("##### ⚡ Restablecer Contraseña")
                st.caption("Selecciona cualquier usuario para restaurar su clave a **'lider123'** o la que prefieras:")
                with st.form("form_reseteo_rapido_admin"):
                    u_sel_reset = st.selectbox(
                        "Selecciona el usuario:",
                        options=list(users_dict.keys()),
                        format_func=lambda u: f"👤 {u} — {users_dict[u].get('nombre', '')} ({users_dict[u].get('rol', '')})"
                    )
                    pass_nueva_admin = st.text_input("Nueva Contraseña:", value="lider123")
                    chk_forzar_admin = st.checkbox("Pedir cambio de clave al iniciar sesión", value=False)
                    btn_run_reset_admin = st.form_submit_button("🔄 Restablecer Contraseña Ahora", type="primary", use_container_width=True)
                    
                    if btn_run_reset_admin:
                        ok_r_a, msg_r_a = restablecer_password_usuario(u_sel_reset, pass_nueva_admin, debe_cambiar=chk_forzar_admin)
                        if ok_r_a:
                            st.success(f"✅ ¡Éxito! Contraseña de **{u_sel_reset}** actualizada a: `{pass_nueva_admin}`")
                            st.session_state['ultimo_reseteo_admin'] = {
                                'usuario': u_sel_reset,
                                'nombre': users_dict[u_sel_reset].get('nombre', ''),
                                'password': pass_nueva_admin
                            }
                            st.rerun()
                        else:
                            st.error(f"❌ {msg_r_a}")
                
                # Compartir accesos por WhatsApp
                ult_a = st.session_state.get('ultimo_reseteo_admin')
                if ult_a:
                    st.markdown("---")
                    st.markdown("###### 📲 Compartir Accesos por WhatsApp")
                    tel_a_in = st.text_input("Número celular (10 dígitos):", placeholder="ej. 3123456789", key="tel_a_in_reset")
                    msg_wa_admin = (
                        f"🌸 ¡Hola {ult_a['nombre'].split()[0].title() if ult_a['nombre'] else 'Líder'}! Te comparto tus credenciales de acceso al Sistema de Gestión:\n\n"
                        f"👤 *Usuario:* `{ult_a['usuario']}`\n"
                        f"🔑 *Contraseña:* `{ult_a['password']}`\n\n"
                        f"🌐 *Enlace:* https://app-lideres-production.up.railway.app\n\n"
                        f"¡Muchos éxitos! ✨"
                    )
                    st.text_area("Mensaje listo para enviar:", msg_wa_admin, height=120, key="txt_wa_admin_box")
                    if tel_a_in and len(tel_a_in.strip()) >= 10:
                        link_wa_a = f"https://api.whatsapp.com/send?phone=57{tel_a_in.strip()}&text={urllib.parse.quote(msg_wa_admin)}"
                        st.link_button("📲 Enviar Credenciales por WhatsApp", url=link_wa_a, use_container_width=True)

            with tab_crear_u:
                st.markdown("##### 👤 Crear o Modificar Cuenta")
                with st.form("form_nuevo_usuario"):
                    nu_username = st.text_input("Usuario (Login)", placeholder="ej. dolly.parra@natura.net o lider9334")
                    nu_nombre = st.text_input("Nombre Completo", placeholder="ej. Dolly Parra")
                    nu_pass = st.text_input("Contraseña", type="password", placeholder="Dejar vacío para mantener contraseña actual")
                    nu_rol = st.selectbox("Rol de Acceso", options=["gerente", "lider", "superadmin", "asesor"])
                    nu_grupo = st.text_input("Código de Grupo (Para Líderes)", placeholder="ej. 9334")
                    nu_sector = st.text_input("Código de Sector (Para Gerentes)", placeholder="ej. 700000466")
                    nu_nom_sec = st.text_input("Nombre del Sector (Para Gerentes/Líderes)", placeholder="ej. EMOCIONES DOLLY")
                    
                    btn_save_u = st.form_submit_button("💾 Guardar / Actualizar Usuario", type="primary", use_container_width=True)
                    if btn_save_u:
                        ok_u, msg_u = registrar_o_actualizar_usuario(
                            nu_username, nu_nombre, nu_pass, nu_rol, nu_grupo, nu_sector, nombre_sector=nu_nom_sec
                        )
                        if ok_u:
                            st.success(f"✅ {msg_u}")
                            st.rerun()
                        else:
                            st.error(f"❌ {msg_u}")
                            
            with tab_eliminar_u:
                st.markdown("##### 🗑️ Eliminar Cuenta de Usuario / Gerente")
                st.caption("Selecciona el perfil a remover. La cuenta del Super Administrador ('admin') está protegida.")
                
                users_disponibles_elim = [u for u in users_dict.keys() if u != 'admin']
                
                if users_disponibles_elim:
                    with st.form("form_eliminar_usuario_perfil"):
                        user_a_eliminar = st.selectbox(
                            "Selecciona el usuario a eliminar:",
                            options=users_disponibles_elim,
                            format_func=lambda u: f"👤 {u} — {users_dict[u].get('nombre', '')} ({users_dict[u].get('rol', '')})"
                        )
                        
                        u_info_sel = users_dict.get(user_a_eliminar, {})
                        es_gerente_sel = (u_info_sel.get("rol") == "gerente")
                        
                        chk_elim_sec = False
                        if es_gerente_sel and u_info_sel.get("codigo_sector"):
                            chk_elim_sec = st.checkbox(
                                f"🗑️ También eliminar el historial del Sector {u_info_sel.get('codigo_sector')} (permite que este sector pueda re-registrarse)",
                                value=True
                            )
                            
                        chk_conf_del_u = st.checkbox("🔒 Confirmo que deseo eliminar este usuario permanentemente", value=False)
                        
                        btn_del_u = st.form_submit_button("🚨 Eliminar Perfil Definitivamente", type="secondary", use_container_width=True)
                        
                        if btn_del_u:
                            if not chk_conf_del_u:
                                st.warning("⚠️ Debes marcar la casilla de confirmación para eliminar.")
                            else:
                                ok_del, msg_del = eliminar_usuario_perfil(user_a_eliminar, eliminar_sector_asociado=chk_elim_sec)
                                if ok_del:
                                    st.success(f"✅ {msg_del}")
                                    st.rerun()
                                else:
                                    st.error(f"❌ {msg_del}")
                else:
                    st.info("No hay usuarios adicionales disponibles para eliminar.")

        st.markdown("---")
        st.subheader("🎛️ Control Global de Permisos de Carga de Archivos")
        st.markdown("Configura si las Líderes de Negocio pueden subir o actualizar archivos Excel en la plataforma, o si esta función permanece restringida a la Gerencia General.")
        
        config_actual = cargar_configuracion()
        estado_permiso = config_actual.get("permitir_carga_lideres", False)
        
        col_p1, col_p2 = st.columns([1.5, 1])
        with col_p1:
            nuevo_permiso = st.toggle(
                "🔓 Permitir a las Líderes de Negocio subir/actualizar archivos",
                value=estado_permiso,
                help="Si está activado, las Líderes podrán ver las opciones para subir y actualizar archivos permitidos."
            )
            
            if nuevo_permiso != estado_permiso:
                config_actual["permitir_carga_lideres"] = nuevo_permiso
                guardar_configuracion(config_actual)
                st.success("✅ ¡Permisos de carga globales actualizados correctamente!")
                st.rerun()
        
        with col_p2:
            if estado_permiso:
                st.warning("⚠️ **Permisos Abiertos**: Las Líderes tienen acceso a subir archivos.")
            else:
                st.info("🔒 **Modo Protegido (Predeterminado)**: Las Líderes y Asesoras tienen bloqueadas las opciones de subida de archivos.")

        st.markdown("---")
        st.subheader("🧹 Mantenimiento & Limpieza de Base de Datos")
        st.markdown("Herramientas de administración para eliminar los datos de un grupo en específico o reiniciar la base de datos completa para un nuevo ciclo.")

        with st.expander("⚠️ Opciones Avanzadas de Borrado (Líder Específico / Base Completa)", expanded=False):
            col_b1, col_b2 = st.columns(2)

            # Opción 1: Borrado por Líder o Grupo Específico
            with col_b1:
                st.markdown("##### 👤 1. Borrar Datos de un Grupo / Líder Específico")
                st.caption("Elimina de SQLite las asesoras, facturación y comentarios de una líder determinada.")
                
                df_grupos_b = consultar_tableau_sql()
                lista_grupos_borrar = sorted([str(g).strip() for g in df_grupos_b['Grupo'].dropna().unique()]) if (df_grupos_b is not None and not df_grupos_b.empty and 'Grupo' in df_grupos_b.columns) else []
                
                grp_a_borrar = st.selectbox("Selecciona el Grupo / Líder a eliminar:", options=["-- Seleccionar --"] + lista_grupos_borrar, key="sel_grp_borrar")
                check_elim_cuenta = st.checkbox("También eliminar la cuenta de usuario (login)", value=False, key="chk_elim_cuenta")
                
                if grp_a_borrar != "-- Seleccionar --":
                    st.warning(f"⚠️ Estás a punto de borrar los datos del **Grupo {grp_a_borrar}**.")
                    if st.button("🗑️ Confirmar y Borrar Datos de este Grupo", type="secondary", key="btn_borrar_grp"):
                        res_del = eliminar_datos_por_grupo_o_usuario(grp_a_borrar, eliminar_cuenta=check_elim_cuenta)
                        st.success(f"✅ ¡Datos eliminados para el Grupo {grp_a_borrar}! Registros removidos: {res_del}")
                        st.cache_data.clear()
                        st.rerun()

            # Opción 2: Vaciar Base de Datos Completa
            with col_b2:
                st.markdown("##### 🚨 2. Vaciar Base de Datos Completa")
                st.caption("Reinicia todas las tablas SQLite y limpia los archivos Excel locales para iniciar un ciclo nuevo desde cero.")
                
                check_elim_excel = st.checkbox("Eliminar archivos Excel locales ('Cómo Vamos' y Tableau)", value=True, key="chk_elim_excel")
                check_vaciar_usuarios = st.checkbox("También eliminar cuentas de usuarios Líderes", value=False, key="chk_vaciar_users")
                confirmacion_seguridad = st.checkbox("🔒 Confirmo que deseo vaciar la Base de Datos completa", value=False, key="chk_conf_seg")
                
                if confirmacion_seguridad:
                    if st.button("🚨 VACIAR BASE DE DATOS COMPLETA AHORA", type="primary", key="btn_vaciar_db_all"):
                        res_vac = vaciar_base_datos_completa(vaciar_usuarios=check_vaciar_usuarios, eliminar_archivos_excel=check_elim_excel)
                        st.success(f"✅ ¡Base de datos y archivos vaciados exitosamente! Resumen: {res_vac}")
                        st.cache_data.clear()
                        st.rerun()
                else:
                    st.info("Marca la casilla de confirmación para habilitar el botón de vaciado.")

        st.markdown("---")
        st.subheader("🎛️ Control de Visibilidad y Accesos por Pestaña / Módulo")
        st.markdown("Configura de manera independiente qué pestañas y módulos son visibles para cada rol (**Gerentes**, **Líderes de Negocio** y **Asesoras**).")

        permisos_dict = config_actual.get("permisos_pestanas", DEFAULT_PERMISOS_PESTANAS)
        cambio_permisos = False

        c_th_name, c_th_ger, c_th_lid, c_th_ase = st.columns([2.5, 1, 1, 1])
        with c_th_name:
            st.markdown("##### 📌 Pestaña / Módulo")
        with c_th_ger:
            st.markdown("##### 👑 Gerentes")
        with c_th_lid:
            st.markdown("##### 👩‍💼 Líderes")
        with c_th_ase:
            st.markdown("##### 👤 Asesoras")

        st.markdown("<hr style='margin: 5px 0 15px 0;'>", unsafe_allow_html=True)

        for t_key, t_info in permisos_dict.items():
            t_nombre = t_info.get("nombre", t_key)
            val_ger = bool(t_info.get("gerente", True))
            val_lid = bool(t_info.get("lider", True))
            val_ase = bool(t_info.get("asesor", True))
            
            c1, c2, c3, c4 = st.columns([2.5, 1, 1, 1])
            with c1:
                st.markdown(f"**{t_nombre}**")
            with c2:
                new_ger = st.toggle("Gerente", value=val_ger, key=f"t_{t_key}_ger", label_visibility="collapsed")
            with c3:
                new_lid = st.toggle("Líder", value=val_lid, key=f"t_{t_key}_lid", label_visibility="collapsed")
            with c4:
                new_ase = st.toggle("Asesora", value=val_ase, key=f"t_{t_key}_ase", label_visibility="collapsed")
                
            if new_ger != val_ger or new_lid != val_lid or new_ase != val_ase:
                permisos_dict[t_key]["gerente"] = new_ger
                permisos_dict[t_key]["lider"] = new_lid
                permisos_dict[t_key]["asesor"] = new_ase
                cambio_permisos = True

        if cambio_permisos:
            config_actual["permisos_pestanas"] = permisos_dict
            guardar_configuracion(config_actual)
            st.success("✅ ¡Permisos de visibilidad por pestaña actualizados exitosamente!")
            st.rerun()

# --- TAB: DIRECTORIO & GESTIÓN DE MIS LÍDERES (EXCLUSIVO GERENTES) ---
with tab_lideres_gerente:
    st.subheader(f"👥 Directorio de Mis Líderes & Gestión de Accesos")
    st.markdown(f"Administración centralizada de usuarios para las líderes de tu Sector **{user_sector if user_sector else 'General'}** (*{user_nombre}*). Consulta sus datos de acceso, descarga el archivo de respaldo o restablece contraseñas en 1 clic.")

    users_dict = cargar_usuarios()
    # Filtrar líderes de este sector
    lideres_sector = []
    for uname, udata in users_dict.items():
        es_de_sector = True if not user_sector else (str(udata.get("codigo_sector", "")).strip() == str(user_sector).strip())
        if udata.get("rol") == "lider" and es_de_sector:
            lideres_sector.append({
                "Usuario (Login)": uname,
                "Nombre Líder": udata.get("nombre", ""),
                "Código de Grupo": str(udata.get("codigo_grupo", "")),
                "Estado Suscripción": udata.get("estado_suscripcion", "activo"),
                "Debe Cambiar Clave": "Sí" if udata.get("debe_cambiar_password") else "No"
            })

    col_ger_u1, col_ger_u2 = st.columns([1.3, 1])

    with col_ger_u1:
        st.markdown("##### 📋 Listado de Líderes Registradas en tu Sector")
        if lideres_sector:
            df_lid_sec = pd.DataFrame(lideres_sector)
            st.dataframe(df_lid_sec, use_container_width=True, hide_index=True)

            csv_lid = df_lid_sec.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Descargar Directorio de Mis Líderes (CSV / Excel)",
                data=csv_lid,
                file_name=f"Directorio_Lideres_Sector_{user_sector if user_sector else 'General'}.csv",
                mime="text/csv",
                use_container_width=True,
                key="btn_descargar_lideres_gerente"
            )
        else:
            st.info(f"No hay cuentas de líderes registradas asociadas al código de sector {user_sector}.")

    with col_ger_u2:
        st.markdown("##### ⚡ Restablecer Contraseña de una Líder")
        st.caption("Si una líder olvidó su contraseña o perdió sus datos, selecciónala y restablécela a **'lider123'** de inmediato:")

        if lideres_sector:
            with st.form("form_reset_gerente"):
                u_sel_ger = st.selectbox(
                    "Selecciona la Líder a gestionar:",
                    options=[l["Usuario (Login)"] for l in lideres_sector],
                    format_func=lambda u: f"👤 {u} — {users_dict[u].get('nombre', '')} (Grupo {users_dict[u].get('codigo_grupo', '')})"
                )
                pass_nueva_ger = st.text_input("Nueva Contraseña:", value="lider123")
                btn_reset_ger = st.form_submit_button("🔄 Restablecer Contraseña", type="primary", use_container_width=True)

                if btn_reset_ger:
                    ok_r_g, msg_r_g = restablecer_password_usuario(u_sel_ger, pass_nueva_ger, debe_cambiar=False)
                    if ok_r_g:
                        st.success(f"✅ ¡Listo! La contraseña de **{u_sel_ger}** ahora es: `{pass_nueva_ger}`")
                        st.session_state['ultimo_reseteo_gerente'] = {
                            'usuario': u_sel_ger,
                            'nombre': users_dict[u_sel_ger].get('nombre', ''),
                            'grupo': users_dict[u_sel_ger].get('codigo_grupo', ''),
                            'password': pass_nueva_ger
                        }
                        st.rerun()
                    else:
                        st.error(f"❌ {msg_r_g}")

            # Sub-panel de WhatsApp directo
            ult_g = st.session_state.get('ultimo_reseteo_gerente')
            if ult_g:
                st.markdown("---")
                st.markdown("###### 📲 Enviar Credenciales a la Líder por WhatsApp")
                # Intentar buscar celular de la líder desde df_tableau
                cel_auto = ""
                if 'df_tableau' in locals() and df_tableau is not None and not df_tableau.empty:
                    match_l = df_tableau[df_tableau['Grupo'].astype(str) == str(ult_g['grupo'])]
                    if not match_l.empty and 'celular' in match_l.columns:
                        cel_val = str(match_l['celular'].iloc[0]).replace('.0', '').strip()
                        cel_auto = "".join(ch for ch in cel_val if ch.isdigit())

                tel_ger_in = st.text_input("Número celular de la líder (10 dígitos):", value=cel_auto, key="tel_ger_reset_in")
                msg_wa_ger = (
                    f"🌸 ¡Hola {ult_g['nombre'].split()[0].title() if ult_g['nombre'] else 'Líder'}! Te comparto tus credenciales de acceso al Sistema de Gestión Natura & Avon:\n\n"
                    f"👤 *Usuario:* `{ult_g['usuario']}`\n"
                    f"🔑 *Contraseña:* `{ult_g['password']}`\n\n"
                    f"🌐 *Enlace de Ingreso:* https://app-lideres-production.up.railway.app\n\n"
                    f"¡Muchos éxitos! ✨ — Tu Gerente {user_nombre}"
                )
                st.text_area("Mensaje listo para WhatsApp:", msg_wa_ger, height=120, key="txt_wa_ger_msg")
                if tel_ger_in and len(tel_ger_in.strip()) >= 10:
                    link_wa_g = f"https://api.whatsapp.com/send?phone=57{tel_ger_in.strip()}&text={urllib.parse.quote(msg_wa_ger)}"
                    st.link_button("📲 Enviar Datos por WhatsApp a la Líder", url=link_wa_g, use_container_width=True)
                else:
                    st.caption("💡 Ingresa el número celular para habilitar el botón de WhatsApp.")

# Footer
st.markdown("---")
st.caption(f"📈 Panel de Control {user_sector_nombre} | Desarrollado por Tao-System")
