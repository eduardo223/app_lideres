import streamlit as st
import os
import urllib.parse
from datetime import datetime
import numpy as np

try:
    import cv2
except ImportError:
    cv2 = None

from procesador import (
    obtener_tarifa_gerente,
    cargar_historico_sectores,
    vincular_referida,
    registrar_reporte_pago,
    obtener_nombre_sector_usuario,
    verificar_estado_suscripcion
)

def renderizar_contenido_pago_gerente(current_user, user_sector, user_nombre):
    """
    Renderiza la interfaz completa de pago de suscripción para Gerentes:
    - Desglose de cuota mensual con descuento dinámico por referidas activas
    - Código de embajadora / patrocinadora
    - Pestañas con medios de pago: Nequi (Bre-B y Personal en fondo blanco), Daviplata (Casita Bre-B ampliada), Bancolombia y Llave Bre-B
    - Reporte de pago ágil por comprobante con lectura automática de código QR de aprobación
    """
    tarifa_info = obtener_tarifa_gerente(user_sector or current_user)
    codigo_ref = tarifa_info.get("codigo_embajadora", f"REF-{user_sector}")
    base_monto = tarifa_info.get("cuota_base", 100000.0)
    dcto_pct = tarifa_info.get("porcentaje_descuento", 0.0)
    dcto_monto = tarifa_info.get("monto_descuento", 0.0)
    total_pagar = tarifa_info.get("total_a_pagar", 100000)
    num_activas = tarifa_info.get("num_referidas_activas", 0)
    num_totales = tarifa_info.get("num_referidas_totales", 0)
    referidas_list = tarifa_info.get("referidas_detalle", [])

    st.markdown(f"""
    <div style="background: linear-gradient(135deg, rgba(227, 0, 123, 0.08) 0%, rgba(255, 107, 0, 0.08) 100%);
                border: 1px solid rgba(227, 0, 123, 0.28); border-radius: 16px; padding: 16px 20px; margin-bottom: 16px;
                box-shadow: 0 4px 16px rgba(227, 0, 123, 0.08);">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
            <div>
                <span style="font-size: 0.8rem; text-transform: uppercase; font-weight: 800; color: #E3007B; letter-spacing: 0.5px;">💎 Suscripción Mensual de Sector</span>
                <h3 style="margin: 2px 0 0 0; color: #1E293B; font-size: 1.35rem;">{obtener_nombre_sector_usuario(current_user)}</h3>
                <span style="font-size: 0.88rem; color: #64748B;">Tarifa Plena Oficial: <b>${base_monto:,.0f} COP</b></span>
            </div>
            <div style="text-align: right;">
                <span style="font-size: 0.82rem; color: #059669; font-weight: 700; background: #ECFDF5; padding: 3px 8px; border-radius: 6px; border: 1px solid #A7F3D0;">
                    🎁 {dcto_pct:.0f}% OFF ({num_activas} referida(s) activa(s))
                </span>
                <div style="font-size: 1.65rem; font-weight: 900; color: #059669; margin-top: 4px;">${total_pagar:,.0f} COP</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander(f"👥 Tu Red de Colegas Referidas ({num_activas} activas de {num_totales} registradas • {dcto_pct:.0f}% dcto)", expanded=(num_totales > 0)):
        st.markdown(f"**Tu Código de Embajadora para Compartir:** <code style='font-size: 1.1rem; color: #E3007B;'>{codigo_ref}</code>", unsafe_allow_html=True)
        st.caption("💡 **Regla de Oro:** Por cada colega Gerente que active su sector con tu código, obtienes **5% de descuento recurrente** mientras ella se mantenga activa y al día. Si una colega se atrasa en su mensualidad, el beneficio se pausa automáticamente y tu cuota se recalcula.")
        if referidas_list:
            for r in referidas_list:
                c_nom = r.get('nombre_gerente', 'Gerente')
                c_sec = r.get('nombre_sector', f"Sector {r.get('codigo_sector')}")
                st.markdown(f"- **{c_sec}** ({c_nom}): {r.get('estado_str')}")
        else:
            st.info("Aún no tienes colegas registradas con tu código. ¡Comparte tu código con otras Gerentes de Natura y reduce tu mensualidad todos los meses!")

        sec_str = str(user_sector or "").strip().replace(".0", "")
        hist_ref = cargar_historico_sectores()
        referido_actual = hist_ref.get(sec_str, {}).get("referido_por")
        if not referido_actual:
            with st.popover("🎁 ¿Alguien te recomendó esta plataforma? Ingresa su código aquí"):
                cod_input_ingresar = st.text_input("Código de la Gerente que te refirió (ej: REF-...):", key="input_cod_patrocinadora_modal")
                if st.button("Vincular Patrocinadora", key="btn_vincular_patrocinadora_modal"):
                    ok_v, msg_v = vincular_referida(sec_str, cod_input_ingresar)
                    if ok_v:
                        st.success(f"✅ {msg_v}")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg_v}")
        else:
            st.caption(f"🤝 Vinculada con la embajadora del Sector: `{referido_actual}`")

    st.markdown("##### 💳 Selecciona tu Método de Pago Preferido:")
    tab_neq, tab_davi, tab_banco, tab_llave, tab_rep = st.tabs([
        "🟣 Nequi", "🔴 Daviplata", "🟡 Bancolombia", "🔵 Llave Bre-B", "📝 Reportar Pago"
    ])

    with tab_neq:
        st.caption("Escanea con tu aplicación de Nequi o billetera digital a la llave Bre-B:")
        c_q1, c_q2 = st.columns(2)
        with c_q1:
            st.markdown("<p style='font-size:0.88rem; font-weight:800; color:#E3007B; margin-bottom:6px; text-align:center;'>💳 QR Bre-B Oficial ($100.000 COP)</p>", unsafe_allow_html=True)
            p_q1 = "data/assets/qr_nequi_breb.png" if os.path.exists("data/assets/qr_nequi_breb.png") else ("assets/qr_nequi_breb.png" if os.path.exists("assets/qr_nequi_breb.png") else None)
            if p_q1:
                st.image(p_q1, use_container_width=True)
            else:
                st.info("Paga por Nequi al número 3057939537.")
        with c_q2:
            st.markdown("<p style='font-size:0.88rem; font-weight:800; color:#E3007B; margin-bottom:6px; text-align:center;'>📱 Tu QR Nequi Personal</p>", unsafe_allow_html=True)
            p_q2 = "data/assets/qr_nequi_personal.png" if os.path.exists("data/assets/qr_nequi_personal.png") else ("assets/qr_nequi_personal.png" if os.path.exists("assets/qr_nequi_personal.png") else None)
            if p_q2:
                st.image(p_q2, use_container_width=True)
            else:
                st.info("Paga por Nequi al número 3057939537.")

        st.markdown(f"""
        <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 14px; margin-top: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.04);">
            <div style="color: #334155; font-size: 0.95rem;"><b>Titular:</b> Raul Sanchez</div>
            <div style="color: #334155; font-size: 1.05rem; margin: 4px 0;"><b>Celular / Llave Nequi:</b> <code style="font-size: 1.15rem; font-weight: 800; color: #E3007B;">3057939537</code></div>
            <div style="color: #64748B; font-size: 0.85rem;">En la descripción de pago puedes colocar: <code>Sector {user_sector}</code></div>
        </div>
        """, unsafe_allow_html=True)

    with tab_davi:
        st.caption("Escanea desde tu app Daviplata o cualquier billetera digital con Bre-B:")
        p_qd = "data/assets/qr_daviplata.png" if os.path.exists("data/assets/qr_daviplata.png") else ("assets/qr_daviplata.png" if os.path.exists("assets/qr_daviplata.png") else None)
        if p_qd:
            col_dv1, col_dv2, col_dv3 = st.columns([1, 2.8, 1])
            with col_dv2:
                st.image(p_qd, use_container_width=True)

        st.markdown(f"""
        <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 14px; margin-top: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.04);">
            <div style="color: #334155; font-size: 0.95rem;"><b>Titular:</b> Raul Eduardo Sanchez</div>
            <div style="color: #334155; font-size: 1.05rem; margin: 4px 0;"><b>Celular / Llave Daviplata:</b> <code style="font-size: 1.15rem; font-weight: 800; color: #DC2626;">3057939537</code></div>
            <div style="color: #64748B; font-size: 0.85rem;">En la descripción de pago puedes colocar: <code>Sector {user_sector}</code></div>
        </div>
        """, unsafe_allow_html=True)

    with tab_banco:
        st.caption("Transferencia directa desde App Bancolombia, cajero o sucursal virtual:")
        st.markdown("""
        <div style="background: linear-gradient(135deg, #FEF3C7 0%, #FFFBEB 100%); border: 1px solid #FCD34D; border-radius: 14px; padding: 20px; margin: 10px 0; box-shadow: 0 4px 12px rgba(217, 119, 6, 0.08);">
            <div style="display: flex; align-items: center; justify-content: space-between;">
                <div style="font-size: 1.25rem; font-weight: 900; color: #92400E;">🏦 Bancolombia</div>
                <span style="background: #FDE68A; color: #78350F; font-size: 0.75rem; font-weight: 800; padding: 3px 8px; border-radius: 6px;">CUENTA OFICIAL</span>
            </div>
            <div style="margin-top: 12px;">
                <p style="margin: 4px 0; color: #78350F; font-size: 0.95rem;"><b>Tipo de Cuenta:</b> Cuenta de Ahorros</p>
                <p style="margin: 6px 0; color: #78350F; font-size: 1.05rem;"><b>Número de Cuenta:</b></p>
                <div style="background: white; border: 1.5px dashed #D97706; border-radius: 8px; padding: 8px 14px; display: inline-block;">
                    <code style="font-size: 1.35rem; font-weight: 900; color: #1E293B; letter-spacing: 1px;">58947425836</code>
                </div>
                <p style="margin: 8px 0 0 0; color: #78350F; font-size: 0.95rem;"><b>Titular:</b> Raul Eduardo Sanchez</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with tab_llave:
        st.caption("Transferencias inmediatas interbancarias Bre-B sin costo:")
        st.markdown("""
        <div style="background: linear-gradient(135deg, #F0F9FF 0%, #FFFFFF 100%); border: 1.5px solid #BAE6FD; border-radius: 14px; padding: 20px; margin: 10px 0; box-shadow: 0 4px 12px rgba(2, 132, 199, 0.06);">
            <div style="font-size: 1.25rem; font-weight: 900; color: #0369A1;">⚡ Llave Bre-B</div>
            <p style="margin: 8px 0 4px 0; color: #0C4A6E; font-size: 0.95rem;">Válido desde cualquier banco o billetera digital que soporte Bre-B (Davivienda, BBVA, Banco de Bogotá, Scotiabank, Nequi, Daviplata, etc.):</p>
            <div style="background: white; border: 1.5px dashed #0284C7; border-radius: 10px; padding: 10px 18px; display: inline-block; margin: 8px 0;">
                <code style="font-size: 1.45rem; font-weight: 900; color: #0284C7; letter-spacing: 1px;">3057939537</code>
            </div>
            <p style="margin: 4px 0; color: #0C4A6E; font-size: 0.95rem;"><b>Destinatario:</b> Raul Eduardo Sanchez</p>
        </div>
        """, unsafe_allow_html=True)

    with tab_rep:
        st.caption("Sube la captura de tu comprobante. El sistema leerá automáticamente la información:")
        with st.form("form_reporte_pago_gerente", clear_on_submit=False):
            f_metodo = st.selectbox("Medio de Pago Utilizado:", ["Nequi", "Daviplata", "Bancolombia Ahorros", "Llave Bre-B", "Otro"])
            f_monto = st.number_input("Valor Pagado ($ COP):", value=float(total_pagar), step=5000.0)
            f_file = st.file_uploader("📎 Adjuntar Comprobante o Captura de Pantalla del Pago:", type=["png", "jpg", "jpeg", "webp", "pdf"], key="upload_comprobante_gerente", help="Sube la foto del comprobante de transferencia. El sistema analizará la imagen y el código QR de confirmación.")
            f_notas = st.text_input("Nota o mensaje adicional (opcional):", placeholder=f"Pago mensualidad {obtener_nombre_sector_usuario(current_user)}")

            btn_enviar_reporte = st.form_submit_button("🚀 Confirmar Pago y Notificar por WhatsApp", type="primary", use_container_width=True)
            if btn_enviar_reporte:
                if not f_file:
                    st.warning("⚠️ Por favor adjunta la captura de pantalla o imagen de tu comprobante de pago.")
                else:
                    file_name_saved = None
                    codigo_qr_detectado = None

                    try:
                        comp_dir = "data/comprobantes"
                        os.makedirs(comp_dir, exist_ok=True)
                        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                        file_name_saved = f"pago_{user_sector}_{timestamp_str}_{f_file.name}"
                        file_bytes = f_file.getbuffer()
                        with open(os.path.join(comp_dir, file_name_saved), "wb") as f_dst:
                            f_dst.write(file_bytes)

                        # Escaneo inteligente con OpenCV para detectar y leer QR en el comprobante
                        if cv2 is not None and f_file.type and "image" in f_file.type:
                            try:
                                img_np = np.frombuffer(file_bytes, np.uint8)
                                cv_img = cv2.imdecode(img_np, cv2.IMREAD_COLOR)
                                if cv_img is not None:
                                    det = cv2.QRCodeDetector()
                                    val_qr, _, _ = det.detectAndDecode(cv_img)
                                    if val_qr and str(val_qr).strip():
                                        codigo_qr_detectado = str(val_qr).strip()
                            except Exception:
                                pass
                    except Exception:
                        pass

                    ok_rep, nuevo_p = registrar_reporte_pago(
                        user_sector,
                        user_nombre,
                        f_metodo,
                        referencia=(codigo_qr_detectado[:20] if codigo_qr_detectado else ""),
                        valor=f_monto,
                        notas=f_notas,
                        comprobante_nombre=file_name_saved,
                        codigo_qr_detectado=codigo_qr_detectado
                    )

                    if ok_rep:
                        st.success("✅ ¡Comprobante registrado exitosamente en el sistema!")
                        if codigo_qr_detectado:
                            st.info(f"🔍 **Código QR del comprobante verificado automáticamente:** `{codigo_qr_detectado[:35]}...`")

                        texto_wa = (
                            f"Hola Raúl, soy {user_nombre} del Sector {user_sector} ({obtener_nombre_sector_usuario(current_user)}).\n\n"
                            f"Acabo de realizar el pago de mi mensualidad:\n"
                            f"💰 Monto: ${f_monto:,.0f} COP\n"
                            f"💳 Canal: {f_metodo}\n"
                            f"📎 Comprobante adjuntado en la plataforma.\n\n"
                            f"Quedo atenta a la confirmación en plataforma. ¡Muchas gracias!"
                        )
                        wa_url = f"https://wa.me/573057939537?text={urllib.parse.quote(texto_wa)}"
                        st.markdown(f"""
                        <a href="{wa_url}" target="_blank" style="text-decoration: none;">
                            <button style="background: linear-gradient(135deg, #25D366 0%, #128C7E 100%); color: white; border: none; padding: 14px 20px; border-radius: 10px; font-weight: 800; font-size: 1.05rem; width: 100%; cursor: pointer; margin-top: 10px;">
                                💬 Abrir WhatsApp y Enviar Notificación a Raúl
                            </button>
                        </a>
                        """, unsafe_allow_html=True)

if hasattr(st, 'dialog'):
    @st.dialog("💳 Centro de Pagos & Suscripción de Sector", width="large")
    def modal_pago_suscripcion_gerente(current_user, user_sector, user_nombre):
        renderizar_contenido_pago_gerente(current_user, user_sector, user_nombre)
else:
    def modal_pago_suscripcion_gerente(current_user, user_sector, user_nombre):
        with st.expander("💳 Centro de Pagos & Suscripción de Sector", expanded=True):
            renderizar_contenido_pago_gerente(current_user, user_sector, user_nombre)

def banner_alerta_suscripcion_gerente(current_user, user_sector, user_nombre, es_movil=False):
    """
    Renderiza un banner elegante y no invasivo ÚNICAMENTE si quedan 3 días o menos
    de suscripción o si la suscripción ya está vencida.
    """
    info_susc = verificar_estado_suscripcion(current_user)
    dias_rest = info_susc.get('dias_restantes', 9999)
    estado = info_susc.get('estado', 'activo')
    vence_str = info_susc.get('fecha_vencimiento_str', '')

    if dias_rest <= 3 or estado in ['vencido', 'prueba']:
        if es_movil:
            st.markdown(f"""
            <div style="background: linear-gradient(135deg, rgba(245, 158, 11, 0.12) 0%, rgba(255, 255, 255, 0.95) 100%);
                        border: 1px solid rgba(245, 158, 11, 0.4); border-radius: 12px; padding: 10px 14px; margin-bottom: 10px;">
                <div style="font-weight: 800; color: #D97706; font-size: 0.95rem;">✨ Suscripción de Sector</div>
                <div style="color: #475569; font-size: 0.82rem; margin: 2px 0 6px 0;">
                    {'Tu periodo concluyó el <b>' + vence_str + '</b>' if estado == 'vencido' else 'Vence en <b>' + str(dias_rest) + ' día(s)</b> (' + vence_str + ')'}.
                </div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("💳 Pagar / Renovar", key="btn_renovar_mob_banner", type="primary", use_container_width=True):
                modal_pago_suscripcion_gerente(current_user, user_sector, user_nombre)
        else:
            col_sb_av1, col_sb_av2 = st.columns([3.4, 1.2])
            with col_sb_av1:
                if estado == 'vencido':
                    st.markdown(f"""
                    <div style="background: linear-gradient(135deg, rgba(239, 68, 68, 0.12) 0%, rgba(255, 255, 255, 0.95) 100%);
                                border: 1px solid rgba(239, 68, 68, 0.35); border-radius: 12px; padding: 10px 16px; margin-bottom: 12px;">
                        <span style="font-weight: 800; color: #DC2626;">⚠️ Suscripción Finalizada:</span>
                        <span style="color: #475569; font-size: 0.9rem; margin-left: 6px;">Tu periodo concluyó el <b>{vence_str}</b>. Renueva tu mes acumulativo para mantener a tu equipo activo.</span>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div style="background: linear-gradient(135deg, rgba(245, 158, 11, 0.12) 0%, rgba(255, 255, 255, 0.95) 100%);
                                border: 1px solid rgba(245, 158, 11, 0.35); border-radius: 12px; padding: 10px 16px; margin-bottom: 12px;">
                        <span style="font-weight: 800; color: #D97706;">✨ Suscripción de Sector:</span>
                        <span style="color: #475569; font-size: 0.9rem; margin-left: 6px;">Vence en <b>{dias_rest} día(s)</b> ({vence_str}). Tu cuota tiene descuento automático por referidas.</span>
                    </div>
                    """, unsafe_allow_html=True)
            with col_sb_av2:
                if st.button("💳 Pagar / Renovar", key="btn_renovar_susc_banner_top", type="primary", use_container_width=True):
                    modal_pago_suscripcion_gerente(current_user, user_sector, user_nombre)
