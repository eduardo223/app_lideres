import pandas as pd
import os
import io
import sys
import json

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

def safe_print(*args, **kwargs):
    try:
        print(*args, **kwargs)
    except Exception:
        try:
            clean_args = [str(a).encode('ascii', 'ignore').decode('ascii') for a in args]
            print(*clean_args, **kwargs)
        except Exception:
            pass

def limpiar_serie_numerica(series):
    if series is None or series.empty:
        return series
    def _limpiar(val):
        if pd.isna(val):
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        try:
            s = str(val).replace('$', '').replace(',', '').replace(' ', '').strip()
            return float(s)
        except:
            return 0.0
    return series.apply(_limpiar)

def normalizar_columnas(df):
    """
    Limpia y estandariza los nombres de las columnas del DataFrame de Excel,
    corregiendo la codificación de caracteres especiales (ej. 'Cdigo' -> 'Código').
    """
    columnas_limpias = []
    for c in df.columns:
        # Reemplazar el carácter especial de reemplazo \ufffd por 'ó'
        s = str(c).replace('\ufffd', 'ó')
        
        # Mapeos específicos de respaldo
        if 'digo' in s and 'consultora' in s:
            s = 'Código de consultora'
        elif 'digo' in s and 'grupo' in s:
            s = 'Código de grupo'
        elif 'Facturac' in s and '_anterior' not in s:
            if 'Objetivo' in s:
                s = 'Objetivo Facturación'
            elif 'Real' in s:
                s = 'Real Facturación'
            elif 'Cumplimiento' in s:
                s = 'Cumplimiento Facturación'
        columnas_limpias.append(s.strip())
        
    df.columns = columnas_limpias

    # Limpiar automáticamente columnas numéricas conocidas para eliminar el '$' de los textos
    cols_numericas = [
        'Objetivo Facturación', 'Real Facturación', 'Objetivo Activas', 'Real Activas',
        'Cumplimiento Facturación', 'Cumplimiento Activas', 'Saldo', 'Disponibles',
        'Inicios', 'Reinicios', 'Recuperos', 'Inactiva 1', 'Inactiva 2', 'Inactiva 3'
    ]
    for col in cols_numericas:
        if col in df.columns:
            df[col] = limpiar_serie_numerica(df[col])

    return df

def calcular_metas_ciclo(origen='Base para el como vamos.xlsx'):
    """
    Calcula las metas de crecimiento (+1, +3, +5, +7, +9) basándose en las activas reales
    y calcula el 'Avance % Facturación' comparando con la hoja 'Como vamos anterior'.
    """
    if isinstance(origen, str):
        if not os.path.exists(origen):
            print(f"Error: No se encontró el archivo '{origen}'.")
            return None
        print("Iniciando lectura de la base de datos...")
        df = pd.read_excel(origen, sheet_name="Base para el como vamos")
        df = normalizar_columnas(df)
        
        # Eliminar filas 'None', 'nan' o totalmente nulas que dañen la presentación de tablas
        col_nom_ref = 'Nombre de consultora' if 'Nombre de consultora' in df.columns else df.columns[0]
        if col_nom_ref in df.columns:
            mask_valida_df = df[col_nom_ref].notna() & (~df[col_nom_ref].astype(str).str.strip().str.lower().isin(['none', 'nan', '', 'null', '0']))
            df = df[mask_valida_df]
        
        # Verificar si existe la hoja 'Como vamos anterior' para calcular el Avance %
        try:
            xl = pd.ExcelFile(origen)
            hojas_lower = {s.lower().strip(): s for s in xl.sheet_names}
            if 'como vamos anterior' in hojas_lower:
                nombre_hoja_anterior = hojas_lower['como vamos anterior']
                df_prev = pd.read_excel(origen, sheet_name=nombre_hoja_anterior)
                df_prev = normalizar_columnas(df_prev)
                
                col_id = 'Código de consultora' if 'Código de consultora' in df_prev.columns else df_prev.columns[0]
                cols_prev = [c for c in [col_id, 'Cumplimiento Facturación', 'Real Facturación'] if c in df_prev.columns]
                
                if col_id in cols_prev and 'Cumplimiento Facturación' in cols_prev:
                    df_prev_sub = df_prev[cols_prev].drop_duplicates(subset=[col_id])
                    df_prev_sub = df_prev_sub.rename(columns={
                        'Cumplimiento Facturación': 'Cumplimiento Facturación_anterior',
                        'Real Facturación': 'Real Facturación_anterior'
                    })
                    
                    df = pd.merge(df, df_prev_sub, on=col_id, how='left')
        except Exception as e:
            print(f"Nota: No se pudo cargar la hoja 'Como vamos anterior': {e}")

    elif isinstance(origen, pd.DataFrame):
        df = normalizar_columnas(origen.copy())
    else:
        print("Error: El origen debe ser una ruta de archivo o un DataFrame de pandas.")
        return None
    
    # Nombre de la columna de activas de referencia
    col_activas = 'Real Activas' if 'Real Activas' in df.columns else 'Objetivo Activas'
    
    if col_activas in df.columns:
        # Cálculo de Metas de Crecimiento sobre las Activas Reales
        df['Meta_Crecer_1plus_150k'] = df[col_activas] + 1
        df['Meta_Crecer_3plus_200k'] = df[col_activas] + 3
        df['Meta_Crecer_5plus_300k'] = df[col_activas] + 5
        df['Meta_Crecer_7plus_500k'] = df[col_activas] + 7
        df['Meta_Crecer_9plus_750k'] = df[col_activas] + 9
    else:
        print(f"Advertencia: No se encontró la columna '{col_activas}'. Usando valor 0 como base.")
        df['Meta_Crecer_1plus_150k'] = 1
        df['Meta_Crecer_3plus_200k'] = 3
        df['Meta_Crecer_5plus_300k'] = 5
        df['Meta_Crecer_7plus_500k'] = 7
        df['Meta_Crecer_9plus_750k'] = 9

    # Cálculo del Avance % Facturación (Cumplimiento Actual - Cumplimiento Anterior)
    if 'Cumplimiento Facturación_anterior' in df.columns and 'Cumplimiento Facturación' in df.columns:
        c_actual = pd.to_numeric(df['Cumplimiento Facturación'], errors='coerce')
        c_anterior = pd.to_numeric(df['Cumplimiento Facturación_anterior'], errors='coerce')
        df['Avance % Facturación'] = c_actual - c_anterior
        
    if 'Real Facturación_anterior' in df.columns and 'Real Facturación' in df.columns:
        r_actual = pd.to_numeric(df['Real Facturación'], errors='coerce')
        r_anterior = pd.to_numeric(df['Real Facturación_anterior'], errors='coerce')
        df['Avance Facturación COP'] = r_actual - r_anterior

    # Cálculo de la columna 'Falta para el 100%', 'Falta para el 110%' y 'Cumplimiento Facturación' (en escala 0-100%)
    if 'Objetivo Facturación' in df.columns and 'Real Facturación' in df.columns:
        obj_f = pd.to_numeric(df['Objetivo Facturación'], errors='coerce')
        real_f = pd.to_numeric(df['Real Facturación'], errors='coerce')
        df['Falta para el 100%'] = obj_f - real_f
        df['Brecha Meta 100%'] = df['Falta para el 100%']
        df['Falta para el 110%'] = (obj_f * 1.10) - real_f
        df['Cumplimiento Facturación'] = (real_f / obj_f * 100.0).where(obj_f > 0, 0.0)

    # Cálculo de Productividad (Real Facturación / Real Activas)
    if 'Real Facturación' in df.columns and 'Real Activas' in df.columns:
        real_f = pd.to_numeric(df['Real Facturación'], errors='coerce')
        real_a = pd.to_numeric(df['Real Activas'], errors='coerce')
        df['Productividad'] = (real_f / real_a).fillna(0.0)

    # Cálculo dinámico de Ganancia Estimada según Matriz y Potencializador de Saldo
    df = calcular_ganancia_estimada_df(df)

    # Si se ejecuta directamente desde archivo, exportamos
    if isinstance(origen, str):
        archivo_salida = 'Resultado_Metas_Procesadas.xlsx'
        try:
            df.to_excel(archivo_salida, index=False)
            safe_print(f"[OK] Resultados guardados exitosamente en: '{archivo_salida}'!")
        except Exception as e:
            safe_print(f"Advertencia al guardar archivo de salida: {e}")

    return df

# --- REGLAS DE GANANCIA ESTIMADA SEGÚN MATRIZ Y POTENCIALIZADOR DE SALDO ---

MATRIZ_GANANCIA = [
    # <94.99% Fact | 95-99.99% Fact | 100-109.99% Fact | >=110% Fact
    [0.0150, 0.0160, 0.0170, 0.0180],  # <95% Activas
    [0.0160, 0.0275, 0.0400, 0.0450],  # 95% a 97.49% Activas
    [0.0170, 0.0300, 0.0425, 0.0525],  # 97.5% a 99.99% Activas
    [0.0180, 0.0325, 0.0500, 0.0575],  # 100% a 102.45% Activas
    [0.0190, 0.0340, 0.0525, 0.0600],  # 102.5% a 104.99% Activas
    [0.0200, 0.0355, 0.0550, 0.0625],  # 105% a 109.99% Activas
    [0.0210, 0.0370, 0.0575, 0.0650],  # >=110% Activas
]

ETIQUETAS_ACTIVAS = [
    "Menor a 95%", "95% a 97,49%", "97,5% a 99,99%",
    "100% a 102,45%", "102,5% a 104,99%", "105% a 109,99%", "110% a más"
]

ETIQUETAS_FACTURACION = [
    "Menor a 94,99%", "95% a 99,99%", "100% a 109,99%", "110% a más"
]

def obtener_indice_activas(cump_activas):
    if cump_activas < 0.95:
        return 0
    elif cump_activas < 0.975:
        return 1
    elif cump_activas < 1.00:
        return 2
    elif cump_activas < 1.025:
        return 3
    elif cump_activas < 1.05:
        return 4
    elif cump_activas < 1.10:
        return 5
    else:
        return 6

def obtener_indice_facturacion(cump_fact):
    if cump_fact < 0.95:
        return 0
    elif cump_fact < 1.00:
        return 1
    elif cump_fact < 1.10:
        return 2
    else:
        return 3

def calcular_matriz_ganancia(cump_activas, cump_fact, inicios=4):
    idx_act = obtener_indice_activas(cump_activas)
    idx_fact = obtener_indice_facturacion(cump_fact)
    pct_base = MATRIZ_GANANCIA[idx_act][idx_fact]
    
    # Penalización por Inicios < 4 (-0.5%)
    penalizacion_inicios = 0.0
    if inicios < 4:
        penalizacion_inicios = 0.005
        pct_base = max(0.0, pct_base - penalizacion_inicios)
        
    return pct_base, idx_act, idx_fact

def limpiar_numero(val, default=0.0):
    if pd.isna(val):
        return default
    if isinstance(val, (int, float)):
        return float(val)
    try:
        s = str(val).replace('$', '').replace(',', '').replace(' ', '').strip()
        return float(s)
    except:
        return default

def obtener_potencializador_saldo(saldo):
    s = limpiar_numero(saldo, 0.0)

    if s <= -4:
        return -0.30
    elif s <= -2:
        return -0.25
    elif s == -1:
        return -0.20
    elif s == 0:
        return -0.15
    elif s == 1:
        return -0.05
    elif s == 2:
        return 0.00
    elif s == 3:
        return 0.05
    elif s == 4:
        return 0.10
    elif s == 5:
        return 0.15
    elif s <= 7:
        return 0.20
    elif s <= 9:
        return 0.25
    else:
        return 0.30

def calcular_ganancia_estimada_df(df):
    ganancias_total = []
    ganancias_matriz = []
    pct_matriz_list = []
    potencializador_list = []
    potencializador_cop_list = []

    for _, row in df.iterrows():
        real_fact = limpiar_numero(row.get('Real Facturación', 0.0))
        obj_fact = limpiar_numero(row.get('Objetivo Facturación', 0.0))
        cump_fact = (real_fact / obj_fact) if obj_fact > 0 else 0.0
        
        real_act = limpiar_numero(row.get('Real Activas', 0.0))
        obj_act = limpiar_numero(row.get('Objetivo Activas', 0.0))
        cump_act = (real_act / obj_act) if obj_act > 0 else 0.0
        
        inicios = limpiar_numero(row.get('Inicios', 4.0), 4.0)
        saldo = limpiar_numero(row.get('Saldo', 0.0))
        
        pct_matriz, _, _ = calcular_matriz_ganancia(cump_act, cump_fact, inicios)
        ganancia_matriz_cop = real_fact * pct_matriz
        
        pct_potencializador = obtener_potencializador_saldo(saldo)
        potencializador_cop = ganancia_matriz_cop * pct_potencializador
        
        ganancia_total = ganancia_matriz_cop + potencializador_cop
        
        ganancias_total.append(ganancia_total)
        ganancias_matriz.append(ganancia_matriz_cop)
        pct_matriz_list.append(pct_matriz)
        potencializador_list.append(pct_potencializador)
        potencializador_cop_list.append(potencializador_cop)

    df['Ganancia_Matriz_Pct'] = pct_matriz_list
    df['Ganancia_Matriz_COP'] = ganancias_matriz
    df['Potencializador_Pct'] = potencializador_list
    df['Potencializador_COP'] = potencializador_cop_list
    df['Ganancia estimada'] = ganancias_total
    
    return df


def generar_analisis_como_vamos(df):
    """
    Realiza un análisis diagnóstico completo y automatizado sobre las métricas clave de
    'Cómo Vamos': Facturación, Saldos, Embudo de Disponibles y Metas de Crecimiento.
    """
    if df is None or df.empty:
        return {}

    df = normalizar_columnas(df.copy())
    
    # 1. Facturación
    obj_fact = float(pd.to_numeric(df['Objetivo Facturación'], errors='coerce').sum()) if 'Objetivo Facturación' in df.columns else 0.0
    real_fact = float(pd.to_numeric(df['Real Facturación'], errors='coerce').sum()) if 'Real Facturación' in df.columns else 0.0
    cump_fact = (real_fact / obj_fact * 100.0) if obj_fact > 0 else 0.0
    meta_110 = obj_fact * 1.10
    brecha_100 = obj_fact - real_fact
    brecha_110 = max(0.0, meta_110 - real_fact)

    # Brecha 100% detallada por líderes
    s_brecha = pd.to_numeric(df['Brecha Meta 100%'], errors='coerce') if 'Brecha Meta 100%' in df.columns else pd.Series(dtype=float)
    faltante_100_lideres = float(s_brecha[s_brecha > 0].sum()) if not s_brecha.empty else 0.0
    superavit_100_lideres = float(s_brecha[s_brecha < 0].abs().sum()) if not s_brecha.empty else 0.0
    
    # Avance % Facturación vs corte anterior
    s_avance_pct = pd.to_numeric(df['Avance % Facturación'], errors='coerce').dropna() if 'Avance % Facturación' in df.columns else pd.Series(dtype=float)
    avance_pct_prom = float(s_avance_pct.mean() * 100.0) if not s_avance_pct.empty else 0.0
    
    s_avance_cop = pd.to_numeric(df['Avance Facturación COP'], errors='coerce').dropna() if 'Avance Facturación COP' in df.columns else pd.Series(dtype=float)
    avance_cop_total = float(s_avance_cop.sum()) if not s_avance_cop.empty else 0.0

    real_activas = float(pd.to_numeric(df['Real Activas'], errors='coerce').sum()) if 'Real Activas' in df.columns else 0.0
    obj_activas = float(pd.to_numeric(df['Objetivo Activas'], errors='coerce').sum()) if 'Objetivo Activas' in df.columns else 0.0
    cump_activas = (real_activas / obj_activas * 100.0) if obj_activas > 0 else 0.0
    productividad_prom = (real_fact / real_activas) if real_activas > 0 else 0.0

    # 2. Saldos
    s_saldo = pd.to_numeric(df['Saldo'], errors='coerce').fillna(0) if 'Saldo' in df.columns else pd.Series([0])
    total_saldo = float(s_saldo.sum())
    lideres_con_saldo = int((s_saldo > 0).sum())
    max_saldo = float(s_saldo.max())
    saldo_prom = float(s_saldo.mean())

    # 3. Embudo de Disponibles
    total_disponibles = float(pd.to_numeric(df['Disponibles'], errors='coerce').sum()) if 'Disponibles' in df.columns else 0.0
    tasa_conversion = (real_activas / total_disponibles * 100.0) if total_disponibles > 0 else 0.0
    pendientes_pedido = max(0.0, total_disponibles - real_activas)
    
    inicios = float(pd.to_numeric(df['Inicios'], errors='coerce').sum()) if 'Inicios' in df.columns else 0.0
    reinicios = float(pd.to_numeric(df['Reinicios'], errors='coerce').sum()) if 'Reinicios' in df.columns else 0.0
    recuperos = float(pd.to_numeric(df['Recuperos'], errors='coerce').sum()) if 'Recuperos' in df.columns else 0.0
    
    inactivas = {}
    for i in range(1, 7):
        col_inact = f'Inactiva {i}'
        inactivas[f'Inactiva_{i}'] = float(pd.to_numeric(df[col_inact], errors='coerce').sum()) if col_inact in df.columns else 0.0
    total_inactivas = sum(inactivas.values())

    # 4. Metas de Crecimiento
    metas = {
        'meta_1plus': float(pd.to_numeric(df['Meta_Crecer_1plus_150k'], errors='coerce').sum()) if 'Meta_Crecer_1plus_150k' in df.columns else 0.0,
        'meta_3plus': float(pd.to_numeric(df['Meta_Crecer_3plus_200k'], errors='coerce').sum()) if 'Meta_Crecer_3plus_200k' in df.columns else 0.0,
        'meta_5plus': float(pd.to_numeric(df['Meta_Crecer_5plus_300k'], errors='coerce').sum()) if 'Meta_Crecer_5plus_300k' in df.columns else 0.0,
        'meta_7plus': float(pd.to_numeric(df['Meta_Crecer_7plus_500k'], errors='coerce').sum()) if 'Meta_Crecer_7plus_500k' in df.columns else 0.0,
        'meta_9plus': float(pd.to_numeric(df['Meta_Crecer_9plus_750k'], errors='coerce').sum()) if 'Meta_Crecer_9plus_750k' in df.columns else 0.0,
    }

    # Estructura del diagnóstico
    analisis = {
        'facturacion': {
            'objetivo': obj_fact,
            'real': real_fact,
            'cumplimiento_pct': cump_fact,
            'meta_110': meta_110,
            'brecha_100': brecha_100,
            'brecha_110': brecha_110,
            'productividad_promedio': productividad_prom,
            'avance_pct_promedio': avance_pct_prom,
            'avance_cop_total': avance_cop_total,
            'faltante_100_lideres': faltante_100_lideres,
            'superavit_100_lideres': superavit_100_lideres
        },
        'activas': {
            'objetivo': obj_activas,
            'real': real_activas,
            'cumplimiento_pct': cump_activas
        },
        'saldos': {
            'total_saldo': total_saldo,
            'lideres_afectados': lideres_con_saldo,
            'max_saldo': max_saldo,
            'saldo_promedio': saldo_prom
        },
        'disponibles': {
            'total_disponibles': total_disponibles,
            'tasa_conversion_pct': tasa_conversion,
            'pendientes_pedido': pendientes_pedido,
            'inicios': inicios,
            'reinicios': reinicios,
            'recuperos': recuperos,
            'inactivas_detalle': inactivas,
            'total_inactivas': total_inactivas,
            'retencion': {
                'inactiva_1': {
                    'total': inactivas.get('Inactiva_1', 0.0),
                    'max_fuga_pct': 12.0,
                    'min_retencion_pct': 88.0,
                    'max_fuga_cant': inactivas.get('Inactiva_1', 0.0) * 0.12,
                    'min_retencion_cant': inactivas.get('Inactiva_1', 0.0) * 0.88,
                },
                'inactiva_2': {
                    'total': inactivas.get('Inactiva_2', 0.0),
                    'max_fuga_pct': 8.0,
                    'min_retencion_pct': 92.0,
                    'max_fuga_cant': inactivas.get('Inactiva_2', 0.0) * 0.08,
                    'min_retencion_cant': inactivas.get('Inactiva_2', 0.0) * 0.92,
                },
                'inactiva_3': {
                    'total': inactivas.get('Inactiva_3', 0.0),
                    'max_fuga_pct': 6.0,
                    'min_retencion_pct': 94.0,
                    'max_fuga_cant': inactivas.get('Inactiva_3', 0.0) * 0.06,
                    'min_retencion_cant': inactivas.get('Inactiva_3', 0.0) * 0.94,
                }
            }
        },
        'metas': metas
    }

def rotar_y_guardar_nuevo_ciclo(nuevo_excel_origen, ruta_destino='Base para el como vamos.xlsx'):
    """
    Toma el archivo de datos del ciclo actual, convierte la hoja 'Base para el como vamos'
    en la hoja 'Como vamos anterior', coloca los nuevos datos en la hoja 'Base para el como vamos'
    y sobrescribe el archivo 'Base para el como vamos.xlsx'.
    """
    df_anterior = None
    if os.path.exists(ruta_destino):
        try:
            xl_existente = pd.ExcelFile(ruta_destino)
            hojas_lower = {s.lower().strip(): s for s in xl_existente.sheet_names}
            if 'base para el como vamos' in hojas_lower:
                df_anterior = pd.read_excel(ruta_destino, sheet_name=hojas_lower['base para el como vamos'])
        except Exception as e:
            print(f"Advertencia al leer ciclo anterior de '{ruta_destino}': {e}")

    # Leer el nuevo archivo entregado
    xl_nuevo = pd.ExcelFile(nuevo_excel_origen)
    hoja_nuevo = xl_nuevo.sheet_names[0]
    for s in xl_nuevo.sheet_names:
        if 'base para el como vamos' in s.lower():
            hoja_nuevo = s
            break
    df_nuevo = pd.read_excel(nuevo_excel_origen, sheet_name=hoja_nuevo)

    # Guardar ambas hojas en el archivo destino
    try:
        with pd.ExcelWriter(ruta_destino, engine='openpyxl') as writer:
            df_nuevo.to_excel(writer, sheet_name='Base para el como vamos', index=False)
            if df_anterior is not None and not df_anterior.empty:
                df_anterior.to_excel(writer, sheet_name='Como vamos anterior', index=False)
    except PermissionError:
        raise PermissionError(f"El archivo '{ruta_destino}' está abierto en Microsoft Excel. Por favor ciérralo en Excel y vuelve a presionar el botón.")
            
    safe_print(f"[OK] ¡Ciclo rotado exitosamente en '{ruta_destino}'!")
    return calcular_metas_ciclo(ruta_destino)

# --- FUNCIONES DE FORMATO CONDICIONAL Y EXPORTACIÓN A COLORES ---

def color_cumplimiento(val):
    """
    Retorna estilo CSS para formatear celdas de porcentaje según semáforo:
    <95%: Rojo (#FEE2E2), 95-99.99%: Amarillo (#FEF9C3), 100-109.99%: Verde (#DCFCE7), >=110%: Morado (#F3E8FF).
    """
    try:
        if pd.isna(val):
            return ""
        s = str(val)
        for char in ['%', '+', '➡️', '⬆️', '⬇️', '$', ' ']:
            s = s.replace(char, '')
        if not s:
            return ""
        num = float(s)
            
        if num < 95.0:
            return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
        elif num < 100.0:
            return 'background-color: #FEF9C3; color: #854D0E; font-weight: bold;'
        elif num < 110.0:
            return 'background-color: #DCFCE7; color: #166534; font-weight: bold;'
        else:
            return 'background-color: #F3E8FF; color: #6B21A8; font-weight: bold;'
    except:
        return ""

def color_avance(val):
    """
    Retorna estilo CSS para Avance %:
    > 0%: Verde (#DCFCE7), < 0%: Rojo (#FEE2E2), 0%: Neutro.
    """
    try:
        if pd.isna(val):
            return ""
        s = str(val)
        for char in ['%', '+', '➡️', '⬆️', '⬇️', '$', ' ']:
            s = s.replace(char, '')
        if not s:
            return ""
        num = float(s)
        if num > 0:
            return 'background-color: #DCFCE7; color: #166534; font-weight: bold;'
        elif num < 0:
            return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
        else:
            return ''
    except:
        return ""

def color_saldo(val):
    """
    Retorna estilo CSS para formatear celdas de saldo:
    0-2: Verde (#DCFCE7), 3-5: Amarillo (#FEF9C3), >=6: Rojo (#FEE2E2).
    """
    try:
        if pd.isna(val):
            return ""
        s = str(val).replace('$', '').replace(',', '').replace(' ', '').strip()
        num = float(s)
        if num <= 2:
            return 'background-color: #DCFCE7; color: #166534; font-weight: bold;'
        elif num <= 5:
            return 'background-color: #FEF9C3; color: #854D0E; font-weight: bold;'
        else:
            return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
    except:
        return ""

def exportar_excel_con_colores(df_dict, buffer_salida=None):
    """
    Genera un archivo Excel profesional (.xlsx) aplicando colores reales en celdas (openpyxl PatternFill).
    df_dict: dict de {sheet_name: dataframe}
    """
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Eliminar hoja inicial por defecto

    # Definir paleta de colores openpyxl
    fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    fill_green = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    font_green = Font(name="Calibri", size=10, bold=True, color="166534")
    
    fill_yellow = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    font_yellow = Font(name="Calibri", size=10, bold=True, color="854D0E")
    
    fill_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    font_red = Font(name="Calibri", size=10, bold=True, color="991B1B")
    
    fill_purple = PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid")
    font_purple = Font(name="Calibri", size=10, bold=True, color="6B21A8")

    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    for sheet_name, df in df_dict.items():
        ws = wb.create_sheet(title=str(sheet_name)[:31])
        
        headers = list(df.columns)
        ws.append(headers)
        
        for col_num, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        for _, row in df.iterrows():
            row_vals = [row[c] for c in headers]
            ws.append(row_vals)
            r_idx = ws.max_row
            
            for col_idx, col_name in enumerate(headers, 1):
                cell = ws.cell(row=r_idx, column=col_idx)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")
                
                val = row[col_name]
                col_lower = str(col_name).lower()
                
                if 'cumplimiento' in col_lower or 'avance' in col_lower:
                    try:
                        clean_str = str(val).replace('%', '').replace('+', '').strip()
                        num = float(clean_str)
                        
                        if num < 95.0:
                            cell.fill = fill_red
                            cell.font = font_red
                        elif num < 100.0:
                            cell.fill = fill_yellow
                            cell.font = font_yellow
                        elif num < 110.0:
                            cell.fill = fill_green
                            cell.font = font_green
                        else:
                            cell.fill = fill_purple
                            cell.font = font_purple
                    except (ValueError, TypeError):
                        pass
                elif 'saldo' in col_lower:
                    try:
                        num = float(val)
                        if num <= 2:
                            cell.fill = fill_green
                            cell.font = font_green
                        elif num <= 5:
                            cell.fill = fill_yellow
                            cell.font = font_yellow
                        else:
                            cell.fill = fill_red
                            cell.font = font_red
                    except (ValueError, TypeError):
                        pass

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    if buffer_salida is not None:
        wb.save(buffer_salida)
        return buffer_salida
    else:
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

# --- MÓDULO INFORME TABLEAU MANAGER ("INFORME TABLEAU CAM") ---

RUTA_COMENTARIOS = 'comentarios_lideres.json'

def cargar_comentarios_lideres():
    """
    Carga el diccionario de comentarios/notas por Código CB desde un JSON persistente.
    """
    if os.path.exists(RUTA_COMENTARIOS):
        try:
            with open(RUTA_COMENTARIOS, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def guardar_comentario_lider(codigo_cb, comentario):
    """
    Guarda o actualiza el comentario de una consultora por su Código CB en el JSON persistente.
    """
    comentarios = cargar_comentarios_lideres()
    codigo_str = str(codigo_cb).strip()
    comentarios[codigo_str] = str(comentario).strip()
    try:
        with open(RUTA_COMENTARIOS, 'w', encoding='utf-8') as f:
            json.dump(comentarios, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Error al guardar comentario: {e}")
        return False

def guardar_todos_comentarios(dict_comentarios):
    """
    Guarda masivamente un diccionario de comentarios {codigo_cb: comentario}.
    """
    comentarios = cargar_comentarios_lideres()
    for cb, nota in dict_comentarios.items():
        cb_str = str(cb).strip()
        nota_str = str(nota).strip()
        if nota_str:
            comentarios[cb_str] = nota_str
        elif cb_str in comentarios and nota_str == "":
            comentarios.pop(cb_str, None)
    try:
        with open(RUTA_COMENTARIOS, 'w', encoding='utf-8') as f:
            json.dump(comentarios, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Error guardando comentarios masivos: {e}")
        return False

def procesar_base_tableau_manager(origen='Base de Datos.xlsx'):
    """
    Carga y procesa la sábana de datos de Tableau (Base de Datos.xlsx).
    - Normaliza nombres de columnas y tipos numéricos.
    - Incluye todos los estados de actividad comercial (Activas, Inactivas 1 a 6).
    - Integra comentarios persistentes por Código CB.
    """
    df = None
    if isinstance(origen, str):
        if os.path.exists(origen):
            df = pd.read_excel(origen, sheet_name=0)
    elif isinstance(origen, pd.DataFrame):
        df = origen.copy()
    elif hasattr(origen, 'read'): # UploadedFile de Streamlit
        try:
            origen.seek(0)
        except Exception:
            pass
        df = pd.read_excel(origen, sheet_name=0)

    if df is None or df.empty:
        return None

    # Detectar si la primera fila son celdas vacías o 'Unnamed' y promover la fila de cabecera real
    if any('unnamed' in str(c).lower() for c in df.columns[:5]):
        for r_idx in range(min(5, len(df))):
            row_vals = [str(x) for x in df.iloc[r_idx].values]
            if any('codigo' in x.lower() or 'código' in x.lower() or 'asesora' in x.lower() for x in row_vals):
                df.columns = df.iloc[r_idx]
                df = df.iloc[r_idx + 1:].reset_index(drop=True)
                break

    # Normalización de caracteres en nombres de columnas
    cols_rename = {}
    for c in df.columns:
        clean_c = str(c).replace('\ufffd', 'ó').strip()
        cols_rename[c] = clean_c
    df = df.rename(columns=cols_rename)

    # Reemplazar posibles problemas de codificación comunes y estandarizar nombres de cabecera
    renombres_clave = {
        'Codigo CB': 'Codigo CB',
        'Código CB': 'Codigo CB',
        'Gerencia ': 'Gerencia',
        'Nombre Gerencia': 'Gerencia',
        'Sector ': 'Sector',
        'Nombre Setor': 'Sector',
        'Nombre Sector': 'Sector',
        'Situación': 'Situación',
        'Deuda Mora ': 'Deuda Mora',
        'Pts para Ascender ': 'Pts Asc',
        'Pts Para Mantener': 'Pts Mant',
        'Pts Acumulados': 'Pts Acum'
    }
    df = df.rename(columns=renombres_clave)

    # Columnas numéricas a limpiar
    cols_num = [
        'Credito Total', 'Credito Disponible', 'Deuda Total', 'Deuda Mora',
        'Pts Natura', 'Pts AVON', 'Pts Total VD', 'Pts VOL', 'Pts Acum', 'Pts Mant', 'Pts Asc',
        'Ped. Pendientes', 'Ped. Mora', 'Fact. Total', 'Grupo', 'Ciclo', 'Cod. Gerencia', 'Cod. Sector', 'Ciclos Inactividad'
    ]
    for col in cols_num:
        if col in df.columns:
            df[col] = limpiar_serie_numerica(df[col])

    # Convertir flotantes enteros y columnas como Ciclo/Grupo a int64 sin decimales (.000000)
    for col in df.columns:
        if pd.api.types.is_float_dtype(df[col]):
            if 'ciclo' in str(col).lower() or 'grupo' in str(col).lower():
                df[col] = df[col].apply(lambda v: int(round(v * 1000)) if (pd.notna(v) and 10 < v < 1000 and (v % 1 != 0)) else (int(round(v)) if pd.notna(v) else 0))
            else:
                non_nulls = df[col].dropna()
                if not non_nulls.empty and (non_nulls.round(4) % 1 == 0).all():
                    df[col] = df[col].fillna(0).round().astype('int64')

    # Unir comentarios persistentes por Código CB
    comentarios = cargar_comentarios_lideres()
    if 'Codigo CB' in df.columns:
        df['Codigo_CB_key'] = df['Codigo CB'].astype(str).str.strip()
        df['Comentarios_Lider'] = df['Codigo_CB_key'].map(lambda k: comentarios.get(k, ''))
    else:
        df['Comentarios_Lider'] = ''

    # Sincronización consistente de estado Activa entre Situación y Sit. Comercial
    if 'Sit. Comercial' in df.columns:
        mask_activa_cycle = pd.Series(False, index=df.index)
        if 'Pts Total VD' in df.columns:
            mask_activa_cycle = mask_activa_cycle | (df['Pts Total VD'] > 0)
        if 'Fact. Total' in df.columns:
            mask_activa_cycle = mask_activa_cycle | (df['Fact. Total'] > 0)
        if 'Situación' in df.columns:
            mask_activa_cycle = mask_activa_cycle | (df['Situación'].astype(str).str.strip().str.lower() == 'activa')
        
        df.loc[mask_activa_cycle, 'Sit. Comercial'] = 'Activa'
        if 'Situación' in df.columns:
            df.loc[mask_activa_cycle, 'Situación'] = 'Activa'

    return df

def color_nivel(val):
    """
    Retorna estilo CSS para formatear la celda de Nivel/Color:
    Bronce, Plata, Oro, Platino, Zafiro, Diamante.
    """
    if pd.isna(val):
        return ""
    v = str(val).strip().lower()
    if 'bronce' in v:
        return 'background-color: #FEF3C7; color: #92400E; font-weight: bold;'
    elif 'plata' in v:
        return 'background-color: #F1F5F9; color: #475569; font-weight: bold;'
    elif 'oro' in v:
        return 'background-color: #FEF9C3; color: #854D0E; font-weight: bold;'
    elif 'platino' in v:
        return 'background-color: #E0F2FE; color: #0369A1; font-weight: bold;'
    elif 'zafiro' in v:
        return 'background-color: #DBEAFE; color: #1E40AF; font-weight: bold;'
    elif 'diamante' in v:
        return 'background-color: #F3E8FF; color: #6B21A8; font-weight: bold;'
    return ""

def color_situacion(val):
    """
    Retorna estilo CSS para formatear la celda de Situación Comercial:
    - Activa: Verde suave (#E6F4EA / texto #137333).
    - Inactiva 1: Amarillo (#FEF9C3 / texto #854D0E).
    - Inactiva 2/3: Naranja (#FFEDD5 / texto #C2410C).
    - Inactiva 4/5: Rojo (#FEE2E2 / texto #991B1B).
    """
    if pd.isna(val):
        return ""
    v = str(val).strip().lower()
    if 'activa' in v and 'inactiva' not in v:
        return 'background-color: #E6F4EA; color: #137333; font-weight: bold;'
    elif 'inactiva 1' in v:
        return 'background-color: #FEF9C3; color: #854D0E; font-weight: bold;'
    elif 'inactiva 2' in v or 'inactiva 3' in v:
        return 'background-color: #FFEDD5; color: #C2410C; font-weight: bold;'
    elif 'inactiva 4' in v or 'inactiva 5' in v or 'inactiva 6' in v:
        return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
    return ""

def color_deuda_mora(val):
    """
    Retorna estilo CSS para formatear celdas de Deuda Mora según semáforo:
    - $0: Sin mora (neutro).
    - $1 a $200.000: Amarillo (#FEF9C3 / texto #854D0E).
    - $200.001 a $500.000: Naranja (#FFEDD5 / texto #C2410C).
    - > $500.000: Rojo (#FEE2E2 / texto #991B1B).
    """
    try:
        if pd.isna(val):
            return ""
        s = str(val)
        for char in ['$', '+', '➡️', '⬆️', '⬇️', ',', ' ', '.']:
            s = s.replace(char, '')
        if not s:
            return ""
        num = float(s)
        if num <= 0:
            return ""
        elif num <= 200000:
            return 'background-color: #FEF9C3; color: #854D0E; font-weight: bold;'
        elif num <= 500000:
            return 'background-color: #FFEDD5; color: #C2410C; font-weight: bold;'
        else:
            return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
    except Exception:
        return ""

def normalizar_estado_mi_grupo(val):
    """
    Normaliza el texto del campo 'ESTADO' independientemente de si viene en MAYÚSCULAS,
    minúsculas o combinación de ambos (ej: 'ACTIVA', 'Activa', 'INACTIVAS 1', 'Inactivas 1', etc.)
    """
    if pd.isna(val):
        return ""
    s = str(val).strip()
    s_lower = s.lower()
    
    # Inactiva 1 a 6 (ej. 'inactivas 1', 'INACTIVAS 1', 'inactiva 1', 'INACTIVA 1')
    for i in range(1, 7):
        if 'inactiv' in s_lower and str(i) in s_lower:
            return f"Inactiva {i}"
            
    if 'activa' in s_lower and 'inactiva' not in s_lower:
        return "Activa"
    elif 'cesada' in s_lower:
        return "Cesada"
    elif 'registrada' in s_lower:
        return "Registrada"
    elif 'intenci' in s_lower:
        return "Intención"
    return s

def actualizar_situacion_comercial_desde_mi_grupo(origen_mi_grupo='mi_grupo.xls', ruta_base='Base de Datos.xlsx'):
    """
    Lee la tabla 'mi_grupo.xls' (o .xlsx), extrae la columna 'ESTADO' / 'Sit. Comercial'
    (reconociendo combinaciones de MAYÚSCULAS y minúsculas), la normaliza y actualiza
    la columna 'Sit. Comercial' en 'Base de Datos.xlsx' vinculando por Código CB.
    """
    if not os.path.exists(ruta_base):
        return {'exito': False, 'error': f"No se encontró el archivo base '{ruta_base}'."}

    df_grupo = None
    if isinstance(origen_mi_grupo, str):
        if os.path.exists(origen_mi_grupo):
            try:
                df_grupo = pd.read_excel(origen_mi_grupo)
            except Exception:
                try:
                    df_grupo = pd.read_excel(origen_mi_grupo, engine='xlrd')
                except Exception:
                    pass
    elif isinstance(origen_mi_grupo, pd.DataFrame):
        df_grupo = origen_mi_grupo.copy()
    elif hasattr(origen_mi_grupo, 'read'):
        try:
            origen_mi_grupo.seek(0)
        except Exception:
            pass
        try:
            df_grupo = pd.read_excel(origen_mi_grupo)
        except Exception:
            try:
                origen_mi_grupo.seek(0)
                df_grupo = pd.read_excel(origen_mi_grupo, engine='xlrd')
            except Exception:
                pass

    if df_grupo is None or df_grupo.empty:
        return {'exito': False, 'error': "No se pudo leer el archivo 'mi_grupo'."}

    # 1. Identificar columna de Código CB en mi_grupo (resistente a mayúsculas/minúsculas)
    col_code_grupo = None
    for c in df_grupo.columns:
        c_up = str(c).strip().upper()
        if ('DIGO' in c_up or 'CODIGO' in c_up or 'CB' in c_up) and 'NOMBRE' not in c_up:
            col_code_grupo = c
            break

    if not col_code_grupo:
        return {'exito': False, 'error': "No se encontró la columna de Código en el archivo 'mi_grupo'."}

    # 2. Identificar columna de Estado / Situación en mi_grupo (resistente a mayúsculas/minúsculas/combinadas)
    col_estado_grupo = None
    # Coincidencia exacta con 'ESTADO' primero
    for c in df_grupo.columns:
        if str(c).strip().upper() == 'ESTADO':
            col_estado_grupo = c
            break
    # Coincidencia parcial si no hay exacta
    if not col_estado_grupo:
        for c in df_grupo.columns:
            c_up = str(c).strip().upper()
            if 'ESTADO' in c_up or 'SITUACION' in c_up or 'SITUACI' in c_up:
                col_estado_grupo = c
                break

    if not col_estado_grupo:
        return {'exito': False, 'error': "No se encontró la columna 'ESTADO' o 'Situación' en 'mi_grupo'."}

    # Crear mapeo {codigo_cb_clean: estado_normalizado}
    df_grupo['cb_clean'] = df_grupo[col_code_grupo].astype(str).str.strip()
    df_grupo['estado_clean'] = df_grupo[col_estado_grupo].apply(normalizar_estado_mi_grupo)

    mapa_estados = dict(zip(df_grupo['cb_clean'], df_grupo['estado_clean']))

    # Cargar la base principal
    df_base = pd.read_excel(ruta_base, sheet_name=0)
    col_code_base = None
    for c in df_base.columns:
        if 'Codigo CB' in str(c) or 'Código CB' in str(c) or 'Código de consultora' in str(c):
            col_code_base = c
            break
    if not col_code_base:
        col_code_base = df_base.columns[0]

    col_sit_comercial = next((c for c in df_base.columns if 'sit. comercial' in str(c).lower() or 'sit comercial' in str(c).lower()), None)
    col_situacion_macro = next((c for c in df_base.columns if str(c).strip().lower() in ['situación', 'situacion']), None)

    if not col_sit_comercial:
        col_sit_comercial = 'Sit. Comercial'
        df_base[col_sit_comercial] = ''

    df_base['cb_clean'] = df_base[col_code_base].astype(str).str.strip()

    coincidencias = 0
    cambios = 0
    detalles_cambios = []

    for idx in df_base.index:
        cb = df_base.at[idx, 'cb_clean']
        if cb in mapa_estados:
            coincidencias += 1
            nuevo_estado = mapa_estados[cb]
            estado_actual = str(df_base.at[idx, col_sit_comercial]).strip() if pd.notna(df_base.at[idx, col_sit_comercial]) else ''
            if nuevo_estado and nuevo_estado != estado_actual:
                nombre = str(df_base.at[idx, 'Nombre'] if 'Nombre' in df_base.columns else cb)
                detalles_cambios.append({
                    'Código CB': cb,
                    'Asesora / Consultora': nombre,
                    'Estado Anterior': estado_actual,
                    'Nuevo Estado (mi_grupo)': nuevo_estado
                })
                df_base.at[idx, col_sit_comercial] = nuevo_estado
                if col_situacion_macro:
                    if nuevo_estado == 'Activa':
                        df_base.at[idx, col_situacion_macro] = 'Activa'
                    elif 'inactiva' in nuevo_estado.lower():
                        if any(k in nuevo_estado.lower() for k in ['1', '2', '3']):
                            df_base.at[idx, col_situacion_macro] = 'Disponible'
                        elif any(k in nuevo_estado.lower() for k in ['4', '5', '6']):
                            df_base.at[idx, col_situacion_macro] = 'Indisponible'
                cambios += 1

    df_base = df_base.drop(columns=['cb_clean'], errors='ignore')

    # Guardar en Base de Datos.xlsx y sincronizar SQLite
    try:
        df_base.to_excel(ruta_base, index=False)
        try:
            sincronizar_excel_tableau_a_sqlite(ruta_base)
        except Exception:
            pass
        return {
            'exito': True,
            'coincidencias': coincidencias,
            'cambios': cambios,
            'detalles': detalles_cambios
        }
    except Exception as e:
        return {'exito': False, 'error': f"Error al guardar '{ruta_base}': {e}"}

def actualizar_base_desde_activas(origen_activas, ruta_base='Base de Datos.xlsx'):
    """
    Cruce inteligente de base de datos con archivo de 'activas' (.xlsx, .xls, .csv).
    - Lee el archivo de activas identificando inteligentemente la columna de Código CB o Documento.
    - Si el archivo trae columna de Estado/Situación, la normaliza; de lo contrario, marca a las asesoras cruzadas como 'Activa'.
    - Si el archivo contiene Facturación, Puntos, Pedidos o Mora, complementa/actualiza las columnas correspondientes.
    - Guarda en 'Base de Datos.xlsx' y sincroniza automáticamente con la base de datos relacional SQLite (base_matices.db).
    """
    if not os.path.exists(ruta_base):
        return {'exito': False, 'error': f"No se encontró el archivo base '{ruta_base}'."}

    df_act = None
    # 1. Cargar archivo multi-formato
    try:
        if isinstance(origen_activas, str):
            if not os.path.exists(origen_activas):
                return {'exito': False, 'error': f"No existe el archivo '{origen_activas}'."}
            if origen_activas.lower().endswith('.csv'):
                try:
                    df_act = pd.read_csv(origen_activas, encoding='utf-8')
                except Exception:
                    df_act = pd.read_csv(origen_activas, encoding='latin1', sep=None, engine='python')
            else:
                try:
                    df_act = pd.read_excel(origen_activas)
                except Exception:
                    df_act = pd.read_excel(origen_activas, engine='xlrd')
        elif isinstance(origen_activas, pd.DataFrame):
            df_act = origen_activas.copy()
        elif hasattr(origen_activas, 'read'):
            name = getattr(origen_activas, 'name', '').lower()
            try:
                origen_activas.seek(0)
            except Exception:
                pass
            if name.endswith('.csv'):
                try:
                    df_act = pd.read_csv(origen_activas, encoding='utf-8')
                except Exception:
                    try:
                        origen_activas.seek(0)
                        df_act = pd.read_csv(origen_activas, encoding='latin1', sep=None, engine='python')
                    except Exception:
                        pass
            else:
                try:
                    df_act = pd.read_excel(origen_activas)
                except Exception:
                    try:
                        origen_activas.seek(0)
                        df_act = pd.read_excel(origen_activas, engine='xlrd')
                    except Exception:
                        pass
    except Exception as e_load:
        return {'exito': False, 'error': f"Error al leer el archivo de activas: {e_load}"}

    if df_act is None or df_act.empty:
        return {'exito': False, 'error': "El archivo de activas está vacío o no se pudo interpretar."}

    # Si la primera fila es encabezado desplazado, buscar la fila con 'CODIGO' o 'DOCUMENTO' o 'NOMBRE'
    if not any(any(k in str(c).upper() for k in ['CODIGO', 'CB', 'DOC', 'ASESORA', 'CONSULTORA', 'NOMBRE']) for c in df_act.columns):
        for r_idx in range(min(10, len(df_act))):
            row_vals = [str(x).upper() for x in df_act.iloc[r_idx].values if pd.notna(x)]
            if any('COD' in v or 'DOC' in v or 'CONSULTORA' in v for v in row_vals):
                df_act.columns = df_act.iloc[r_idx]
                df_act = df_act.iloc[r_idx + 1:].reset_index(drop=True)
                break

    # 2. Identificar columna de Código CB en el archivo de activas
    col_code_act = None
    for c in df_act.columns:
        c_up = str(c).strip().upper()
        if ('DIGO' in c_up or 'CODIGO' in c_up or 'CB' in c_up) and 'NOMBRE' not in c_up and 'GERENCIA' not in c_up and 'SECTOR' not in c_up and 'GRUPO' not in c_up:
            col_code_act = c
            break
    if not col_code_act:
        for c in df_act.columns:
            c_up = str(c).strip().upper()
            if 'DOC' in c_up or 'CEDULA' in c_up or 'IDENT' in c_up:
                col_code_act = c
                break

    if not col_code_act:
        return {'exito': False, 'error': "No se identificó la columna de Código o Documento en el archivo de activas."}

    # 3. Detectar columnas complementarias opcionales en el archivo de activas
    col_estado_act = None
    for c in df_act.columns:
        c_up = str(c).strip().upper()
        if str(c).strip().upper() == 'ESTADO' or 'SIT. COM' in c_up or 'SITUACION' in c_up or 'SITUACI' in c_up:
            col_estado_act = c
            break

    col_fact_act = None
    for c in df_act.columns:
        c_up = str(c).strip().upper()
        if ('FACT' in c_up or 'VENTA' in c_up or 'VALOR PEDIDO' in c_up or 'VALOR' in c_up) and 'ANTERIOR' not in c_up and 'OBJETIVO' not in c_up:
            col_fact_act = c
            break

    col_pts_act = None
    for c in df_act.columns:
        c_up = str(c).strip().upper()
        if ('PTS' in c_up or 'PUNTOS' in c_up) and 'ANTERIOR' not in c_up:
            col_pts_act = c
            break

    col_ped_act = None
    for c in df_act.columns:
        c_up = str(c).strip().upper()
        if 'PEDIDO' in c_up or 'PED.' in c_up or 'ORDEN' in c_up:
            col_ped_act = c
            break

    # Función helper para limpiar códigos eliminando .0 y espacios
    def _limpiar_cb_str(v):
        if pd.isna(v):
            return ""
        s = str(v).strip()
        if s.endswith('.0'):
            s = s[:-2]
        return s

    df_act['cb_clean'] = df_act[col_code_act].apply(_limpiar_cb_str)
    df_act = df_act[df_act['cb_clean'] != '']

    # Diccionarios de mapeo desde el archivo de activas
    mapa_estados = {}
    mapa_fact = {}
    mapa_pts = {}
    mapa_ped = {}

    for _, row in df_act.iterrows():
        cb = row['cb_clean']
        if not cb:
            continue
        # Estado: Si trae columna la normalizamos; de lo contrario 'Activa'
        if col_estado_act and pd.notna(row[col_estado_act]):
            mapa_estados[cb] = normalizar_estado_mi_grupo(row[col_estado_act])
        else:
            mapa_estados[cb] = "Activa"

        if col_fact_act and pd.notna(row[col_fact_act]):
            try:
                mapa_fact[cb] = float(limpiar_numero(row[col_fact_act], 0.0))
            except Exception:
                pass

        if col_pts_act and pd.notna(row[col_pts_act]):
            try:
                mapa_pts[cb] = int(round(limpiar_numero(row[col_pts_act], 0)))
            except Exception:
                pass

        if col_ped_act and pd.notna(row[col_ped_act]):
            try:
                mapa_ped[cb] = int(round(limpiar_numero(row[col_ped_act], 0)))
            except Exception:
                pass

    # 4. Cargar Base de Datos.xlsx principal
    df_base = pd.read_excel(ruta_base, sheet_name=0)
    col_code_base = None
    for c in df_base.columns:
        if 'Codigo CB' in str(c) or 'Código CB' in str(c) or 'Código de consultora' in str(c):
            col_code_base = c
            break
    if not col_code_base:
        col_code_base = df_base.columns[0]

    col_sit_comercial = next((c for c in df_base.columns if 'sit. comercial' in str(c).lower() or 'sit comercial' in str(c).lower()), None)
    col_situacion_macro = next((c for c in df_base.columns if str(c).strip().lower() in ['situación', 'situacion']), None)

    if not col_sit_comercial:
        col_sit_comercial = 'Sit. Comercial'
        df_base[col_sit_comercial] = ''

    df_base['cb_clean'] = df_base[col_code_base].apply(_limpiar_cb_str)

    # Identificar columnas opcionales en Base de Datos.xlsx
    col_fact_base = next((c for c in df_base.columns if 'Fact. Total' in str(c) or 'Facturación Total' in str(c)), None)
    col_pts_base = next((c for c in df_base.columns if 'Pts Acumulados' in str(c) or 'Pts Acum' in str(c) or 'Pts Total VD' in str(c)), None)
    col_ped_base = next((c for c in df_base.columns if 'Ped. Pendientes' in str(c) or 'Pedidos' in str(c)), None)

    coincidencias = 0
    cambios_estado = 0
    cambios_totales = 0
    detalles_cambios = []
    cbs_base_set = set(df_base['cb_clean'].dropna())

    for idx in df_base.index:
        cb = df_base.at[idx, 'cb_clean']
        if cb in mapa_estados:
            coincidencias += 1
            nuevo_est = mapa_estados[cb]
            est_actual = str(df_base.at[idx, col_sit_comercial]).strip() if pd.notna(df_base.at[idx, col_sit_comercial]) else ''
            
            campos_modificados = []
            
            if nuevo_est:
                df_base.at[idx, col_sit_comercial] = nuevo_est
                if col_situacion_macro:
                    if nuevo_est == 'Activa':
                        df_base.at[idx, col_situacion_macro] = 'Activa'
                    elif 'inactiva' in nuevo_est.lower():
                        if any(k in nuevo_est.lower() for k in ['1', '2', '3']):
                            df_base.at[idx, col_situacion_macro] = 'Disponible'
                        elif any(k in nuevo_est.lower() for k in ['4', '5', '6']):
                            df_base.at[idx, col_situacion_macro] = 'Indisponible'
                
                if nuevo_est != est_actual:
                    cambios_estado += 1
                    campos_modificados.append(f"Estado: {est_actual or 'N/D'} ➔ {nuevo_est}")

            # Actualizar Facturación si viene en activas
            if cb in mapa_fact and col_fact_base:
                val_f = mapa_fact[cb]
                if val_f > 0:
                    df_base.at[idx, col_fact_base] = val_f
                    campos_modificados.append(f"Facturación: ${val_f:,.0f}".replace(",", "."))

            # Actualizar Puntos si viene en activas
            if cb in mapa_pts and col_pts_base:
                val_p = mapa_pts[cb]
                if val_p > 0:
                    df_base.at[idx, col_pts_base] = val_p
                    campos_modificados.append(f"Puntos: {val_p:,}".replace(",", "."))

            # Actualizar Pedidos si viene en activas
            if cb in mapa_ped and col_ped_base:
                val_ped = mapa_ped[cb]
                df_base.at[idx, col_ped_base] = val_ped
                campos_modificados.append(f"Pedidos: {val_ped}")

            if campos_modificados:
                cambios_totales += 1
                nombre = str(df_base.at[idx, 'Nombre'] if 'Nombre' in df_base.columns else cb)
                detalles_cambios.append({
                    'Código CB': cb,
                    'Asesora / Consultora': nombre,
                    'Estado Anterior': est_actual,
                    'Nuevo Estado': nuevo_est,
                    'Campos Actualizados': ", ".join(campos_modificados)
                })

    df_base = df_base.drop(columns=['cb_clean'], errors='ignore')

    # Guardar en Base de Datos.xlsx y refrescar SQLite
    try:
        df_base.to_excel(ruta_base, index=False)
        try:
            sincronizar_excel_tableau_a_sqlite(ruta_base)
        except Exception as e_sql:
            safe_print(f"Advertencia al sincronizar SQLite tras cruce de activas: {e_sql}")

        no_encontradas = [cb for cb in mapa_estados.keys() if cb not in cbs_base_set]

        return {
            'exito': True,
            'total_activas_archivo': len(mapa_estados),
            'coincidencias': coincidencias,
            'cambios_estado': cambios_estado,
            'cambios_totales': cambios_totales,
            'detalles': detalles_cambios,
            'no_encontradas_count': len(no_encontradas)
        }
    except Exception as e_save:
        return {'exito': False, 'error': f"Error al guardar '{ruta_base}': {e_save}"}

# --- MÓDULO DE AUTENTICACIÓN Y GESTIÓN DE USUARIOS POR ROL ---
RUTA_USUARIOS = 'usuarios.json'

def hashlib_sha256(texto):
    import hashlib
    return hashlib.sha256(str(texto).encode('utf-8')).hexdigest()

def cargar_usuarios():
    """
    Carga el diccionario de usuarios desde usuarios.json.
    Si no existe, crea usuarios predeterminados:
    - gerente (pass: admin123, rol: gerente)
    - lider8425 (pass: lider123, rol: lider, codigo_grupo: 8425)
    - lider7841 (pass: lider123, rol: lider, codigo_grupo: 7841)
    - asesor (pass: asesor123, rol: asesor)
    """
    if os.path.exists(RUTA_USUARIOS):
        try:
            with open(RUTA_USUARIOS, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass

    # Usuarios predeterminados si el archivo no existe
    def_pass_super = hashlib_sha256("superadmin123")
    def_pass_admin = hashlib_sha256("admin123")
    def_pass_lider = hashlib_sha256("lider123")
    def_pass_asesor = hashlib_sha256("asesor123")

    usuarios_default = {
        "admin": {
            "nombre": "Super Administrador del Sistema",
            "password_hash": def_pass_super,
            "rol": "superadmin",
            "codigo_grupo": None
        },
        "gerente": {
            "nombre": "Gerencia General",
            "password_hash": def_pass_admin,
            "rol": "gerente",
            "codigo_grupo": None
        },
        "lider8425": {
            "nombre": "Luz Dary Chacon Gaitan",
            "password_hash": def_pass_lider,
            "rol": "lider",
            "codigo_grupo": "8425"
        },
        "lider7841": {
            "nombre": "Carmenza Roncancio Gachancipa",
            "password_hash": def_pass_lider,
            "rol": "lider",
            "codigo_grupo": "7841"
        },
        "asesor": {
            "nombre": "Usuario Consulta Facturación",
            "password_hash": def_pass_asesor,
            "rol": "asesor",
            "codigo_grupo": None
        }
    }
    guardar_usuarios(usuarios_default)
    return usuarios_default

def guardar_usuarios(dict_usuarios):
    """
    Guarda el diccionario de usuarios en usuarios.json.
    """
    try:
        with open(RUTA_USUARIOS, 'w', encoding='utf-8') as f:
            json.dump(dict_usuarios, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Error al guardar usuarios: {e}")
        return False

def autenticar_usuario(username, password):
    """
    Valida credenciales. Retorna el diccionario del usuario si es correcto o None.
    """
    u_clean = str(username).strip().lower()
    usuarios = cargar_usuarios()
    if u_clean in usuarios:
        user_info = usuarios[u_clean]
        p_hash = hashlib_sha256(password)
        if user_info.get("password_hash") == p_hash:
            user_copy = user_info.copy()
            user_copy["username"] = u_clean
            return user_copy
    return None

def cambiar_password_usuario(username, nueva_password):
    """
    Actualiza la contraseña de un usuario y remueve la bandera de cambio obligatorio.
    """
    u_clean = str(username).strip().lower()
    usuarios = cargar_usuarios()
    if u_clean in usuarios:
        usuarios[u_clean]["password_hash"] = hashlib_sha256(nueva_password)
        usuarios[u_clean]["debe_cambiar_password"] = False
        if guardar_usuarios(usuarios):
            sincronizar_usuarios_a_sqlite()
            return True, "¡Contraseña actualizada exitosamente! Ya puedes navegar por el sistema."
    return False, "Error al guardar la nueva contraseña."

def registrar_o_actualizar_usuario(username, nombre, password, rol, codigo_grupo=None, codigo_sector=None, debe_cambiar_password=None):
    """
    Permite al Gerente o Superadmin crear o actualizar un usuario.
    Si se proporciona una contraseña (usuario nuevo o reseteo), debe_cambiar_password se establece en True por defecto.
    """
    u_clean = str(username).strip().lower()
    if not u_clean:
        return False, "El nombre de usuario no puede estar vacío."
    
    usuarios = cargar_usuarios()
    usuario_previo = usuarios.get(u_clean, {})
    
    if password:
        p_hash = hashlib_sha256(password)
        req_cambio = True if debe_cambiar_password is None else debe_cambiar_password
    else:
        p_hash = usuario_previo.get("password_hash", hashlib_sha256("123456"))
        req_cambio = usuario_previo.get("debe_cambiar_password", False) if debe_cambiar_password is None else debe_cambiar_password

    usr_data = {
        "nombre": nombre,
        "password_hash": p_hash,
        "rol": rol,
        "codigo_grupo": str(codigo_grupo).strip() if codigo_grupo else None,
        "debe_cambiar_password": req_cambio
    }
    if codigo_sector:
        usr_data["codigo_sector"] = str(codigo_sector).strip()
    elif u_clean in usuarios and "codigo_sector" in usuarios[u_clean]:
        usr_data["codigo_sector"] = usuarios[u_clean]["codigo_sector"]

    usuarios[u_clean] = usr_data
    if guardar_usuarios(usuarios):
        sincronizar_usuarios_a_sqlite()
        return True, f"Usuario '{u_clean}' guardado exitosamente."
    return False, "Error al guardar el archivo de usuarios."

def validar_sector_archivo(origen_file, sector_esperado):
    """
    Inspecciona un archivo subido (o ruta) en memoria y verifica que el Cod. Sector coincida con el sector del usuario en sesión.
    Retorna (valido: bool, sector_encontrado: str, nombre_sector: str, mensaje: str).
    """
    if not sector_esperado:
        return False, None, None, "⚠️ Tu perfil de Gerente no tiene un código de sector configurado en el sistema. Por favor solicita la asignación de tu código de sector."
    
    sector_esp_str = str(sector_esperado).strip()
    df_check = None
    
    try:
        if isinstance(origen_file, str):
            df_check = pd.read_excel(origen_file, nrows=10)
        elif hasattr(origen_file, 'read'):
            try:
                origen_file.seek(0)
            except Exception:
                pass
            df_check = pd.read_excel(origen_file, nrows=10)
            try:
                origen_file.seek(0)
            except Exception:
                pass
        elif isinstance(origen_file, pd.DataFrame):
            df_check = origen_file.head(10)
    except Exception as e:
        return False, None, None, f"Error al leer la estructura del archivo para validación: {e}"
        
    if df_check is None or df_check.empty:
        return False, None, None, "El archivo está vacío o no tiene datos válidos."

    # Detectar si la primera fila contiene 'Unnamed' y promover la cabecera real
    if any('unnamed' in str(c).lower() for c in df_check.columns[:5]):
        for r_idx in range(min(5, len(df_check))):
            row_vals = [str(x) for x in df_check.iloc[r_idx].values]
            if any('sector' in x.lower() or 'setor' in x.lower() for x in row_vals):
                df_check.columns = df_check.iloc[r_idx]
                df_check = df_check.iloc[r_idx + 1:].reset_index(drop=True)
                break

    # Buscar columna de sector
    col_sec = None
    col_nom_sec = None
    for c in df_check.columns:
        c_clean = str(c).replace('\ufffd', 'ó').strip()
        c_low = c_clean.lower()
        if ('cod' in c_low or 'cd' in c_low) and ('sector' in c_low or 'setor' in c_low):
            col_sec = c
        elif 'sector' in c_low or 'setor' in c_low:
            if not col_sec:
                col_sec = c
            else:
                col_nom_sec = c
                
    if not col_sec:
        return True, None, None, "No se encontró columna explícita de sector para validar."

    # Extraer el primer código de sector no nulo
    s_val = df_check[col_sec].dropna().astype(str).str.strip()
    if s_val.empty:
        return True, None, None, "No se encontraron valores numéricos de sector en las primeras filas."

    val_encontrado = s_val.iloc[0].split('.')[0]
    nom_encontrado = ""
    if col_nom_sec and not df_check[col_nom_sec].dropna().empty:
        nom_encontrado = str(df_check[col_nom_sec].dropna().iloc[0]).strip()

    if val_encontrado != sector_esp_str:
        msg = (
            f"❌ **Acceso Denegado - Validación de Sector**\n\n"
            f"El archivo subido pertenece a una Gerencia de otro Sector.\n\n"
            f"Se canceló la carga para proteger y evitar sobreescribir la información de tu sector.\n\n"
            f"📲 **¿Necesitas ayuda?** Por favor comunícate con el servicio de Soporte por WhatsApp al **3057939537**."
        )
        return False, val_encontrado, nom_encontrado, msg

    return True, val_encontrado, nom_encontrado, "Validación de sector exitosa."

def generar_password_aleatoria(longitud=8):
    import random, string
    chars = string.ascii_letters + string.digits
    return 'Lider' + ''.join(random.choice(chars) for _ in range(longitud))

def auto_crear_usuarios_lideres_desde_bases(ruta_tableau='Base de Datos.xlsx', ruta_como_vamos='Base para el como vamos.xlsx'):
    """
    Detecta automáticamente los grupos de líderes y crea o actualiza sus cuentas de usuario en el sistema.
    """
    creados = []
    if not os.path.exists(ruta_como_vamos):
        return creados

    try:
        df_metas = pd.read_excel(ruta_como_vamos, sheet_name=0)
        df_metas = normalizar_columnas(df_metas)
    except Exception:
        return creados

    col_grp = 'Código de grupo' if 'Código de grupo' in df_metas.columns else None
    col_nom = 'Nombre de consultora' if 'Nombre de consultora' in df_metas.columns else None

    if not col_grp:
        return creados

    df_tab = None
    if os.path.exists(ruta_tableau):
        try:
            df_tab = pd.read_excel(ruta_tableau, sheet_name=0)
        except Exception:
            pass

    grupos_procesados = set()
    usuarios_existentes = cargar_usuarios()

    for _, row in df_metas.iterrows():
        g_raw = row.get(col_grp)
        if pd.isna(g_raw):
            continue
        g_str = str(int(limpiar_numero(g_raw))) if limpiar_numero(g_raw) > 0 else str(g_raw).strip()
        if not g_str or g_str in grupos_procesados or g_str.lower() in ['nan', 'none', '0']:
            continue

        nom_lider = str(row.get(col_nom, '')).strip()
        if not nom_lider or nom_lider.lower() in ['nan', 'none', 'null', '0', '']:
            continue

        grupos_procesados.add(g_str)

        correo_lider = None
        if df_tab is not None:
            mask_g = (df_tab['Grupo'].astype(str).str.strip() == g_str) if 'Grupo' in df_tab.columns else pd.Series(False, index=df_tab.index)
            if 'Correo' in df_tab.columns and mask_g.any():
                df_g = df_tab[mask_g].dropna(subset=['Correo'])
                if not df_g.empty:
                    correo_lider = str(df_g['Correo'].iloc[0]).strip().lower()

        username = correo_lider if (correo_lider and '@' in correo_lider) else f"lider{g_str}"
        ya_existe = (username in usuarios_existentes)

        if ya_existe:
            # Líder existente: actualizar datos silenciosamente manteniendo clave activa
            registrar_o_actualizar_usuario(
                username=username,
                nombre=nom_lider,
                password=None,
                rol="lider",
                codigo_grupo=g_str
            )
        else:
            # Líder nueva: crear cuenta y generar clave temporal
            pass_gen = generar_password_aleatoria()
            exito, msg = registrar_o_actualizar_usuario(
                username=username,
                nombre=nom_lider,
                password=pass_gen,
                rol="lider",
                codigo_grupo=g_str
            )
            if exito:
                creados.append({
                    "Código Grupo": g_str,
                    "Nombre Líder": nom_lider,
                    "Usuario (Login / Correo)": username,
                    "Contraseña Generada": pass_gen,
                    "Resultado": "✨ Nueva Cuenta Creada"
                })

    return creados

# --- CONFIGURACIÓN DE PERMISOS GLOBALES DE CARGA ---
RUTA_CONFIG = 'configuracion.json'

DEFAULT_PERMISOS_PESTANAS = {
    "tab_tableau": {"nombre": "📊 Informe Tableau Cam", "gerente": True, "lider": True, "asesor": True},
    "tab_resumen": {"nombre": "📊 Resumen & KPIs", "gerente": True, "lider": True, "asesor": True},
    "tab_ganancia": {"nombre": "💵 Simulador de Ganancia", "gerente": True, "lider": True, "asesor": True},
    "tab_diagnostico": {"nombre": "🔎 Diagnóstico 'Cómo Vamos'", "gerente": True, "lider": True, "asesor": True},
    "tab_metas": {"nombre": "🎯 Metas de Crecimiento", "gerente": True, "lider": True, "asesor": True},
    "tab_detalle": {"nombre": "👥 Detalle Completo", "gerente": True, "lider": True, "asesor": True},
    "tab_exportar": {"nombre": "📤 Exportar Datos", "gerente": True, "lider": True, "asesor": True}
}

def cargar_configuracion():
    """
    Carga la configuración global de la aplicación.
    Por defecto, incluye los permisos de visibilidad por pestaña y la subida de archivos.
    """
    config = {
        "permitir_carga_lideres": False,
        "permisos_pestanas": DEFAULT_PERMISOS_PESTANAS.copy()
    }
    
    if os.path.exists(RUTA_CONFIG):
        try:
            with open(RUTA_CONFIG, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
                if isinstance(loaded, dict):
                    if "permitir_carga_lideres" in loaded:
                        config["permitir_carga_lideres"] = loaded["permitir_carga_lideres"]
                    if "permisos_pestanas" in loaded and isinstance(loaded["permisos_pestanas"], dict):
                        for tab_key, tab_val in DEFAULT_PERMISOS_PESTANAS.items():
                            if tab_key in loaded["permisos_pestanas"]:
                                # Combinar claves existentes preservando estructura
                                for r_key in ["gerente", "lider", "asesor"]:
                                    if r_key in loaded["permisos_pestanas"][tab_key]:
                                        config["permisos_pestanas"][tab_key][r_key] = bool(loaded["permisos_pestanas"][tab_key][r_key])
        except Exception as e:
            print(f"Nota al cargar configuración: {e}")
            
    guardar_configuracion(config)
    return config

def guardar_configuracion(dict_config):
    """
    Guarda la configuración global en configuracion.json.
    """
    try:
        with open(RUTA_CONFIG, 'w', encoding='utf-8') as f:
            json.dump(dict_config, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"Error al guardar configuración: {e}")
        return False

# --- MOTOR DE BASE DE DATOS RELACIONAL SQLITE (base_matices.db) ---
import sqlite3

RUTA_DB_SQLITE = 'base_matices.db'

def obtener_conexion_db():
    conn = sqlite3.connect(RUTA_DB_SQLITE, timeout=30.0)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
    except Exception:
        pass
    return conn

def inicializar_db_sqlite():
    """
    Inicializa la base de datos relacional SQLite 'base_matices.db' y crea las tablas indexadas.
    Sincroniza automáticamente los datos de Excel y JSON si la base se crea por primera vez.
    """
    conn = obtener_conexion_db()
    cursor = conn.cursor()

    # 1. Tabla de Usuarios
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        username TEXT PRIMARY KEY,
        nombre TEXT,
        password_hash TEXT,
        rol TEXT,
        codigo_grupo TEXT,
        codigo_sector TEXT
    )
    """)
    
    try:
        cursor.execute("ALTER TABLE usuarios ADD COLUMN codigo_sector TEXT")
    except Exception:
        pass

    # 2. Tabla de Configuración
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS configuracion (
        clave TEXT PRIMARY KEY,
        valor TEXT
    )
    """)

    # 3. Tabla de Consultoras Tableau
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS consultoras_tableau (
        codigo_cb TEXT PRIMARY KEY,
        nombre TEXT,
        documento_gpp TEXT,
        cod_gerencia TEXT,
        gerencia TEXT,
        cod_sector TEXT,
        sector TEXT,
        grupo TEXT,
        ciclo INTEGER,
        color TEXT,
        ascenso_cb TEXT,
        situacion TEXT,
        sit_comercial TEXT,
        ciclos_inactividad INTEGER,
        actividad_convergencia TEXT,
        inicios_completos TEXT,
        fact_vol REAL,
        fact_natura REAL,
        fact_avon REAL,
        fact_ce REAL,
        facturacion_total REAL,
        pts_natura INTEGER,
        pts_avon INTEGER,
        pts_total_vd INTEGER,
        pts_vol INTEGER,
        pts_acum INTEGER,
        pts_mant INTEGER,
        pts_asc INTEGER,
        tienda_online TEXT,
        tienda_sf TEXT,
        ciclo_primer_pedido INTEGER,
        fecha_nacimiento TEXT,
        mes_cumpleanos TEXT,
        deuda_total REAL,
        deuda_mora REAL,
        credito_total REAL,
        credito_disponible REAL,
        pedidos_pendientes INTEGER,
        pedidos_mora INTEGER,
        celular TEXT,
        correo TEXT,
        dpto_residencia TEXT,
        ciudad_residencia TEXT,
        barrio_residencia TEXT,
        direccion_residencia TEXT,
        complemento_residencia TEXT,
        referencia_residencia TEXT,
        dpto_entrega TEXT,
        ciudad_entrega TEXT,
        barrio_entrega TEXT,
        direccion_entrega TEXT,
        complemento_entrega TEXT,
        referencia_entrega TEXT,
        tiempo_casa INTEGER,
        origen_cb TEXT,
        notas_lider TEXT
    )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_tableau_grupo ON consultoras_tableau (grupo)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_tableau_sector ON consultoras_tableau (sector)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_tableau_cod_sector ON consultoras_tableau (cod_sector)")

    # 4. Tabla de Metas "Cómo Vamos"
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS metas_como_vamos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo_cb TEXT,
        nombre_consultora TEXT,
        nombre_gerencia TEXT,
        nombre_sector TEXT,
        codigo_grupo TEXT,
        color TEXT,
        obj_facturacion REAL,
        real_facturacion REAL,
        cump_facturacion REAL,
        obj_activas REAL,
        real_activas REAL,
        cump_activas REAL,
        saldo REAL,
        disponibles INTEGER,
        inicios INTEGER,
        reinicios INTEGER,
        recuperos INTEGER,
        ganancia_estimada REAL
    )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_metas_grupo ON metas_como_vamos (codigo_grupo)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_metas_sector ON metas_como_vamos (nombre_sector)")

    conn.commit()

    # Sincronización inicial
    sincronizar_usuarios_a_sqlite(conn)
    sincronizar_configuracion_a_sqlite(conn)

    conn.close()

def sincronizar_usuarios_a_sqlite(conn=None):
    close_at_end = False
    if conn is None:
        conn = obtener_conexion_db()
        close_at_end = True
    
    usuarios = cargar_usuarios()
    cursor = conn.cursor()
    
    try:
        cursor.execute("ALTER TABLE usuarios ADD COLUMN codigo_sector TEXT")
    except Exception:
        pass

    for uname, uinfo in usuarios.items():
        cursor.execute("""
        INSERT OR REPLACE INTO usuarios (username, nombre, password_hash, rol, codigo_grupo, codigo_sector)
        VALUES (?, ?, ?, ?, ?, ?)
        """, (uname, uinfo.get("nombre"), uinfo.get("password_hash"), uinfo.get("rol"), uinfo.get("codigo_grupo"), uinfo.get("codigo_sector")))
    conn.commit()
    if close_at_end:
        conn.close()

def sincronizar_configuracion_a_sqlite(conn=None):
    close_at_end = False
    if conn is None:
        conn = obtener_conexion_db()
        close_at_end = True
    
    cfg = cargar_configuracion()
    cursor = conn.cursor()
    for k, v in cfg.items():
        cursor.execute("""
        INSERT OR REPLACE INTO configuracion (clave, valor)
        VALUES (?, ?)
        """, (str(k), json.dumps(v)))
    conn.commit()
    if close_at_end:
        conn.close()

def sincronizar_excel_tableau_a_sqlite(ruta_excel='Base de Datos.xlsx', conn=None):
    """
    Convierte y vuelca el archivo Excel de Tableau hacia la tabla consultoras_tableau en SQLite.
    """
    df = procesar_base_tableau_manager(ruta_excel)
    if df is None or df.empty:
        return False
    
    close_at_end = False
    if conn is None:
        conn = obtener_conexion_db()
        close_at_end = True
    
    cursor = conn.cursor()
    
    # Extraer el código de sector del dataframe subido para borrar ÚNICAMENTE los registros de ese sector
    col_sec_check = None
    for c in df.columns:
        c_low = str(c).lower().replace('ó', 'o')
        if ('cod' in c_low or 'cd' in c_low) and ('sector' in c_low or 'setor' in c_low):
            col_sec_check = c
            break
        elif 'sector' in c_low or 'setor' in c_low:
            col_sec_check = c
            
    if col_sec_check and not df[col_sec_check].dropna().empty:
        col_sec_found_val = str(df[col_sec_check].dropna().iloc[0]).strip().split('.')[0]
        cursor.execute("DELETE FROM consultoras_tableau WHERE cod_sector = ? OR sector LIKE ?", (col_sec_found_val, f"%{col_sec_found_val}%"))
    else:
        cursor.execute("DELETE FROM consultoras_tableau")
    
    for _, row in df.iterrows():
        cb = str(row.get('Codigo CB') if 'Codigo CB' in df.columns else row.get('Código CB', '')).strip()
        if not cb or cb.lower() == 'nan':
            continue
        
        nom = str(row.get('Asesora / Consultora') if 'Asesora / Consultora' in df.columns else row.get('Nombre', ''))
        col = str(row.get('Nivel / Color') if 'Nivel / Color' in df.columns else row.get('Color', ''))
        
        cursor.execute("""
        INSERT OR REPLACE INTO consultoras_tableau (
            codigo_cb, nombre, documento_gpp, cod_gerencia, gerencia, cod_sector, sector, grupo, ciclo, color,
            ascenso_cb, situacion, sit_comercial, ciclos_inactividad, actividad_convergencia, inicios_completos,
            fact_vol, fact_natura, fact_avon, fact_ce, facturacion_total,
            pts_natura, pts_avon, pts_total_vd, pts_vol, pts_acum, pts_mant, pts_asc,
            tienda_online, tienda_sf, ciclo_primer_pedido, fecha_nacimiento, mes_cumpleanos,
            deuda_total, deuda_mora, credito_total, credito_disponible, pedidos_pendientes, pedidos_mora,
            celular, correo,
            dpto_residencia, ciudad_residencia, barrio_residencia, direccion_residencia, complemento_residencia, referencia_residencia,
            dpto_entrega, ciudad_entrega, barrio_entrega, direccion_entrega, complemento_entrega, referencia_entrega,
            tiempo_casa, origen_cb, notas_lider
        ) VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?,
            ?, ?,
            ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?,
            ?, ?, ?
        )
        """, (
            cb,
            nom,
            str(row.get('DocumentoGPP', '')),
            str(row.get('Cod. Gerencia', '')),
            str(row.get('Gerencia', '')),
            str(row.get('Cod. Sector', '')),
            str(row.get('Sector', '')),
            str(int(limpiar_numero(row.get('Grupo', 0)))) if limpiar_numero(row.get('Grupo', 0)) > 0 else str(row.get('Grupo', '')),
            int(limpiar_numero(row.get('Ciclo', 0))),
            col,
            str(row.get('Ascenso CB', '')),
            str(row.get('Situación', '')),
            str(row.get('Sit. Comercial', '')),
            int(limpiar_numero(row.get('Ciclos Inactividad', 0))),
            str(row.get('Actividad Convergencia', '')),
            str(row.get('Inicios Completos', '')),
            float(limpiar_numero(row.get('Fact. VOL', 0.0))),
            float(limpiar_numero(row.get('Fact. Natura', 0.0))),
            float(limpiar_numero(row.get('Fact. AVON', 0.0))),
            float(limpiar_numero(row.get('Fact. C&E', 0.0))),
            float(limpiar_numero(row.get('Fact. Total', 0.0))),
            int(limpiar_numero(row.get('Pts Natura', 0))),
            int(limpiar_numero(row.get('Pts AVON', 0))),
            int(limpiar_numero(row.get('Pts Total VD', 0))),
            int(limpiar_numero(row.get('Pts VOL', 0))),
            int(limpiar_numero(row.get('Pts Acum', 0))),
            int(limpiar_numero(row.get('Pts Mant', 0))),
            int(limpiar_numero(row.get('Pts Asc', 0))),
            str(row.get('Tienda Online', '')),
            str(row.get('TiendaSF', '')),
            int(limpiar_numero(row.get('Ciclo Primer Pedido', 0))),
            str(row.get('Fecha De Nacimiento', '')),
            str(row.get('Mes Cumpleaños', '')),
            float(limpiar_numero(row.get('Deuda Total', 0.0))),
            float(limpiar_numero(row.get('Deuda Mora', 0.0))),
            float(limpiar_numero(row.get('Credito Total', 0.0))),
            float(limpiar_numero(row.get('Credito Disponible', 0.0))),
            int(limpiar_numero(row.get('Ped. Pendientes', 0))),
            int(limpiar_numero(row.get('Ped. Mora', 0))),
            str(row.get('celular', '')),
            str(row.get('Correo', '')),
            str(row.get('Dpto - Residencia', '')),
            str(row.get('Ciudad - Residencia', '')),
            str(row.get('Barrio - Residencia', '')),
            str(row.get('Dirección - Residencia', '')),
            str(row.get('Complemento - Residencia', '')),
            str(row.get('Referencia - Residencia', '')),
            str(row.get('Dpto - Entrega', '')),
            str(row.get('Ciudad - Entrega', '')),
            str(row.get('Barrio - Entrega', '')),
            str(row.get('Dirección - Entrega', '')),
            str(row.get('Complemento - Entrega', '')),
            str(row.get('Referencia - Entrega', '')),
            int(limpiar_numero(row.get('Tiempo de Casa (Ciclo)', 0))),
            str(row.get('Origen CB', '')),
            str(row.get('Notas / Comentarios Líder', ''))
        ))
    
    conn.commit()
    if close_at_end:
        conn.close()
    return True

def sincronizar_excel_metas_a_sqlite(df_metas, conn=None):
    """
    Convierte y vuelca el DataFrame de metas hacia la tabla metas_como_vamos en SQLite.
    """
    if df_metas is None or df_metas.empty:
        return False
    
    close_at_end = False
    if conn is None:
        conn = obtener_conexion_db()
        close_at_end = True
    
    col_sec = 'Nombre Setor' if 'Nombre Setor' in df_metas.columns else 'Sector'
    if not df_metas[col_sec].dropna().empty:
        sec_val = str(df_metas[col_sec].dropna().iloc[0]).strip()
        cursor.execute("DELETE FROM metas_como_vamos WHERE nombre_sector = ? OR nombre_sector LIKE ?", (sec_val, f"%{sec_val}%"))
    else:
        cursor.execute("DELETE FROM metas_como_vamos")
    
    col_cb = 'Código de consultora' if 'Código de consultora' in df_metas.columns else 'Cd Consultora'
    col_nom = 'Nombre de consultora' if 'Nombre de consultora' in df_metas.columns else 'Nombre Consultora'
    col_ger = 'Nombre Gerencia' if 'Nombre Gerencia' in df_metas.columns else 'Gerencia'
    col_grp = 'Código de grupo' if 'Código de grupo' in df_metas.columns else 'Cód. Grupo'
    
    for _, row in df_metas.iterrows():
        cb = str(row.get(col_cb, '')).strip()
        cursor.execute("""
        INSERT INTO metas_como_vamos (
            codigo_cb, nombre_consultora, nombre_gerencia, nombre_sector, codigo_grupo, color,
            obj_facturacion, real_facturacion, cump_facturacion, obj_activas, real_activas, cump_activas,
            saldo, disponibles, inicios, reinicios, recuperos, ganancia_estimada
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            cb,
            str(row.get(col_nom, '')),
            str(row.get(col_ger, '')),
            str(row.get(col_sec, '')),
            str(int(limpiar_numero(row.get(col_grp, 0)))) if limpiar_numero(row.get(col_grp, 0)) > 0 else str(row.get(col_grp, '')),
            str(row.get('Color', '')),
            float(limpiar_numero(row.get('Objetivo Facturación', 0.0))),
            float(limpiar_numero(row.get('Real Facturación', 0.0))),
            float(limpiar_numero(row.get('Cumplimiento Facturación', 0.0))),
            float(limpiar_numero(row.get('Objetivo Activas', 0.0))),
            float(limpiar_numero(row.get('Real Activas', 0.0))),
            float(limpiar_numero(row.get('Cumplimiento Activas', 0.0))),
            float(limpiar_numero(row.get('Saldo', 0.0))),
            int(limpiar_numero(row.get('Disponibles', 0))),
            int(limpiar_numero(row.get('Inicios', 0))),
            int(limpiar_numero(row.get('Reinicios', 0))),
            int(limpiar_numero(row.get('Recuperos', 0))),
            float(limpiar_numero(row.get('Ganancia estimada', 0.0)))
        ))
        
    conn.commit()
    if close_at_end:
        conn.close()
    return True

def consultar_tableau_sql(grupo=None, sector=None):
    """
    Ejecuta consulta SQL ultra-rápida indexada sobre consultoras_tableau en base_matices.db,
    filtrando por grupo de líder o sector de gerencia.
    """
    conn = obtener_conexion_db()
    query = """
    SELECT 
        codigo_cb AS 'Código CB',
        nombre AS 'Asesora / Consultora',
        documento_gpp AS 'DocumentoGPP',
        cod_gerencia AS 'Cod. Gerencia',
        gerencia AS 'Gerencia',
        cod_sector AS 'Cod. Sector',
        sector AS 'Sector',
        grupo AS 'Grupo',
        ciclo AS 'Ciclo',
        color AS 'Nivel / Color',
        ascenso_cb AS 'Ascenso CB',
        situacion AS 'Situación',
        sit_comercial AS 'Sit. Comercial',
        ciclos_inactividad AS 'Ciclos Inactividad',
        actividad_convergencia AS 'Actividad Convergencia',
        inicios_completos AS 'Inicios Completos',
        fact_vol AS 'Fact. VOL',
        fact_natura AS 'Fact. Natura',
        fact_avon AS 'Fact. AVON',
        fact_ce AS 'Fact. C&E',
        facturacion_total AS 'Fact. Total',
        pts_natura AS 'Pts Natura',
        pts_avon AS 'Pts AVON',
        pts_total_vd AS 'Pts Total VD',
        pts_vol AS 'Pts VOL',
        pts_acum AS 'Pts Acum',
        pts_mant AS 'Pts Mant',
        pts_asc AS 'Pts Asc',
        tienda_online AS 'Tienda Online',
        tienda_sf AS 'TiendaSF',
        ciclo_primer_pedido AS 'Ciclo Primer Pedido',
        fecha_nacimiento AS 'Fecha De Nacimiento',
        mes_cumpleanos AS 'Mes Cumpleaños',
        deuda_total AS 'Deuda Total',
        deuda_mora AS 'Deuda Mora',
        credito_total AS 'Credito Total',
        credito_disponible AS 'Credito Disponible',
        pedidos_pendientes AS 'Ped. Pendientes',
        pedidos_mora AS 'Ped. Mora',
        celular AS 'celular',
        correo AS 'Correo',
        dpto_residencia AS 'Dpto - Residencia',
        ciudad_residencia AS 'Ciudad - Residencia',
        barrio_residencia AS 'Barrio - Residencia',
        direccion_residencia AS 'Dirección - Residencia',
        complemento_residencia AS 'Complemento - Residencia',
        referencia_residencia AS 'Referencia - Residencia',
        dpto_entrega AS 'Dpto - Entrega',
        ciudad_entrega AS 'Ciudad - Entrega',
        barrio_entrega AS 'Barrio - Entrega',
        direccion_entrega AS 'Dirección - Entrega',
        complemento_entrega AS 'Complemento - Entrega',
        referencia_entrega AS 'Referencia - Entrega',
        tiempo_casa AS 'Tiempo de Casa (Ciclo)',
        origen_cb AS 'Origen CB',
        notas_lider AS 'Notas / Comentarios Líder'
    FROM consultoras_tableau
    """
    where_clauses = []
    params = []
    
    if grupo:
        where_clauses.append("(grupo LIKE ? OR sector LIKE ?)")
        params.extend([f"%{grupo}%", f"%{grupo}%"])
        
    if sector and str(sector).strip():
        sec_str = str(sector).strip()
        where_clauses.append("(cod_sector = ? OR sector LIKE ?)")
        params.extend([sec_str, f"%{sec_str}%"])
        
    if where_clauses:
        query += " WHERE " + " AND ".join(where_clauses)
    
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()

    # Añadir alias de columnas para máxima compatibilidad con las pestañas de app.py
    if 'Código CB' in df.columns and 'Codigo CB' not in df.columns:
        df['Codigo CB'] = df['Código CB']
    if 'Asesora / Consultora' in df.columns and 'Nombre' not in df.columns:
        df['Nombre'] = df['Asesora / Consultora']
    if 'Nivel / Color' in df.columns and 'Color' not in df.columns:
        df['Color'] = df['Nivel / Color']
    if 'Notas / Comentarios Líder' in df.columns and 'Comentarios_Lider' not in df.columns:
        df['Comentarios_Lider'] = df['Notas / Comentarios Líder']

    return df

# Ejecutamos la función si se invoca el script directamente
if __name__ == "__main__":
    df_resultado = calcular_metas_ciclo()
    if df_resultado is not None:
        diagnostico = generar_analisis_como_vamos(df_resultado)
        print("=== DIAGNÓSTICO AUTOMÁTICO DE 'CÓMO VAMOS' ===")
        print(f"Facturación Real vs Obj : ${diagnostico['facturacion']['real']:,.0f} / ${diagnostico['facturacion']['objetivo']:,.0f} ({diagnostico['facturacion']['cumplimiento_pct']:.1f}%)")
        print(f"Productividad Promedio  : ${diagnostico['facturacion']['productividad_promedio']:,.0f} COP / activa")
        print(f"Tasa Conversión Disponib: {diagnostico['disponibles']['tasa_conversion_pct']:.1f}% ({diagnostico['activas']['real']:.0f} de {diagnostico['disponibles']['total_disponibles']:.0f})")
        print(f"Total Saldo Pendiente   : {diagnostico['saldos']['total_saldo']:.0f} pedidos/cartera ({diagnostico['saldos']['lideres_afectados']} líderes con saldo)")

