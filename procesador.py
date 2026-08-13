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
    - Excluye consultoras en 'Sit. Comercial' == 'Inactiva 6' (o Ciclos Inactividad == 6).
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

    # Filtrar automáticamente 'Inactiva 6' (Estado No 6)
    if 'Sit. Comercial' in df.columns:
        df = df[df['Sit. Comercial'].astype(str).str.strip().str.lower() != 'inactiva 6'].copy()
    if 'Ciclos Inactividad' in df.columns:
        df = df[pd.to_numeric(df['Ciclos Inactividad'], errors='coerce').fillna(0) != 6].copy()

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

    col_sit_base = None
    for c in df_base.columns:
        if 'Sit. Comercial' in str(c) or 'Situación' in str(c):
            col_sit_base = c
            break
    if not col_sit_base:
        col_sit_base = 'Sit. Comercial'
        df_base[col_sit_base] = ''

    df_base['cb_clean'] = df_base[col_code_base].astype(str).str.strip()

    coincidencias = 0
    cambios = 0
    detalles_cambios = []

    for idx in df_base.index:
        cb = df_base.at[idx, 'cb_clean']
        if cb in mapa_estados:
            coincidencias += 1
            nuevo_estado = mapa_estados[cb]
            estado_actual = str(df_base.at[idx, col_sit_base]).strip()
            if nuevo_estado and nuevo_estado != estado_actual:
                nombre = str(df_base.at[idx, 'Nombre'] if 'Nombre' in df_base.columns else cb)
                detalles_cambios.append({
                    'Código CB': cb,
                    'Asesora / Consultora': nombre,
                    'Estado Anterior': estado_actual,
                    'Nuevo Estado (mi_grupo)': nuevo_estado
                })
                df_base.at[idx, col_sit_base] = nuevo_estado
                cambios += 1

    df_base = df_base.drop(columns=['cb_clean'], errors='ignore')

    # Guardar en Base de Datos.xlsx
    try:
        df_base.to_excel(ruta_base, index=False)
        return {
            'exito': True,
            'coincidencias': coincidencias,
            'cambios': cambios,
            'detalles': detalles_cambios
        }
    except Exception as e:
        return {'exito': False, 'error': f"Error al guardar '{ruta_base}': {e}"}

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

def registrar_o_actualizar_usuario(username, nombre, password, rol, codigo_grupo=None):
    """
    Permite al Gerente crear un usuario de Líder o Asesora o actualizar su contraseña.
    """
    u_clean = str(username).strip().lower()
    if not u_clean:
        return False, "El nombre de usuario no puede estar vacío."
    
    usuarios = cargar_usuarios()
    p_hash = hashlib_sha256(password) if password else (usuarios.get(u_clean, {}).get("password_hash", hashlib_sha256("123456")))
    
    usuarios[u_clean] = {
        "nombre": nombre,
        "password_hash": p_hash,
        "rol": rol,
        "codigo_grupo": str(codigo_grupo).strip() if codigo_grupo else None
    }
    if guardar_usuarios(usuarios):
        return True, f"Usuario '{u_clean}' guardado exitosamente."
    return False, "Error al guardar el archivo de usuarios."

# --- CONFIGURACIÓN DE PERMISOS GLOBALES DE CARGA ---
RUTA_CONFIG = 'configuracion.json'

def cargar_configuracion():
    """
    Carga la configuración global de la aplicación.
    Por defecto, las Líderes y Asesoras tienen bloqueada la subida de archivos (permitir_carga_lideres = False).
    """
    if os.path.exists(RUTA_CONFIG):
        try:
            with open(RUTA_CONFIG, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    config_default = {
        "permitir_carga_lideres": False
    }
    guardar_configuracion(config_default)
    return config_default

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

