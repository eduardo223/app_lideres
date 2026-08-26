import pandas as pd
import os
import io
import sys
import json
import re

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
                cols_prev = [c for c in [col_id, 'Cumplimiento Facturación', 'Real Facturación', 'Inactiva 1', 'Inactiva 2', 'Inactiva 3', 'Inactivas 1', 'Inactivas 2', 'Inactivas 3'] if c in df_prev.columns]
                
                if col_id in cols_prev:
                    df_prev_sub = df_prev[cols_prev].drop_duplicates(subset=[col_id])
                    renames_prev = {c_p: f"{c_p}_anterior" for c_p in cols_prev if c_p != col_id}
                    df_prev_sub = df_prev_sub.rename(columns=renames_prev)
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

    # Cálculo de Cumplimiento Activas (en escala 0-100%)
    if 'Objetivo Activas' in df.columns and 'Real Activas' in df.columns:
        obj_a = pd.to_numeric(df['Objetivo Activas'], errors='coerce')
        real_a = pd.to_numeric(df['Real Activas'], errors='coerce')
        df['Cumplimiento Activas'] = (real_a / obj_a * 100.0).where(obj_a > 0, 0.0)
    elif 'Cumplimiento Activas' in df.columns:
        c_a = pd.to_numeric(df['Cumplimiento Activas'], errors='coerce')
        df['Cumplimiento Activas'] = c_a.apply(lambda v: v * 100.0 if pd.notna(v) and 0 < v <= 2.5 else (v if pd.notna(v) else 0.0))

    # Cálculo dinámico de Activas Frecuentes y % Actividad Frecuente
    if 'Real Activas' in df.columns:
        r_act = pd.to_numeric(df['Real Activas'], errors='coerce').fillna(0.0)
        r_rec = pd.to_numeric(df['Recuperos'], errors='coerce').fillna(0.0) if 'Recuperos' in df.columns else 0.0
        r_ini = pd.to_numeric(df['Inicios'], errors='coerce').fillna(0.0) if 'Inicios' in df.columns else 0.0
        r_rei = pd.to_numeric(df['Reinicios'], errors='coerce').fillna(0.0) if 'Reinicios' in df.columns else 0.0
        df['Activas Frecuentes'] = (r_act - r_rec - r_ini - r_rei).clip(lower=0)
        if 'Disponibles' in df.columns:
            disp_s = pd.to_numeric(df['Disponibles'], errors='coerce').fillna(0.0)
            df['%Actividad Frecuente'] = (df['Activas Frecuentes'] / disp_s.replace(0, pd.NA) * 100.0).fillna(0.0)

    # Cálculo dinámico de Ganancia Estimada según Matriz y Potencializador de Saldo
    df = calcular_ganancia_estimada_df(df)

    # Integrar Metas de Objetivos Arte (Desafíos LNN: Inicios + Reinicios y Recuperos)
    try:
        mapa_arte = cargar_objetivos_arte()
        mapa_grp = mapa_arte.get('por_grupo', {})
        mapa_nom = mapa_arte.get('por_nombre', {})
        
        col_cv_grp = next((c for c in df.columns if 'grupo' in str(c).lower()), None)
        col_cv_nom = next((c for c in df.columns if 'nombre' in str(c).lower() and 'consultora' in str(c).lower()), None)
        
        def _obtener_meta_ini(row):
            g = str(row.get(col_cv_grp, '')).strip().split('.')[0] if col_cv_grp else ''
            nom = str(row.get(col_cv_nom, '')).strip().lower() if col_cv_nom else ''
            target = mapa_grp.get(g) or mapa_nom.get(nom)
            if target:
                return target.get('meta_inicios_reinicios', 0)
            return 0

        def _obtener_meta_rec(row):
            g = str(row.get(col_cv_grp, '')).strip().split('.')[0] if col_cv_grp else ''
            nom = str(row.get(col_cv_nom, '')).strip().lower() if col_cv_nom else ''
            target = mapa_grp.get(g) or mapa_nom.get(nom)
            if target:
                return target.get('meta_recuperos', 0)
            return 0

        def _obtener_meta_disp(row):
            g = str(row.get(col_cv_grp, '')).strip().split('.')[0] if col_cv_grp else ''
            nom = str(row.get(col_cv_nom, '')).strip().lower() if col_cv_nom else ''
            target = mapa_grp.get(g) or mapa_nom.get(nom)
            if target:
                return target.get('disponibles_proyectadas', 0) or target.get('disponibles_esperadas', 0)
            return 0

        def _obtener_desafio_act(row):
            g = str(row.get(col_cv_grp, '')).strip().split('.')[0] if col_cv_grp else ''
            nom = str(row.get(col_cv_nom, '')).strip().lower() if col_cv_nom else ''
            target = mapa_grp.get(g) or mapa_nom.get(nom)
            if target:
                return target.get('desafio_activas', 0)
            return 0

        df['Meta Inicios + Reinicios'] = df.apply(_obtener_meta_ini, axis=1)
        df['Meta Recuperos'] = df.apply(_obtener_meta_rec, axis=1)
        df['Meta Disponibles Esperadas'] = df.apply(_obtener_meta_disp, axis=1)
        df['Desafío Activas Arte'] = df.apply(_obtener_desafio_act, axis=1)
    except Exception as e_arte:
        safe_print(f"Nota al integrar Objetivos Arte en calcular_metas_ciclo: {e_arte}")

    # Si se ejecuta directamente desde archivo, exportamos
    if isinstance(origen, str):
        archivo_salida = 'Resultado_Metas_Procesadas.xlsx'
        try:
            df.to_excel(archivo_salida, index=False)
            safe_print(f"[OK] Resultados guardados exitosamente en: '{archivo_salida}'!")
        except Exception as e:
            safe_print(f"Advertencia al guardar archivo de salida: {e}")

    return df

# --- MÓDULO DE GESTIÓN DE OBJETIVOS ARTE (DESAFÍOS LNN: INICIOS, REINICIOS, RECUPEROS) ---

RUTA_OBJETIVOS_ARTE_JSON = 'objetivos_arte.json'

def cargar_objetivos_arte():
    """
    Carga el diccionario mapeado de Objetivos Arte desde objetivos_arte.json o lo genera desde Objetivos Arte.xlsx si existe.
    """
    if os.path.exists(RUTA_OBJETIVOS_ARTE_JSON):
        try:
            with open(RUTA_OBJETIVOS_ARTE_JSON, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if data and isinstance(data, dict):
                    return data
        except Exception as e:
            safe_print(f"Nota al cargar {RUTA_OBJETIVOS_ARTE_JSON}: {e}")

    # Fallback automático: procesar Objetivos Arte.xlsx si existe localmente
    if os.path.exists('Objetivos Arte.xlsx'):
        res = procesar_archivo_objetivos_arte('Objetivos Arte.xlsx')
        if res.get('exito'):
            return res.get('data', {})
            
    return {}

def guardar_objetivos_arte(dict_data):
    """
    Guarda el diccionario mapeado de Objetivos Arte en objetivos_arte.json.
    """
    try:
        with open(RUTA_OBJETIVOS_ARTE_JSON, 'w', encoding='utf-8') as f:
            json.dump(dict_data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        safe_print(f"Error al guardar {RUTA_OBJETIVOS_ARTE_JSON}: {e}")
        return False

def procesar_archivo_objetivos_arte(origen_arte, ruta_guardar_excel='Objetivos Arte.xlsx'):
    """
    Lee y procesa el archivo Objetivos Arte (.xlsx, .xls), extrayendo las metas de:
    - 'Inicios + reinicios'
    - 'Recuperos'
    de la hoja 'Desafíos LNN'.
    Guarda el mapeo persistente en 'objetivos_arte.json' y actualiza el archivo 'Objetivos Arte.xlsx'.
    """
    try:
        if hasattr(origen_arte, 'read'):
            try:
                origen_arte.seek(0)
            except Exception:
                pass
            with open(ruta_guardar_excel, 'wb') as f_out:
                f_out.write(origen_arte.read())
            try:
                origen_arte.seek(0)
            except Exception:
                pass
            ruta_leer = ruta_guardar_excel
        elif isinstance(origen_arte, str):
            ruta_leer = origen_arte
        else:
            return {'exito': False, 'error': "Formato de origen no soportado."}

        xl = pd.ExcelFile(ruta_leer)
        
        # Buscar hoja Desafíos LNN (insensible a mayúsculas/tildes)
        hoja_des = next((s for s in xl.sheet_names if 'desaf' in s.lower() and 'ln' in s.lower()), None)
        if not hoja_des:
            hoja_des = next((s for s in xl.sheet_names if 'desaf' in s.lower()), xl.sheet_names[0])

        df_raw = xl.parse(hoja_des)
        
        # Detectar cabeceras desplazadas si vienen con 'Unnamed'
        if any('unnamed' in str(c).lower() for c in df_raw.columns[:5]):
            for r_idx in range(min(5, len(df_raw))):
                row_vals = [str(x).lower() for x in df_raw.iloc[r_idx].values if pd.notna(x)]
                if any('sector' in v or 'grupo' in v or 'lider' in v or 'líder' in v for v in row_vals):
                    df_raw.columns = [str(col_name).strip() for col_name in df_raw.iloc[r_idx]]
                    df_raw = df_raw.iloc[r_idx + 1:].reset_index(drop=True)
                    break

        col_grp = next((c for c in df_raw.columns if any(k in str(c).lower() for k in ['cód. grupo', 'cod grupo', 'cód grupo', 'grupo', 'codigo grupo'])), None)
        
        # Priorizar Nombre de Líder exacto evitando 'Cód. Líder'
        col_lider = next((c for c in df_raw.columns if 'nombre' in str(c).lower() and ('lider' in str(c).lower() or 'líder' in str(c).lower())), None)
        if not col_lider:
            col_lider = next((c for c in df_raw.columns if ('lider' in str(c).lower() or 'líder' in str(c).lower()) and 'cód' not in str(c).lower() and 'cod' not in str(c).lower()), None)
            
        col_sec = next((c for c in df_raw.columns if str(c).lower() in ['sector', 'nombre setor', 'nombre sector']), None)
        
        col_ini_meta = next((c for c in df_raw.columns if 'inicios + reinicios' in str(c).lower() or ('inicio' in str(c).lower() and 'meta' in str(c).lower())), None)
        if not col_ini_meta:
            col_ini_meta = next((c for c in df_raw.columns if 'inicios' in str(c).lower() and 'reinicio' in str(c).lower()), None)
        
        col_rec_meta = next((c for c in df_raw.columns if 'recupero' in str(c).lower()), None)
        
        col_disp_esp = next((c for c in df_raw.columns if 'disponibles esperadas' in str(c).lower() and '202612' in str(c).lower()), None)
        if not col_disp_esp:
            col_disp_esp = next((c for c in df_raw.columns if 'disponibles esperadas' in str(c).lower()), None)
            
        col_disp_proy = next((c for c in df_raw.columns if 'disponibles proyectadas' in str(c).lower()), None)
        col_desafio_act = next((c for c in df_raw.columns if 'desafío de activas' in str(c).lower() or 'desafio de activas' in str(c).lower() or 'desafio activas' in str(c).lower() or 'desafío activas' in str(c).lower()), None)
        col_desafio_fact = next((c for c in df_raw.columns if 'desafío facturación' in str(c).lower() or 'desafio facturacion' in str(c).lower() or 'desafio facturación' in str(c).lower()), None)

        if not col_grp and not col_lider:
            return {'exito': False, 'error': "No se identificó la columna de Grupo o Líder en la hoja 'Desafíos LNN'."}

        mapa_por_grupo = {}
        mapa_por_nombre = {}
        sectores_encontrados = set()

        for _, row in df_raw.iterrows():
            g_raw = str(row.get(col_grp, '')).strip().split('.')[0] if col_grp else ''
            nom = str(row.get(col_lider, '')).strip() if col_lider else ''
            sec = str(row.get(col_sec, '')).strip() if col_sec else ''
            if sec:
                sectores_encontrados.add(sec)

            val_ini = int(round(limpiar_numero(row.get(col_ini_meta, 0), 0))) if col_ini_meta else 0
            val_rec = int(round(limpiar_numero(row.get(col_rec_meta, 0), 0))) if col_rec_meta else 0
            val_disp_esp = int(round(limpiar_numero(row.get(col_disp_esp, 0), 0))) if col_disp_esp else 0
            val_disp_proy = int(round(limpiar_numero(row.get(col_disp_proy, 0), 0))) if col_disp_proy else 0
            val_desafio_act = int(round(limpiar_numero(row.get(col_desafio_act, 0), 0))) if col_desafio_act else 0
            val_desafio_fact = float(limpiar_numero(row.get(col_desafio_fact, 0), 0.0)) if col_desafio_fact else 0.0

            nom_limpio = nom
            if ' - ' in nom_limpio:
                nom_limpio = nom_limpio.split(' - ', 1)[1].strip()

            data_lider = {
                'meta_inicios_reinicios': val_ini,
                'meta_recuperos': val_rec,
                'disponibles_esperadas': val_disp_esp if val_disp_esp > 0 else val_disp_proy,
                'disponibles_proyectadas': val_disp_proy,
                'desafio_activas': val_desafio_act,
                'desafio_facturacion': val_desafio_fact,
                'nombre_lider': nom_limpio,
                'grupo': g_raw,
                'sector': sec
            }

            if g_raw and g_raw not in ['-', 'nan', '']:
                mapa_por_grupo[g_raw] = data_lider
            if nom_limpio and nom_limpio not in ['-', 'nan', '']:
                mapa_por_nombre[nom_limpio.lower()] = data_lider

        dict_final = {
            'por_grupo': mapa_por_grupo,
            'por_nombre': mapa_por_nombre
        }

        guardar_objetivos_arte(dict_final)
        extraer_catalogo_sectores_desde_arte(ruta_leer)

        return {
            'exito': True,
            'total_mapeados': len(mapa_por_grupo),
            'sectores': sorted(list(sectores_encontrados)),
            'data': dict_final
        }
    except Exception as e:
        return {'exito': False, 'error': f"Error al procesar Objetivos Arte: {e}"}

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
    Saldo negativo (< 0): Rojo (#FEE2E2), Saldo >= 0: Verde (#DCFCE7).
    """
    try:
        if pd.isna(val):
            return ""
        s = str(val).replace('$', '').replace(',', '').replace(' ', '').strip()
        num = float(s)
        if num < 0:
            return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
        else:
            return 'background-color: #DCFCE7; color: #166534; font-weight: bold;'
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
    wb.remove(wb.active)  # Eliminar hoja por defecto

    # Estilos profesionales
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
                        if 0 < abs(num) <= 2.5:
                            num = num * 100.0
                        
                        if num < 90.0:
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
                        clean_s = str(val).replace('$', '').replace(',', '').strip()
                        num = float(clean_s)
                        if num < 0:
                            cell.fill = fill_red
                            cell.font = font_red
                        else:
                            cell.fill = fill_green
                            cell.font = font_green
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

COLUMNAS_ORDEN_TABLEAU = [
    'Código CB',
    'Líder / Grupo',
    'Asesora / Consultora',
    'Nivel / Color',
    'Sit. Comercial',
    'Pts Acum',
    'Pts Mant',
    'Pts Asc',
    'Deuda Total',
    'Deuda Mora',
    'Credito Total',
    'Credito Disponible',
    'Pts Natura',
    'Pts AVON',
    'Ped. Pendientes',
    'Notas / Comentarios Líder'
]

def limpiar_y_ordenar_columnas_tableau(df_raw, mapa_lideres=None):
    """
    Estandariza cualquier DataFrame de Tableau para que conserve exactamente el mismo orden
    y cantidad de columnas limpias de la tabla 'Base de Datos' / 'Base Maestra Gestionable',
    eliminando información repetida, columnas duplicadas o campos técnicos.
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=COLUMNAS_ORDEN_TABLEAU)

    df = df_raw.copy()

    if mapa_lideres is None:
        mapa_lideres = obtener_mapa_lideres()

    # 1. Enriquecer columna 'Líder / Grupo'
    if 'Líder / Grupo' not in df.columns:
        col_grp = next((c for c in ['Grupo', 'grupo', 'codigo_grupo', 'Código de grupo'] if c in df.columns), None)
        if col_grp:
            df['Líder / Grupo'] = df[col_grp].apply(
                lambda g: f"Grupo {g} — {mapa_lideres[str(g).strip()]}" if str(g).strip() in mapa_lideres else f"Grupo {g}"
            )
        else:
            df['Líder / Grupo'] = ''

    # 2. Diccionario de normalización a nombres de cabecera canónicos
    rename_dict = {}
    
    # Código CB
    if 'Código CB' not in df.columns:
        c_cb = next((c for c in ['Codigo CB', 'codigo_cb', 'Código de consultora', 'Cd Consultora'] if c in df.columns), None)
        if c_cb: rename_dict[c_cb] = 'Código CB'

    # Asesora / Consultora
    if 'Asesora / Consultora' not in df.columns:
        c_nom = next((c for c in ['Nombre', 'nombre', 'Nombre de consultora', 'Consultora'] if c in df.columns), None)
        if c_nom: rename_dict[c_nom] = 'Asesora / Consultora'

    # Nivel / Color
    if 'Nivel / Color' not in df.columns:
        c_col = next((c for c in ['Color', 'color', 'Nivel', 'nivel'] if c in df.columns), None)
        if c_col: rename_dict[c_col] = 'Nivel / Color'

    # Sit. Comercial
    if 'Sit. Comercial' not in df.columns:
        c_sit = next((c for c in ['sit_comercial', 'Sit Comercial', 'Situación', 'situacion', 'Situacion'] if c in df.columns), None)
        if c_sit: rename_dict[c_sit] = 'Sit. Comercial'

    # Puntos
    if 'Pts Acum' not in df.columns:
        c_pts_ac = next((c for c in ['pts_acum', 'Pts Acumulados'] if c in df.columns), None)
        if c_pts_ac: rename_dict[c_pts_ac] = 'Pts Acum'

    if 'Pts Mant' not in df.columns:
        c_pts_mt = next((c for c in ['pts_mant', 'Pts Para Mantener'] if c in df.columns), None)
        if c_pts_mt: rename_dict[c_pts_mt] = 'Pts Mant'

    if 'Pts Asc' not in df.columns:
        c_pts_as = next((c for c in ['pts_asc', 'Pts para Ascender', 'Pts para Ascender '] if c in df.columns), None)
        if c_pts_as: rename_dict[c_pts_as] = 'Pts Asc'

    # Deuda y Crédito
    if 'Deuda Total' not in df.columns:
        c_dt = next((c for c in ['deuda_total'] if c in df.columns), None)
        if c_dt: rename_dict[c_dt] = 'Deuda Total'

    if 'Deuda Mora' not in df.columns:
        c_dm = next((c for c in ['deuda_mora', 'Deuda Mora '] if c in df.columns), None)
        if c_dm: rename_dict[c_dm] = 'Deuda Mora'

    if 'Credito Total' not in df.columns:
        c_ct = next((c for c in ['credito_total', 'Crédito Total'] if c in df.columns), None)
        if c_ct: rename_dict[c_ct] = 'Credito Total'

    if 'Credito Disponible' not in df.columns:
        c_cd = next((c for c in ['credito_disponible', 'Crédito Disponible'] if c in df.columns), None)
        if c_cd: rename_dict[c_cd] = 'Credito Disponible'

    if 'Pts Natura' not in df.columns:
        c_pn = next((c for c in ['pts_natura'] if c in df.columns), None)
        if c_pn: rename_dict[c_pn] = 'Pts Natura'

    if 'Pts AVON' not in df.columns:
        c_pa = next((c for c in ['pts_avon'] if c in df.columns), None)
        if c_pa: rename_dict[c_pa] = 'Pts AVON'

    if 'Ped. Pendientes' not in df.columns:
        c_pp = next((c for c in ['pedidos_pendientes', 'Pedidos Pendientes'] if c in df.columns), None)
        if c_pp: rename_dict[c_pp] = 'Ped. Pendientes'

    if 'Notas / Comentarios Líder' not in df.columns:
        c_nl = next((c for c in ['Comentarios_Lider', 'notas_lider', 'Notas / Comentarios'] if c in df.columns), None)
        if c_nl: rename_dict[c_nl] = 'Notas / Comentarios Líder'

    if rename_dict:
        df = df.rename(columns=rename_dict)

    # Eliminar duplicados de columnas
    df = df.loc[:, ~df.columns.duplicated()].copy()

    # Seleccionar única y exclusivamente las columnas en el orden canónico
    cols_existentes = [c for c in COLUMNAS_ORDEN_TABLEAU if c in df.columns]
    df_resultado = df[cols_existentes].copy()

    return df_resultado

def obtener_base_tableau_completa_original(grupo=None, sector=None, cbs_filtrados=None):
    """
    Retorna el DataFrame completo de Tableau con las 61+ columnas originales en su orden exacto,
    incluyendo los comentarios persistentes de la líder y aplicando los filtros y ordenamiento correspondientes.
    """
    if os.path.exists('Base de Datos.xlsx'):
        try:
            df = pd.read_excel('Base de Datos.xlsx')
            
            # Detectar y promover cabecera real si viene desplazada
            if any('unnamed' in str(c).lower() for c in df.columns[:5]):
                for r_idx in range(min(10, len(df))):
                    row_vals = [str(x).lower() for x in df.iloc[r_idx].values if pd.notna(x)]
                    if any('codigo' in x or 'código' in x or 'cb' in x for x in row_vals):
                        df.columns = [str(col_name).strip() for col_name in df.iloc[r_idx]]
                        df = df.iloc[r_idx + 1:].reset_index(drop=True)
                        break

            # Limpiar codificaciones comunes en encabezados
            df = df.rename(columns=lambda c: str(c).replace('Situacin', 'Situación').replace('Mes Cumpleaos', 'Mes Cumpleaños').replace('Direccin', 'Dirección'))
            
            col_cb = next((c for c in ['Codigo CB', 'Código CB', 'codigo_cb'] if c in df.columns), df.columns[0])
            comentarios = cargar_comentarios_lideres()
            
            df['cb_clean_tmp'] = df[col_cb].apply(limpiar_codigo_cb_estandar)
            df['Notas / Comentarios Líder'] = df['cb_clean_tmp'].map(lambda k: comentarios.get(k, ''))
            
            # 1. Si se pasan CBs filtrados directamente de la consulta activa en pantalla
            if cbs_filtrados is not None:
                cbs_clean_list = [limpiar_codigo_cb_estandar(x) for x in cbs_filtrados if limpiar_codigo_cb_estandar(x) != '']
                cbs_set = set(cbs_clean_list)
                df = df[df['cb_clean_tmp'].isin(cbs_set)].copy()
                
                # Mantener estrictamente el orden de los filtros de la pantalla
                order_map = {cb: idx for idx, cb in enumerate(cbs_clean_list)}
                df['__sort_rank__'] = df['cb_clean_tmp'].map(order_map)
                df = df.sort_values(by='__sort_rank__').drop(columns=['cb_clean_tmp', '__sort_rank__'], errors='ignore')
                return df

            # 2. Si se proporciona grupo o sector
            if grupo and str(grupo).strip() and str(grupo).strip() != 'Todas las Líderes (Consolidado Zona)':
                g_str = str(grupo).strip()
                mask_grp = pd.Series(False, index=df.index)
                for c_grp in ['Grupo', 'grupo', 'codigo_grupo', 'Código de grupo', 'Cód. Grupo']:
                    if c_grp in df.columns:
                        mask_grp = mask_grp | df[c_grp].astype(str).str.strip().str.split('.').str[0].str.contains(g_str, case=False, na=False)
                if mask_grp.any():
                    df = df[mask_grp].copy()

            if sector and str(sector).strip() and str(sector).strip() != 'Todos':
                s_str = str(sector).strip()
                mask_sec = pd.Series(False, index=df.index)
                for c_sec in ['Cod. Sector', 'cod_sector', 'Codigo Sector', 'Código Sector']:
                    if c_sec in df.columns:
                        mask_sec = mask_sec | df[c_sec].astype(str).str.strip().str.split('.').str[0].str.contains(s_str, case=False, na=False)
                for c_sec in ['Sector ', 'Sector', 'sector', 'Nombre Setor', 'Nombre Sector']:
                    if c_sec in df.columns:
                        mask_sec = mask_sec | df[c_sec].astype(str).str.strip().str.contains(s_str, case=False, na=False)
                if mask_sec.any():
                    df = df[mask_sec].copy()
                    
            df = df.drop(columns=['cb_clean_tmp'], errors='ignore')
            return df
        except Exception as e:
            safe_print(f"Nota al leer Base de Datos.xlsx para exportación: {e}")

    # Fallback a SQLite
    df_sql = consultar_tableau_sql(grupo=grupo, sector=sector)
    if df_sql is not None and not df_sql.empty:
        col_cb_sql = next((c for c in ['Código CB', 'Codigo CB', 'codigo_cb'] if c in df_sql.columns), df_sql.columns[0])
        df_sql['cb_clean_tmp'] = df_sql[col_cb_sql].apply(limpiar_codigo_cb_estandar)
        if cbs_filtrados is not None:
            cbs_clean_list = [limpiar_codigo_cb_estandar(x) for x in cbs_filtrados if limpiar_codigo_cb_estandar(x) != '']
            cbs_set = set(cbs_clean_list)
            df_sql = df_sql[df_sql['cb_clean_tmp'].isin(cbs_set)].copy()
            order_map = {cb: idx for idx, cb in enumerate(cbs_clean_list)}
            df_sql['__sort_rank__'] = df_sql['cb_clean_tmp'].map(order_map)
            df_sql = df_sql.sort_values(by='__sort_rank__').drop(columns=['cb_clean_tmp', '__sort_rank__'], errors='ignore')
        else:
            df_sql = df_sql.drop(columns=['cb_clean_tmp'], errors='ignore')
    return df_sql

def exportar_tableau_excel_con_colores(df, nombre_hoja="Base_Consultoras", buffer_salida=None):
    """
    Genera un archivo Excel profesional (.xlsx) con los colores exactos de la consulta,
    manteniendo estrictamente el orden y todas las columnas originales de la base de datos.
    Optimizado para alta velocidad (bulk append y pre-indexación de columnas).
    """
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    df = df.loc[:, ~df.columns.duplicated()].copy()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = str(nombre_hoja)[:31]

    # Paleta de colores para niveles
    fill_bronce = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    font_bronce = Font(name="Calibri", size=10, bold=True, color="92400E")
    
    fill_plata = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    font_plata = Font(name="Calibri", size=10, bold=True, color="334155")
    
    fill_oro = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    font_oro = Font(name="Calibri", size=10, bold=True, color="854D0E")
    
    fill_platino = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
    font_platino = Font(name="Calibri", size=10, bold=True, color="0369A1")
    
    fill_zafiro = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    font_zafiro = Font(name="Calibri", size=10, bold=True, color="1E40AF")
    
    fill_diamante = PatternFill(start_color="F3E8FF", end_color="F3E8FF", fill_type="solid")
    font_diamante = Font(name="Calibri", size=10, bold=True, color="6B21A8")

    # Paleta para situación comercial
    fill_activa = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    font_activa = Font(name="Calibri", size=10, bold=True, color="166534")
    
    fill_inactiva_leve = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    font_inactiva_leve = Font(name="Calibri", size=10, bold=True, color="854D0E")
    
    fill_inactiva_critica = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    font_inactiva_critica = Font(name="Calibri", size=10, bold=True, color="991B1B")

    fill_mora_media = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")
    font_mora_media = Font(name="Calibri", size=10, bold=True, color="9A3412")

    fill_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    headers = list(df.columns)
    ws.append(headers)

    col_nivel_indices = []
    col_sit_indices = []
    col_mora_indices = []
    col_money_indices = []
    col_int_indices = []

    for col_idx, col_name in enumerate(headers, 1):
        c_low = str(col_name).lower()
        if c_low in ['color', 'nivel', 'nivel / color']:
            col_nivel_indices.append(col_idx)
        elif any(k in c_low for k in ['sit. comercial', 'sit comercial', 'situacion', 'situación']):
            col_sit_indices.append(col_idx)
        elif 'deuda mora' in c_low or c_low.startswith('mora'):
            col_mora_indices.append(col_idx)
        elif any(k in c_low for k in ['deuda', 'credito', 'crédito', 'fact']):
            col_money_indices.append(col_idx)
        elif any(k in c_low for k in ['pts', 'puntos', 'pedidos', 'ped.', 'ciclos']):
            col_int_indices.append(col_idx)

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_vals in df.itertuples(index=False):
        ws.append(list(row_vals))

    num_filas = len(df)
    for r_idx in range(2, num_filas + 2):
        for col_idx in col_nivel_indices:
            cell = ws.cell(row=r_idx, column=col_idx)
            v_str = str(cell.value or '').strip().lower()
            if 'bronce' in v_str:
                cell.fill, cell.font = fill_bronce, font_bronce
            elif 'plata' in v_str:
                cell.fill, cell.font = fill_plata, font_plata
            elif 'oro' in v_str:
                cell.fill, cell.font = fill_oro, font_oro
            elif 'platino' in v_str:
                cell.fill, cell.font = fill_platino, font_platino
            elif 'zafiro' in v_str:
                cell.fill, cell.font = fill_zafiro, font_zafiro
            elif 'diamante' in v_str:
                cell.fill, cell.font = fill_diamante, font_diamante

        for col_idx in col_sit_indices:
            cell = ws.cell(row=r_idx, column=col_idx)
            v_str = str(cell.value or '').strip().lower()
            if 'activa' in v_str or 'inicio' in v_str or 'reinicio' in v_str or 'recupero' in v_str:
                cell.fill, cell.font = fill_activa, font_activa
            elif any(x in v_str for x in ['inactiva 1', 'inactiva 2', 'inactiva 3', 'inactiva1', 'inactiva2', 'inactiva3']):
                cell.fill, cell.font = fill_inactiva_leve, font_inactiva_leve
            elif any(x in v_str for x in ['inactiva 4', 'inactiva 5', 'inactiva 6', 'indisponible', 'inactiva4', 'inactiva5', 'inactiva6']):
                cell.fill, cell.font = fill_inactiva_critica, font_inactiva_critica
            elif 'fuga' in v_str:
                cell.fill, cell.font = fill_mora_media, font_mora_media

        for col_idx in col_mora_indices:
            cell = ws.cell(row=r_idx, column=col_idx)
            num = limpiar_numero(cell.value, 0.0)
            if num <= 0:
                cell.fill, cell.font = fill_activa, font_activa
            elif num <= 200000:
                cell.fill, cell.font = fill_inactiva_leve, font_inactiva_leve
            elif num <= 500000:
                cell.fill, cell.font = fill_mora_media, font_mora_media
            else:
                cell.fill, cell.font = fill_inactiva_critica, font_inactiva_critica
            cell.number_format = '$#,##0'

        for col_idx in col_money_indices:
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.number_format = '$#,##0'

        for col_idx in col_int_indices:
            cell = ws.cell(row=r_idx, column=col_idx)
            cell.number_format = '#,##0'

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = 16

    if buffer_salida is not None:
        wb.save(buffer_salida)
        return buffer_salida
    else:
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

def exportar_tabla_ods(df_dict_or_df, buffer_salida=None):
    """
    Exporta un DataFrame o diccionario de DataFrames a formato OpenDocument Spreadsheet (.ods).
    """
    out = buffer_salida if buffer_salida is not None else io.BytesIO()
    with pd.ExcelWriter(out, engine='odf') as writer:
        if isinstance(df_dict_or_df, dict):
            for s_name, df_item in df_dict_or_df.items():
                df_item.to_excel(writer, sheet_name=str(s_name)[:31], index=False)
        else:
            df_dict_or_df.to_excel(writer, sheet_name='Datos', index=False)
    if buffer_salida is not None:
        return buffer_salida
    return out.getvalue()

def exportar_tabla_pdf(df, titulo="Reporte Ejecutivo - Panel Matices", subtitulo="", columnas=None, max_filas=1000, buffer_salida=None):
    """
    Genera un documento PDF profesional horizontal (Landscape) con tabla estilizada,
    encabezados corporativos y semáforo de colores en celdas de estado, nivel y mora.
    """
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    out = buffer_salida if buffer_salida is not None else io.BytesIO()
    doc = SimpleDocTemplate(
        out,
        pagesize=landscape(letter),
        leftMargin=20,
        rightMargin=20,
        topMargin=25,
        bottomMargin=25
    )

    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=15,
        leading=17,
        textColor=colors.HexColor('#1E293B'),
        spaceAfter=3
    )

    sub_style = ParagraphStyle(
        'SubStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor('#64748B'),
        spaceAfter=10
    )

    cell_style = ParagraphStyle(
        'CellStyle',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=7,
        leading=8,
        textColor=colors.HexColor('#1E293B')
    )

    cell_header_style = ParagraphStyle(
        'CellHeaderStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=7.5,
        leading=9,
        textColor=colors.white,
        alignment=1
    )

    story.append(Paragraph(f"<b>{titulo}</b>", title_style))
    if subtitulo:
        story.append(Paragraph(subtitulo, sub_style))

    # Seleccionar columnas prioritarias si no se pasan explícitas
    if columnas:
        cols_usar = [c for c in columnas if c in df.columns]
    else:
        cols_prioritarias = [
            'Código CB', 'Líder / Grupo', 'Asesora / Consultora', 'Nivel / Color', 'Sit. Comercial',
            'Pts Acum', 'Deuda Total', 'Deuda Mora', 'Credito Disponible', 'Ped. Pendientes', 'Notas / Comentarios Líder'
        ]
        cols_usar = [c for c in cols_prioritarias if c in df.columns]
        if not cols_usar:
            cols_usar = list(df.columns[:10])

    df_pdf = df[cols_usar].head(max_filas).copy()

    table_data = []
    header_row = [Paragraph(f"<b>{c}</b>", cell_header_style) for c in cols_usar]
    table_data.append(header_row)

    custom_table_styles = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('TOPPADDING', (0, 0), (-1, -1), 2.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 2.5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2.5),
    ]

    for r_idx, (_, row) in enumerate(df_pdf.iterrows(), start=1):
        row_cells = []
        for c_idx, col_name in enumerate(cols_usar):
            val = row[col_name]
            val_str = str(val) if pd.notna(val) else ""
            col_lower = str(col_name).lower()

            # Formatear números
            if any(k in col_lower for k in ['deuda', 'credito', 'crédito', 'fact']):
                num = limpiar_numero(val, 0.0)
                val_str = f"${num:,.0f}".replace(",", ".")
            elif any(k in col_lower for k in ['pts', 'puntos', 'pedidos', 'ped.']):
                num = int(limpiar_numero(val, 0))
                val_str = f"{num:,}".replace(",", ".")

            # Colorear celdas según estado / nivel / mora
            if 'color' in col_lower or 'nivel' in col_lower:
                v_low = str(val).lower()
                if 'bronce' in v_low:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#FEF3C7')))
                elif 'plata' in v_low:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#E2E8F0')))
                elif 'oro' in v_low:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#FEF08A')))
                elif 'platino' in v_low:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#E0F2FE')))
                elif 'zafiro' in v_low:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#DBEAFE')))
                elif 'diamante' in v_low:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#F3E8FF')))

            elif 'sit. comercial' in col_lower or 'sit comercial' in col_lower or 'situacion' in col_lower or 'situación' in col_lower:
                v_low = str(val).lower()
                if 'activa' in v_low or 'inicio' in v_low or 'reinicio' in v_low or 'recupero' in v_low:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#DCFCE7')))
                elif any(x in v_low for x in ['inactiva 1', 'inactiva 2', 'inactiva 3']):
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#FEF9C3')))
                elif any(x in v_low for x in ['inactiva 4', 'inactiva 5', 'inactiva 6', 'indisponible']):
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#FEE2E2')))
                elif 'fuga' in v_low:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#FFEDD5')))

            elif 'deuda mora' in col_lower or 'mora' in col_lower:
                num = limpiar_numero(val, 0.0)
                if num <= 0:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#DCFCE7')))
                elif num <= 200000:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#FEF9C3')))
                elif num <= 500000:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#FFEDD5')))
                else:
                    custom_table_styles.append(('BACKGROUND', (c_idx, r_idx), (c_idx, r_idx), colors.HexColor('#FEE2E2')))

            row_cells.append(Paragraph(val_str[:30], cell_style))
        table_data.append(row_cells)

    disponible_width = 750
    col_w = disponible_width / max(1, len(cols_usar))
    t = Table(table_data, colWidths=[col_w] * len(cols_usar), repeatRows=1)
    t.setStyle(TableStyle(custom_table_styles))
    story.append(t)

    doc.build(story)
    if buffer_salida is not None:
        return buffer_salida
    return out.getvalue()

# --- MÓDULO INFORME TABLEAU MANAGER ("INFORME TABLEAU CAM") ---

RUTA_COMENTARIOS = 'comentarios_lideres.json'

def autocorregir_texto_espanol(texto):
    """
    Normaliza el texto de notas eliminando espacios innecesarios.
    """
    return str(texto).strip() if texto else ""

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
    nota_limpia = str(comentario).strip()
    comentarios[codigo_str] = nota_limpia
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
    Actualiza de forma directa y liviana tanto 'comentarios_lideres.json' como la base SQLite 'consultoras_tableau'.
    """
    comentarios = cargar_comentarios_lideres()
    dict_limpio = {}
    for cb, nota in dict_comentarios.items():
        cb_str = str(cb).strip()
        nota_str = str(nota).strip()
        dict_limpio[cb_str] = nota_str
        if nota_str:
            comentarios[cb_str] = nota_str
        elif cb_str in comentarios and nota_str == "":
            comentarios.pop(cb_str, None)
    
    try:
        with open(RUTA_COMENTARIOS, 'w', encoding='utf-8') as f:
            json.dump(comentarios, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Error guardando comentarios masivos en JSON: {e}")

    try:
        conn = obtener_conexion_db()
        cursor = conn.cursor()
        for cb, nota_str in dict_limpio.items():
            cursor.execute("UPDATE consultoras_tableau SET notas_lider = ? WHERE codigo_cb = ?", (nota_str, cb))
        conn.commit()
        conn.close()
    except Exception as e_sql:
        print(f"Nota actualizando SQLite en guardar_todos_comentarios: {e_sql}")

    return True

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
    if 'Sit. Comercial' in df.columns or 'Situación' in df.columns:
        if 'Sit. Comercial' not in df.columns:
            df['Sit. Comercial'] = ''
        if 'Situación' not in df.columns:
            df['Situación'] = ''
            
        mask_activa_cycle = pd.Series(False, index=df.index)
        if 'Pts Total VD' in df.columns:
            mask_activa_cycle = mask_activa_cycle | (pd.to_numeric(df['Pts Total VD'], errors='coerce').fillna(0) > 0)
        if 'Fact. Total' in df.columns:
            mask_activa_cycle = mask_activa_cycle | (pd.to_numeric(df['Fact. Total'], errors='coerce').fillna(0) > 0)
        if 'Situación' in df.columns:
            mask_activa_cycle = mask_activa_cycle | (df['Situación'].astype(str).str.strip().str.lower() == 'activa')
        if 'Sit. Comercial' in df.columns:
            mask_activa_cycle = mask_activa_cycle | (df['Sit. Comercial'].astype(str).str.strip().str.lower() == 'activa')
        if 'Indicador' in df.columns:
            mask_activa_cycle = mask_activa_cycle | (df['Indicador'].astype(str).str.strip().str.lower() == 'activas')
        
        df.loc[mask_activa_cycle, 'Sit. Comercial'] = 'Activa'
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

def limpiar_codigo_cb_estandar(v):
    """
    Limpia y estandariza cualquier código CB o documento:
    - Remueve NaN/None.
    - Remueve sufijo .0 de conversiones a float.
    - Quita espacios en blanco.
    - Remueve ceros a la izquierda si es un valor numérico para asegurar coincidencia consistente.
    """
    if pd.isna(v):
        return ""
    s = str(v).strip()
    if s.endswith('.0'):
        s = s[:-2].strip()
    if s.isdigit():
        s_num = s.lstrip('0')
        return s_num if s_num else "0"
    return s

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
    df_grupo['cb_clean'] = df_grupo[col_code_grupo].apply(limpiar_codigo_cb_estandar)
    df_grupo['estado_clean'] = df_grupo[col_estado_grupo].apply(normalizar_estado_mi_grupo)

    mapa_estados = dict(zip(df_grupo['cb_clean'], df_grupo['estado_clean']))

    # Cargar la base principal
    df_base = pd.read_excel(ruta_base, sheet_name=0)
    # Detectar y promover cabecera real si viene con 'Unnamed' (formato estándar Tableau)
    if any('unnamed' in str(c).lower() for c in df_base.columns[:5]):
        for r_idx in range(min(10, len(df_base))):
            row_vals = [str(x).lower() for x in df_base.iloc[r_idx].values if pd.notna(x)]
            if any('codigo' in x or 'código' in x or 'cb' in x or 'asesora' in x or 'consultora' in x for x in row_vals):
                df_base.columns = [str(col_name).strip() for col_name in df_base.iloc[r_idx]]
                df_base = df_base.iloc[r_idx + 1:].reset_index(drop=True)
                break

    df_base.columns = [str(c).replace('\ufffd', 'ó').strip() for c in df_base.columns]

    col_code_base = None
    for c in df_base.columns:
        c_str = str(c).strip().lower()
        if any(k in c_str for k in ['codigo cb', 'código cb', 'código de consultora', 'codigo de consultora', 'cd consultora']):
            col_code_base = c
            break
    if not col_code_base:
        col_code_base = df_base.columns[0]

    col_sit_comercial = next((c for c in df_base.columns if 'sit. comercial' in str(c).lower() or 'sit comercial' in str(c).lower()), None)
    col_situacion_macro = next((c for c in df_base.columns if str(c).strip().lower() in ['situación', 'situacion']), None)

    if not col_sit_comercial:
        col_sit_comercial = 'Sit. Comercial'
        df_base[col_sit_comercial] = ''
    if not col_situacion_macro:
        col_situacion_macro = 'Situación'
        df_base[col_situacion_macro] = ''

    df_base['cb_clean'] = df_base[col_code_base].apply(limpiar_codigo_cb_estandar)

    for c in [col_sit_comercial, col_situacion_macro]:
        if c and c in df_base.columns:
            df_base[c] = df_base[c].astype(object)

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
                nombre = str(df_base.at[idx, 'Nombre'] if 'Nombre' in df_base.columns else (df_base.at[idx, 'Asesora / Consultora'] if 'Asesora / Consultora' in df_base.columns else cb))
                detalles_cambios.append({
                    'Código CB': cb,
                    'Asesora / Consultora': nombre,
                    'Estado Anterior': estado_actual,
                    'Nuevo Estado (mi_grupo)': nuevo_estado
                })
                df_base.at[idx, col_sit_comercial] = nuevo_estado
                if col_situacion_macro:
                    if 'activa' in nuevo_estado.lower():
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
    Cruce del archivo de 'activas' (.xlsx, .xls, .csv) con la base de datos de Tableau:
    - Identifica los Códigos CB del archivo de activas.
    - Actualiza únicamente el campo 'Indicador' con el valor 'Activas'.
    - Sincroniza 'Sit. Comercial' y 'Situación' a 'Activa'.
    - No modifica otros campos (Facturación, Puntos, Pedidos se preservan intactos).
    - Guarda en 'Base de Datos.xlsx' y sincroniza la tabla SQLite consultoras_tableau.
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
                for enc in ['utf-8', 'utf-8-sig', 'latin1', 'iso-8859-1']:
                    try:
                        df_act = pd.read_csv(origen_activas, encoding=enc, sep=None, engine='python')
                        if df_act is not None and not df_act.empty:
                            break
                    except Exception:
                        continue
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
                for enc in ['utf-8', 'utf-8-sig', 'latin1', 'iso-8859-1']:
                    try:
                        origen_activas.seek(0)
                        df_act = pd.read_csv(origen_activas, encoding=enc, sep=None, engine='python')
                        if df_act is not None and not df_act.empty:
                            break
                    except Exception:
                        continue
            else:
                try:
                    origen_activas.seek(0)
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
    if not any(any(k in str(c).upper() for k in ['CODIGO', 'CB', 'DOC', 'ASESORA', 'CONSULTORA', 'NOMBRE', 'CEDULA', 'IDENT']) for c in df_act.columns):
        for r_idx in range(min(15, len(df_act))):
            row_vals = [str(x).upper() for x in df_act.iloc[r_idx].values if pd.notna(x)]
            if any('COD' in v or 'DOC' in v or 'CONSULTORA' in v or 'ASESORA' in v or 'CEDULA' in v for v in row_vals):
                df_act.columns = df_act.iloc[r_idx]
                df_act = df_act.iloc[r_idx + 1:].reset_index(drop=True)
                break

    # 2. Identificar columna de Código CB o Documento en el archivo de activas
    col_code_act = None
    for c in df_act.columns:
        c_up = str(c).strip().upper()
        if ('DIGO' in c_up or 'CODIGO' in c_up or 'CB' in c_up) and 'NOMBRE' not in c_up and 'GERENCIA' not in c_up and 'SECTOR' not in c_up and 'GRUPO' not in c_up:
            col_code_act = c
            break
    if not col_code_act:
        for c in df_act.columns:
            c_up = str(c).strip().upper()
            if any(k in c_up for k in ['DOC', 'CEDULA', 'IDENT', 'CONSULTORA', 'ASESORA', 'ID_CONSULTORA', 'ID_ASESORA']):
                if 'NOMBRE' not in c_up and 'GERENCIA' not in c_up and 'SECTOR' not in c_up and 'GRUPO' not in c_up:
                    col_code_act = c
                    break

    if not col_code_act:
        return {'exito': False, 'error': "No se identificó la columna de Código o Documento en el archivo de activas."}

    # Obtener el conjunto de Códigos CB del archivo de activas
    df_act['cb_clean'] = df_act[col_code_act].apply(limpiar_codigo_cb_estandar)
    cbs_activas = set(df_act['cb_clean'].dropna().loc[lambda s: s != ''])

    # 3. Cargar Base de Datos.xlsx principal
    df_base = pd.read_excel(ruta_base, sheet_name=0)
    # Detectar y promover cabecera real si viene con 'Unnamed' (formato estándar Tableau)
    if any('unnamed' in str(c).lower() for c in df_base.columns[:5]):
        for r_idx in range(min(10, len(df_base))):
            row_vals = [str(x).lower() for x in df_base.iloc[r_idx].values if pd.notna(x)]
            if any('codigo' in x or 'código' in x or 'cb' in x or 'asesora' in x or 'consultora' in x for x in row_vals):
                df_base.columns = [str(col_name).strip() for col_name in df_base.iloc[r_idx]]
                df_base = df_base.iloc[r_idx + 1:].reset_index(drop=True)
                break

    df_base.columns = [str(c).replace('\ufffd', 'ó').strip() for c in df_base.columns]

    col_code_base = None
    for c in df_base.columns:
        c_str = str(c).strip().lower()
        if any(k in c_str for k in ['codigo cb', 'código cb', 'código de consultora', 'codigo de consultora', 'cd consultora']):
            col_code_base = c
            break
    if not col_code_base:
        col_code_base = df_base.columns[0]

    # Identificar o crear columna 'Indicador'
    col_indicador = next((c for c in df_base.columns if str(c).strip().lower() in ['indicador', 'indicadores']), None)
    if not col_indicador:
        col_indicador = 'Indicador'
        df_base[col_indicador] = ''

    col_sit_comercial = next((c for c in df_base.columns if 'sit. comercial' in str(c).lower() or 'sit comercial' in str(c).lower()), None)
    col_situacion_macro = next((c for c in df_base.columns if str(c).strip().lower() in ['situación', 'situacion']), None)

    if not col_sit_comercial:
        col_sit_comercial = 'Sit. Comercial'
        df_base[col_sit_comercial] = ''
    if not col_situacion_macro:
        col_situacion_macro = 'Situación'
        df_base[col_situacion_macro] = ''

    df_base['cb_clean'] = df_base[col_code_base].apply(limpiar_codigo_cb_estandar)

    # Asegurar compatibilidad de tipos (PyArrow / Pandas 2+ / Python 3.14)
    for c in [col_indicador, col_sit_comercial, col_situacion_macro]:
        if c and c in df_base.columns:
            df_base[c] = df_base[c].astype(object)

    coincidencias = 0
    cambios_totales = 0
    detalles_cambios = []
    cbs_base_set = set(df_base['cb_clean'].dropna())

    for idx in df_base.index:
        cb = df_base.at[idx, 'cb_clean']
        if cb in cbs_activas:
            coincidencias += 1
            ind_actual = str(df_base.at[idx, col_indicador]).strip() if pd.notna(df_base.at[idx, col_indicador]) else ''
            
            # Actualizar Indicador a 'Activas'
            df_base.at[idx, col_indicador] = 'Activas'
            df_base.at[idx, col_sit_comercial] = 'Activa'
            if col_situacion_macro:
                df_base.at[idx, col_situacion_macro] = 'Activa'

            if ind_actual != 'Activas':
                cambios_totales += 1
                nombre = str(df_base.at[idx, 'Nombre'] if 'Nombre' in df_base.columns else (df_base.at[idx, 'Asesora / Consultora'] if 'Asesora / Consultora' in df_base.columns else cb))
                detalles_cambios.append({
                    'Código CB': cb,
                    'Asesora / Consultora': nombre,
                    'Indicador Anterior': ind_actual or 'N/D',
                    'Nuevo Indicador': 'Activas',
                    'Sit. Comercial': 'Activa'
                })

    df_base = df_base.drop(columns=['cb_clean'], errors='ignore')

    # Guardar en Base de Datos.xlsx y refrescar SQLite
    try:
        df_base.to_excel(ruta_base, index=False)
        try:
            sincronizar_excel_tableau_a_sqlite(ruta_base)
        except Exception as e_sql:
            safe_print(f"Advertencia al sincronizar SQLite tras cruce de activas: {e_sql}")

        no_encontradas = [cb for cb in cbs_activas if cb not in cbs_base_set]

        return {
            'exito': True,
            'total_activas_archivo': len(cbs_activas),
            'coincidencias': coincidencias,
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

# --- MÓDULO DE SUSCRIPCIONES, PRUEBAS GRATIS (15 DÍAS) Y CONTROL ANTI-FRAUDE ---
RUTA_HISTORICO_SECTORES = 'sectores_historico.json'
RUTA_MARCA_AGUA_TIEMPO = 'marca_agua_sistema.json'

def cargar_historico_sectores():
    """
    Carga el historial de sectores registrados para el control de pruebas únicas de 15 días.
    """
    if os.path.exists(RUTA_HISTORICO_SECTORES):
        try:
            with open(RUTA_HISTORICO_SECTORES, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    
    sectores_init = {
        "700000459": {
            "codigo_sector": "700000459",
            "nombre_sector": "MATICES CLERY",
            "primera_prueba_fecha": "2026-01-01T00:00:00",
            "ha_consumido_prueba": True,
            "ha_pagado": True,
            "estado": "activo",
            "fecha_vencimiento": None,
            "correo_gerente": "gerente",
            "telefono_gerente": "3057939537"
        },
        "700000466": {
            "codigo_sector": "700000466",
            "nombre_sector": "DOLLY",
            "primera_prueba_fecha": "2026-01-01T00:00:00",
            "ha_consumido_prueba": True,
            "ha_pagado": True,
            "estado": "activo",
            "fecha_vencimiento": None,
            "correo_gerente": "gerente2",
            "telefono_gerente": ""
        }
    }
    guardar_historico_sectores(sectores_init)
    return sectores_init

def guardar_historico_sectores(dict_sectores):
    """
    Persiste el histórico de sectores en sectores_historico.json.
    """
    try:
        with open(RUTA_HISTORICO_SECTORES, 'w', encoding='utf-8') as f:
            json.dump(dict_sectores, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        safe_print(f"Error al guardar histórico de sectores: {e}")
def obtener_nombre_sector_usuario(user_info):
    """
    Retorna el nombre del sector del usuario de forma dinámica.
    Prioridad:
    1. Campo 'nombre_sector' en el perfil del usuario.
    2. Lookup en sectores_historico.json mediante 'codigo_sector'.
    3. Si es superadmin -> 'Gestión Corporativa Global'
    4. Fallback -> 'Sector {codigo_sector}' o 'Liderazgo Empresarial'.
    """
    if not user_info or not isinstance(user_info, dict):
        return "Liderazgo Empresarial"
        
    rol = user_info.get("rol", "")
    if rol == "superadmin":
        return "Gestión Corporativa Global"
        
    nom_sec = user_info.get("nombre_sector")
    if nom_sec and str(nom_sec).strip() and str(nom_sec).lower() not in ["none", "nan", "null", ""]:
        return str(nom_sec).strip()
        
    cod_sec = str(user_info.get("codigo_sector") or "").strip()
    if cod_sec and cod_sec.lower() not in ["none", "nan", "null", ""]:
        historico = cargar_historico_sectores()
        if cod_sec in historico:
            nom_hist = historico[cod_sec].get("nombre_sector")
            if nom_hist and str(nom_hist).strip() and str(nom_hist).lower() not in ["none", "nan", "null", ""]:
                return str(nom_hist).strip()
        return f"Sector {cod_sec}"
        
    return "Liderazgo Empresarial"

def verificar_estado_suscripcion(user_info_o_sector):
    """
    Determina si un usuario o sector tiene acceso permitido al sistema.
    Retorna un diccionario con:
    - permitido: bool
    - estado: 'superadmin' | 'activo' | 'prueba' | 'vencido' | 'bloqueado'
    - dias_restantes: int
    - fecha_vencimiento_str: str
    - motivo: str
    """
    from datetime import datetime

    if isinstance(user_info_o_sector, dict):
        rol = user_info_o_sector.get("rol", "")
        if rol == "superadmin":
            return {
                "permitido": True,
                "estado": "superadmin",
                "dias_restantes": 9999,
                "fecha_vencimiento_str": "Permanente (Super Administrador)",
                "motivo": "Acceso Total Super Administrador"
            }
        cod_sector = str(user_info_o_sector.get("codigo_sector") or "").strip()
        user_estado = user_info_o_sector.get("estado_suscripcion")
        user_vence = user_info_o_sector.get("fecha_vencimiento")
    else:
        cod_sector = str(user_info_o_sector or "").strip()
        user_estado = None
        user_vence = None

    if not cod_sector:
        return {
            "permitido": True,
            "estado": "activo",
            "dias_restantes": 9999,
            "fecha_vencimiento_str": "Activo",
            "motivo": "Cuenta activa"
        }

    historico = cargar_historico_sectores()
    sec_info = historico.get(cod_sector, {})
    
    estado = sec_info.get("estado", user_estado or "activo")
    vence_iso = sec_info.get("fecha_vencimiento", user_vence)
    
    # 1. Si está bloqueado explícitamente por el Administrador
    if estado == "bloqueado":
        return {
            "permitido": False,
            "estado": "bloqueado",
            "dias_restantes": 0,
            "fecha_vencimiento_str": "Bloqueado",
            "motivo": "El acceso para este sector ha sido suspendido por el Administrador."
        }

    # 2. Si no tiene fecha de vencimiento (acceso permanente / cliente activo)
    if not vence_iso:
        return {
            "permitido": True,
            "estado": "activo",
            "dias_restantes": 9999,
            "fecha_vencimiento_str": "Suscripción Activa (Permanente)",
            "motivo": "Suscripción activa y vigente"
        }

    # 3. Evaluar fecha de vencimiento
    try:
        dt_vence = datetime.fromisoformat(vence_iso)
        dt_now = datetime.now()
        
        diff = (dt_vence - dt_now).total_seconds()
        dias_restantes = max(0, int(diff // 86400) + 1)
        fecha_str = dt_vence.strftime("%d/%m/%Y a las %H:%M")

        if diff >= 0:
            return {
                "permitido": True,
                "estado": estado,
                "dias_restantes": dias_restantes,
                "fecha_vencimiento_str": fecha_str,
                "motivo": f"Vigente hasta el {fecha_str}"
            }
        else:
            return {
                "permitido": False,
                "estado": "vencido",
                "dias_restantes": 0,
                "fecha_vencimiento_str": fecha_str,
                "motivo": f"Tu periodo de {'prueba de 15 días' if estado == 'prueba' else 'suscripción'} ha finalizado el {fecha_str}."
            }
    except Exception as e:
        return {
            "permitido": True,
            "estado": "activo",
            "dias_restantes": 9999,
            "fecha_vencimiento_str": "Activo",
            "motivo": f"Vigente ({e})"
        }

# --- CATÁLOGO CORPORATIVO DE SECTORES Y AUTO-APROVISIONAMIENTO DE LÍDERES ---

RUTA_CATALOGO_SECTORES_JSON = 'catalogo_sectores.json'

def limpiar_nombre_sector_solo(nombre_sector):
    """
    Retorna únicamente el nombre del sector sin el nombre de la persona ni códigos adicionales.
    Ejemplo: 'SECTOR COLORES  KAREN' -> 'SECTOR COLORES'
             'SECTOR ARTESANÍA FERNANDA' -> 'SECTOR ARTESANÍA'
             'SECTOR MATICES CLERY' -> 'SECTOR MATICES'
             'Sector ABÁNICO Judy' -> 'SECTOR ABÁNICO'
    """
    if not nombre_sector:
        return ""
    partes = str(nombre_sector).strip().split()
    if len(partes) >= 2:
        if partes[0].lower() == 'sector':
            return f"SECTOR {partes[1].upper()}"
    return str(nombre_sector).strip().upper()

def cargar_catalogo_sectores():
    """
    Retorna el diccionario estructurado de sectores y líderes desde catalogo_sectores.json.
    Si no existe, intenta extraerlo desde Objetivos Arte.xlsx.
    """
    if os.path.exists(RUTA_CATALOGO_SECTORES_JSON):
        try:
            with open(RUTA_CATALOGO_SECTORES_JSON, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if data and isinstance(data, dict):
                    return data
        except Exception as e:
            safe_print(f"Nota al cargar {RUTA_CATALOGO_SECTORES_JSON}: {e}")

    if os.path.exists('Objetivos Arte.xlsx'):
        return extraer_catalogo_sectores_desde_arte('Objetivos Arte.xlsx')

    return {}

def extraer_catalogo_sectores_desde_arte(ruta_o_buffer='Objetivos Arte.xlsx'):
    """
    Analiza la hoja 'Desafíos LNN' de Objetivos Arte y extrae todos los sectores con sus códigos oficiales
    y la lista completa de sus líderes de negocio (con código de grupo y nombre oficial).
    """
    catalogo = {}
    try:
        xl = pd.ExcelFile(ruta_o_buffer)
        hoja_des = next((s for s in xl.sheet_names if 'desaf' in s.lower() and 'ln' in s.lower()), None)
        if not hoja_des:
            hoja_des = next((s for s in xl.sheet_names if 'desaf' in s.lower()), xl.sheet_names[0])
            
        df_raw = xl.parse(hoja_des)

        if any('unnamed' in str(c).lower() for c in df_raw.columns[:5]):
            for r_idx in range(min(5, len(df_raw))):
                row_vals = [str(x).lower() for x in df_raw.iloc[r_idx].values if pd.notna(x)]
                if any('sector' in v or 'grupo' in v or 'lider' in v or 'líder' in v for v in row_vals):
                    df_raw.columns = [str(col_name).strip() for col_name in df_raw.iloc[r_idx]]
                    df_raw = df_raw.iloc[r_idx + 1:].reset_index(drop=True)
                    break

        col_sec = next((c for c in df_raw.columns if str(c).lower() in ['sector', 'nombre setor', 'nombre sector']), None)
        col_cod = next((c for c in df_raw.columns if 'cod' in str(c).lower() and 'sector' in str(c).lower()), None)
        col_grp = next((c for c in df_raw.columns if any(k in str(c).lower() for k in ['cód. grupo', 'cod grupo', 'grupo'])), None)
        col_lc = next((c for c in df_raw.columns if 'cód. líder' in str(c).lower() or 'cod lider' in str(c).lower()), None)
        col_ln = next((c for c in df_raw.columns if 'nombre líder' in str(c).lower() or 'nombre lider' in str(c).lower()), None)

        if not col_sec:
            return catalogo

        for sec_name, grp in df_raw.groupby(col_sec):
            sec_clean = str(sec_name).strip()
            if not sec_clean or sec_clean.lower() in ['nan', '-', 'none', '']:
                continue
                
            cod_s = ""
            if col_cod and not grp[col_cod].dropna().empty:
                cod_s = str(grp[col_cod].dropna().iloc[0]).split('.')[0].strip()
                
            lideres = []
            for _, r in grp.iterrows():
                g = str(r.get(col_grp, '')).split('.')[0].strip() if col_grp else ''
                lc = str(r.get(col_lc, '')).split('.')[0].strip() if col_lc else ''
                ln = str(r.get(col_ln, '')).strip() if col_ln else ''
                
                if not ln or ln.lower() in ['nan', '-', 'none'] or 'fuera de grupo' in ln.lower():
                    continue
                    
                clean_name = ln.split(' - ', 1)[1].strip() if ' - ' in ln else ln
                lideres.append({
                    'codigo_grupo': g,
                    'codigo_consultora': lc if lc != '-' else '',
                    'nombre_lider': clean_name,
                    'nombre_original': ln
                })

            nombre_solo = limpiar_nombre_sector_solo(sec_clean)
            clave_sector = cod_s if cod_s else nombre_solo
            catalogo[clave_sector] = {
                'codigo_sector': cod_s,
                'nombre_sector': nombre_solo,
                'nombre_sector_original': sec_clean,
                'total_lideres': len(lideres),
                'lideres': lideres
            }

        with open(RUTA_CATALOGO_SECTORES_JSON, 'w', encoding='utf-8') as f:
            json.dump(catalogo, f, ensure_ascii=False, indent=2)

    except Exception as e:
        safe_print(f"Error al extraer catálogo de sectores desde Objetivos Arte: {e}")

    return catalogo

def auto_aprovisionar_lideres_sector(cod_sector, nombre_sector=""):
    """
    Crea o vincula automáticamente las cuentas de usuario de todas las Líderes de Negocio
    asociadas a un sector específico basándose en catalogo_sectores.json.
    """
    creadas = []
    sec_clean = str(cod_sector).strip()
    if not sec_clean:
        return creadas

    catalogo = cargar_catalogo_sectores()
    sec_info = catalogo.get(sec_clean)
    if not sec_info:
        for k, v in catalogo.items():
            if str(v.get('nombre_sector', '')).strip().lower() == str(nombre_sector).strip().lower():
                sec_info = v
                break

    if not sec_info or not sec_info.get('lideres'):
        return creadas

    usuarios = cargar_usuarios()
    cambios = False

    for lider in sec_info.get('lideres', []):
        g = lider.get('codigo_grupo', '').strip()
        nom = lider.get('nombre_lider', '').strip()
        if not g or g.lower() in ['nan', '-', 'none', '0']:
            continue

        username = f"lider{g}".lower()
        if username not in usuarios:
            usuarios[username] = {
                "nombre": nom,
                "password_hash": hashlib_sha256(f"Lider{g}*2026"),
                "rol": "lider",
                "codigo_grupo": g,
                "codigo_sector": sec_clean,
                "nombre_sector": nombre_sector or sec_info.get('nombre_sector', ''),
                "telefono": "",
                "debe_cambiar_password": False
            }
            creadas.append({'username': username, 'nombre': nom, 'grupo': g})
            cambios = True
        else:
            user_u = usuarios[username]
            if not user_u.get('codigo_sector') or user_u.get('codigo_sector') != sec_clean:
                user_u['codigo_sector'] = sec_clean
                user_u['nombre_sector'] = nombre_sector or sec_info.get('nombre_sector', '')
                cambios = True

    if cambios:
        guardar_usuarios(usuarios)

    return creadas

def registrar_nueva_gerente(nombre, correo, password, telefono, cod_sector, nombre_sector=""):
    """
    Registra a una nueva Gerente de forma autónoma con 15 días de prueba gratis.
    Valida que el correo no exista y que el código de sector no haya usado ya la prueba gratuita.
    Además, aprovisiona automáticamente las cuentas de todas las líderes de su sector.
    """
    from datetime import datetime, timedelta
    
    u_clean = str(correo).strip().lower()
    sec_clean = str(cod_sector).strip()
    nom_clean = str(nombre).strip()
    tel_clean = str(telefono).strip()
    sec_nom_clean = str(nombre_sector).strip() or f"Sector {sec_clean}"

    if not nom_clean:
        return False, "Por favor ingresa tu nombre completo.", None
    if not u_clean or '@' not in u_clean or '.' not in u_clean:
        return False, "Por favor ingresa un correo electrónico válido (ej. tu_correo@gmail.com).", None
    if not password or len(str(password).strip()) < 6:
        return False, "La contraseña debe tener mínimo 6 caracteres.", None
    if not sec_clean or len(sec_clean) < 3:
        return False, "Por favor ingresa un Código de Sector válido (ej. 700000459).", None

    usuarios = cargar_usuarios()
    if u_clean in usuarios:
        return False, f"El correo '{u_clean}' ya está registrado. Por favor inicia sesión o recupera tu contraseña.", None

    # Candado Anti-Fraude: Verificar si el sector ya disfrutó de su prueba gratis
    historico = cargar_historico_sectores()
    if sec_clean in historico:
        sec_reg = historico[sec_clean]
        if sec_reg.get("ha_consumido_prueba", False) and not sec_reg.get("ha_pagado", False):
            return False, (
                f"⚠️ **El Código de Sector {sec_clean} ({sec_reg.get('nombre_sector', '')}) ya utilizó su periodo de prueba gratuita de 15 días.**\n\n"
                f"Para activar la suscripción de este sector y habilitar a todo tu equipo de líderes, "
                f"comunícate con Soporte Administrativo al WhatsApp **3057939537**."
            ), None

    now = datetime.now()
    vencimiento = now + timedelta(days=15)
    
    nuevo_usuario = {
        "nombre": nom_clean,
        "password_hash": hashlib_sha256(password),
        "rol": "gerente",
        "codigo_grupo": None,
        "codigo_sector": sec_clean,
        "nombre_sector": sec_nom_clean,
        "telefono": tel_clean,
        "fecha_registro": now.isoformat(),
        "fecha_vencimiento": vencimiento.isoformat(),
        "estado_suscripcion": "prueba",
        "debe_cambiar_password": False
    }

    usuarios[u_clean] = nuevo_usuario
    guardar_usuarios(usuarios)

    # Actualizar histórico de sectores
    historico[sec_clean] = {
        "codigo_sector": sec_clean,
        "nombre_sector": sec_nom_clean,
        "primera_prueba_fecha": now.isoformat(),
        "ha_consumido_prueba": True,
        "ha_pagado": False,
        "estado": "prueba",
        "fecha_vencimiento": vencimiento.isoformat(),
        "correo_gerente": u_clean,
        "telefono_gerente": tel_clean
    }
    guardar_historico_sectores(historico)

    # Auto-aprovisionar instantáneamente las cuentas de las líderes de este sector
    lideres_creadas = auto_aprovisionar_lideres_sector(sec_clean, sec_nom_clean)
    msg_lideres = f" y {len(lideres_creadas)} cuentas de Líderes aprovisionadas automáticamente" if lideres_creadas else ""

    user_session_info = nuevo_usuario.copy()
    user_session_info["username"] = u_clean

    return True, f"¡Bienvenida {nom_clean}! Tu cuenta de Gerente y tu prueba gratis de 15 días han sido activadas exitosamente{msg_lideres}.", user_session_info

def actualizar_suscripcion_sector(cod_sector, nuevo_estado, dias_extension=0, es_pago=False):
    """
    Permite al Super Administrador:
    - Activar plan pagado (+30, +90, +365 días o permanente)
    - Dar prórroga de prueba (+5 días)
    - Suspender o desbloquear un sector
    Aplica el cambio en cascada a la Gerente y a todas sus Líderes en usuarios.json y sectores_historico.json.
    """
    from datetime import datetime, timedelta

    sec_clean = str(cod_sector).strip()
    historico = cargar_historico_sectores()
    usuarios = cargar_usuarios()

    now = datetime.now()

    if sec_clean not in historico:
        historico[sec_clean] = {
            "codigo_sector": sec_clean,
            "nombre_sector": f"Sector {sec_clean}",
            "primera_prueba_fecha": now.isoformat(),
            "ha_consumido_prueba": True,
            "ha_pagado": es_pago,
            "estado": nuevo_estado
        }

    vence_iso = None
    if dias_extension > 0:
        dt_base = now
        curr_vence = historico[sec_clean].get("fecha_vencimiento")
        if curr_vence:
            try:
                dt_curr = datetime.fromisoformat(curr_vence)
                if dt_curr > now:
                    dt_base = dt_curr
            except Exception:
                pass
        vence_iso = (dt_base + timedelta(days=dias_extension)).isoformat()
    elif dias_extension == -1:
        vence_iso = None
    else:
        vence_iso = historico[sec_clean].get("fecha_vencimiento")

    historico[sec_clean]["estado"] = nuevo_estado
    historico[sec_clean]["fecha_vencimiento"] = vence_iso
    if es_pago:
        historico[sec_clean]["ha_pagado"] = True
    guardar_historico_sectores(historico)

    cambiados = 0
    for u_k, u_v in usuarios.items():
        if str(u_v.get("codigo_sector") or "").strip() == sec_clean:
            u_v["estado_suscripcion"] = nuevo_estado
            u_v["fecha_vencimiento"] = vence_iso
            cambiados += 1

    guardar_usuarios(usuarios)
    return True, f"Sector {sec_clean} actualizado a '{nuevo_estado}'. {cambiados} cuentas asociadas actualizadas."

def obtener_resumen_suscripciones():
    """
    Retorna un DataFrame con todos los sectores registrados y el estado de sus suscripciones para el panel de Super Admin.
    """
    from datetime import datetime
    historico = cargar_historico_sectores()
    usuarios = cargar_usuarios()

    filas = []
    now = datetime.now()

    for sec_id, info in historico.items():
        nom_sec = info.get("nombre_sector", f"Sector {sec_id}")
        estado = info.get("estado", "activo")
        vence_iso = info.get("fecha_vencimiento")
        correo_g = info.get("correo_gerente", "")
        tel_g = info.get("telefono_gerente", "")
        ha_pagado = info.get("ha_pagado", False)

        nom_gerente = ""
        total_lideres = 0
        for u_id, u_data in usuarios.items():
            if str(u_data.get("codigo_sector") or "").strip() == str(sec_id).strip():
                if u_data.get("rol") == "gerente":
                    nom_gerente = u_data.get("nombre", "")
                    if not correo_g:
                        correo_g = u_id
                    if not tel_g:
                        tel_g = u_data.get("telefono", "")
                elif u_data.get("rol") == "lider":
                    total_lideres += 1

        dias_rest = "Indefinido"
        vence_str = "Permanente"
        estado_label = "🟢 Activo (Pagado)" if ha_pagado else "🟢 Activo"

        if vence_iso:
            try:
                dt_vence = datetime.fromisoformat(vence_iso)
                diff = (dt_vence - now).total_seconds()
                dias_num = max(0, int(diff // 86400) + 1)
                vence_str = dt_vence.strftime("%d/%m/%Y")
                
                if estado == "bloqueado":
                    estado_label = "⛔ Suspendido / Bloqueado"
                    dias_rest = "0 días"
                elif diff < 0:
                    estado_label = "🔴 Vencido (Requiere Pago)"
                    dias_rest = "0 días (Expirado)"
                elif estado == "prueba":
                    estado_label = f"⏳ En Prueba ({dias_num} días)"
                    dias_rest = f"{dias_num} días restantes"
                else:
                    estado_label = f"🟢 Activo ({dias_num} días)"
                    dias_rest = f"{dias_num} días restantes"
            except Exception:
                pass
        elif estado == "bloqueado":
            estado_label = "⛔ Suspendido / Bloqueado"
            dias_rest = "0 días"

        filas.append({
            "Código Sector": sec_id,
            "Nombre Sector": nom_sec,
            "Gerente Responsable": nom_gerente or correo_g or "N/D",
            "Contacto (WhatsApp)": tel_g or "N/D",
            "Líderes Activas": total_lideres,
            "Estado": estado_label,
            "Vence el": vence_str,
            "Tiempo Restante": dias_rest,
            "_raw_estado": estado,
            "_raw_sector": sec_id
        })

    return pd.DataFrame(filas)

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

def registrar_o_actualizar_usuario(username, nombre, password=None, rol="lider", codigo_grupo=None, codigo_sector=None, debe_cambiar_password=None, nombre_sector=None):
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
    
    sec_id = None
    if codigo_sector:
        sec_id = str(codigo_sector).strip()
        usr_data["codigo_sector"] = sec_id
    elif u_clean in usuarios and "codigo_sector" in usuarios[u_clean]:
        sec_id = str(usuarios[u_clean]["codigo_sector"]).strip()
        usr_data["codigo_sector"] = sec_id

    # Asignar o heredar nombre del sector
    if nombre_sector:
        usr_data["nombre_sector"] = str(nombre_sector).strip()
    elif u_clean in usuarios and "nombre_sector" in usuarios[u_clean]:
        usr_data["nombre_sector"] = usuarios[u_clean]["nombre_sector"]

    # Si pertenece a un sector, heredar estado de suscripción, fecha de vencimiento y nombre de sector si no lo tiene
    if sec_id:
        historico = cargar_historico_sectores()
        if sec_id in historico:
            usr_data["estado_suscripcion"] = historico[sec_id].get("estado", "activo")
            usr_data["fecha_vencimiento"] = historico[sec_id].get("fecha_vencimiento")
            if "nombre_sector" not in usr_data or not usr_data["nombre_sector"]:
                usr_data["nombre_sector"] = historico[sec_id].get("nombre_sector", f"Sector {sec_id}")
        elif nombre_sector:
            # Registrar nuevo sector en el histórico si no existía
            from datetime import datetime
            historico[sec_id] = {
                "codigo_sector": sec_id,
                "nombre_sector": str(nombre_sector).strip(),
                "primera_prueba_fecha": datetime.now().isoformat(),
                "ha_consumido_prueba": True,
                "ha_pagado": True,
                "estado": "activo"
            }
            guardar_historico_sectores(historico)

    usuarios[u_clean] = usr_data
    if guardar_usuarios(usuarios):
        sincronizar_usuarios_a_sqlite()
        return True, f"Usuario '{u_clean}' guardado exitosamente."
    return False, "Error al guardar el archivo de usuarios."

def eliminar_usuario_perfil(username, eliminar_sector_asociado=False):
    """
    Elimina la cuenta de un usuario de usuarios.json y sincroniza con SQLite.
    Protege la cuenta 'admin' para evitar auto-bloqueo.
    Si eliminar_sector_asociado es True, también limpia el sector de sectores_historico.json.
    """
    u_clean = str(username).strip().lower()
    if not u_clean:
        return False, "Nombre de usuario no válido."

    if u_clean == "admin":
        return False, "⚠️ No se puede eliminar la cuenta principal del Super Administrador ('admin')."

    usuarios = cargar_usuarios()
    if u_clean not in usuarios:
        return False, f"El usuario '{u_clean}' no existe en el sistema."

    u_data = usuarios[u_clean]
    sec_id = u_data.get("codigo_sector")
    rol = u_data.get("rol")

    del usuarios[u_clean]
    guardar_usuarios(usuarios)
    sincronizar_usuarios_a_sqlite()

    # Si se solicitó eliminar el registro del sector en el histórico
    if eliminar_sector_asociado and sec_id and rol == "gerente":
        historico = cargar_historico_sectores()
        if str(sec_id) in historico:
            del historico[str(sec_id)]
            guardar_historico_sectores(historico)

    return True, f"¡Perfil de usuario '{u_clean}' eliminado exitosamente!"

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
    Detecta automáticamente los grupos de líderes y crea o actualiza sus cuentas de usuario en el sistema,
    vinculándolas al código de sector correspondiente.
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

    # Extraer código de sector si existe en los archivos
    sec_detectado = None
    if df_tab is not None:
        col_s_tab = next((c for c in ['Cod. Sector', 'cod_sector', 'Codigo Sector'] if c in df_tab.columns), None)
        if col_s_tab and not df_tab[col_s_tab].dropna().empty:
            sec_detectado = str(df_tab[col_s_tab].dropna().iloc[0]).split('.')[0].strip()
    if not sec_detectado and 'Cod. Sector' in df_metas.columns and not df_metas['Cod. Sector'].dropna().empty:
        sec_detectado = str(df_metas['Cod. Sector'].dropna().iloc[0]).split('.')[0].strip()

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
        sec_lider = sec_detectado
        if df_tab is not None:
            mask_g = (df_tab['Grupo'].astype(str).str.strip() == g_str) if 'Grupo' in df_tab.columns else pd.Series(False, index=df_tab.index)
            if mask_g.any():
                df_g = df_tab[mask_g]
                if 'Correo' in df_g.columns and not df_g.dropna(subset=['Correo']).empty:
                    correo_lider = str(df_g['Correo'].dropna().iloc[0]).strip().lower()
                col_s_g = next((c for c in ['Cod. Sector', 'cod_sector'] if c in df_g.columns), None)
                if col_s_g and not df_g[col_s_g].dropna().empty:
                    sec_lider = str(df_g[col_s_g].dropna().iloc[0]).split('.')[0].strip()

        username = correo_lider if (correo_lider and '@' in correo_lider) else f"lider{g_str}"
        ya_existe = (username in usuarios_existentes)

        if ya_existe:
            registrar_o_actualizar_usuario(
                username=username,
                nombre=nom_lider,
                password=None,
                rol="lider",
                codigo_grupo=g_str,
                codigo_sector=sec_lider
            )
        else:
            pass_gen = generar_password_aleatoria()
            exito, msg = registrar_o_actualizar_usuario(
                username=username,
                nombre=nom_lider,
                password=pass_gen,
                rol="lider",
                codigo_grupo=g_str,
                codigo_sector=sec_lider
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

def obtener_mapa_lideres():
    """
    Retorna un diccionario { 'codigo_grupo': 'Nombre Líder' }
    recuperando nombres desde usuarios.json, la base de datos SQLite y 'Base para el como vamos.xlsx'.
    """
    mapa = {}

    # 1. Desde usuarios.json
    try:
        usuarios = cargar_usuarios()
        for u_id, u_data in usuarios.items():
            grp = str(u_data.get('codigo_grupo') or '').strip()
            nom = str(u_data.get('nombre') or '').strip()
            if grp and nom and grp.lower() not in ['none', 'nan', 'null', ''] and nom.lower() not in ['none', 'nan', 'null', '']:
                mapa[grp] = nom
    except Exception:
        pass

    # 2. Desde la tabla SQLite metas_como_vamos
    try:
        conn = obtener_conexion_db()
        cursor = conn.cursor()
        rows = cursor.execute("SELECT DISTINCT codigo_grupo, nombre_consultora FROM metas_como_vamos WHERE codigo_grupo IS NOT NULL AND nombre_consultora IS NOT NULL").fetchall()
        for r in rows:
            grp = str(r[0] or '').strip().split('.')[0]
            nom = str(r[1] or '').strip()
            if grp and nom and grp.lower() not in ['none', 'nan', 'null', ''] and nom.lower() not in ['none', 'nan', 'null', '']:
                mapa[grp] = nom
        conn.close()
    except Exception:
        pass

    # 3. Desde el archivo físico 'Base para el como vamos.xlsx' si existe
    try:
        if os.path.exists('Base para el como vamos.xlsx'):
            df_m = pd.read_excel('Base para el como vamos.xlsx', sheet_name=0)
            df_m = normalizar_columnas(df_m)
            col_grp = 'Código de grupo' if 'Código de grupo' in df_m.columns else None
            col_nom = 'Nombre de consultora' if 'Nombre de consultora' in df_m.columns else None
            if col_grp and col_nom:
                for _, row in df_m.iterrows():
                    g = str(row[col_grp] or '').strip().split('.')[0]
                    n = str(row[col_nom] or '').strip()
                    if g and n and g.lower() not in ['none', 'nan', 'null', ''] and n.lower() not in ['none', 'nan', 'null', '']:
                        mapa[g] = n
    except Exception:
        pass

    return mapa

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
        notas_lider TEXT,
        indicador TEXT
    )
    """)
    try:
        cursor.execute("ALTER TABLE consultoras_tableau ADD COLUMN indicador TEXT")
    except Exception:
        pass
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
    try:
        cursor.execute("ALTER TABLE consultoras_tableau ADD COLUMN indicador TEXT")
    except Exception:
        pass
    
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
        cb = limpiar_codigo_cb_estandar(row.get('Codigo CB') if 'Codigo CB' in df.columns else row.get('Código CB', ''))
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
            tiempo_casa, origen_cb, notas_lider, indicador
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
            ?, ?, ?, ?
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
            str(row.get('Notas / Comentarios Líder', '')),
            str(row.get('Indicador', ''))
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
    try:
        conn.cursor().execute("ALTER TABLE consultoras_tableau ADD COLUMN indicador TEXT")
    except Exception:
        pass
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
        notas_lider AS 'Notas / Comentarios Líder',
        indicador AS 'Indicador'
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

def eliminar_datos_por_grupo_o_usuario(codigo_grupo, eliminar_cuenta=False):
    """
    Elimina todos los registros asociados a un código de grupo en las tablas SQLite
    (consultoras_tableau, metas_como_vamos, comentarios_lideres), en los archivos Excel locales
    (Base para el como vamos.xlsx) y opcionalmente su usuario en usuarios.json.
    """
    conn = obtener_conexion_db()
    cursor = conn.cursor()
    grp_str = str(codigo_grupo).strip()

    borrados = {}

    # 1. Borrar en consultoras_tableau
    try:
        cursor.execute("DELETE FROM consultoras_tableau WHERE TRIM(grupo) = ?", (grp_str,))
        borrados['consultoras_tableau'] = cursor.rowcount
    except Exception:
        borrados['consultoras_tableau'] = 0

    # 2. Borrar en metas_como_vamos
    try:
        cursor.execute("DELETE FROM metas_como_vamos WHERE TRIM(codigo_grupo) = ?", (grp_str,))
        borrados['metas_como_vamos'] = cursor.rowcount
    except Exception:
        borrados['metas_como_vamos'] = 0

    # 3. Borrar en comentarios_lideres
    try:
        cursor.execute("DELETE FROM comentarios_lideres WHERE TRIM(grupo) = ? OR TRIM(username) = ?", (grp_str, grp_str))
        borrados['comentarios_lideres'] = cursor.rowcount
    except Exception:
        borrados['comentarios_lideres'] = 0

    # 4. Borrar en usuarios (SQLite y usuarios.json)
    if eliminar_cuenta:
        try:
            cursor.execute("DELETE FROM usuarios WHERE TRIM(codigo_grupo) = ? OR TRIM(username) = ?", (grp_str, grp_str))
            borrados['usuarios_sqlite'] = cursor.rowcount
        except Exception:
            borrados['usuarios_sqlite'] = 0
        
        try:
            usuarios = cargar_usuarios()
            a_eliminar = [u for u, info in usuarios.items() if str(info.get('codigo_grupo')).strip() == grp_str or str(u).strip() == grp_str]
            for u in a_eliminar:
                del usuarios[u]
            guardar_usuarios(usuarios)
            borrados['usuarios_json'] = len(a_eliminar)
        except Exception:
            borrados['usuarios_json'] = 0

    conn.commit()
    conn.close()

    # 5. Filtrar y actualizar el archivo Excel físico 'Base para el como vamos.xlsx'
    try:
        if os.path.exists("Base para el como vamos.xlsx"):
            xl = pd.ExcelFile("Base para el como vamos.xlsx")
            with pd.ExcelWriter("Base para el como vamos.xlsx", engine="openpyxl") as writer:
                for sheet_name in xl.sheet_names:
                    df_s = pd.read_excel("Base para el como vamos.xlsx", sheet_name=sheet_name)
                    col_g = None
                    for c in df_s.columns:
                        if 'grupo' in str(c).lower() or 'codigo_grupo' in str(c).lower():
                            col_g = c
                            break
                    if col_g:
                        df_s = df_s[df_s[col_g].astype(str).str.strip() != grp_str]
                    df_s.to_excel(writer, sheet_name=sheet_name, index=False)
            borrados['excel_como_vamos'] = "Actualizado"
    except Exception as e_xl:
        borrados['excel_como_vamos'] = f"Nota: {e_xl}"

    return borrados

def vaciar_base_datos_completa(vaciar_usuarios=False, eliminar_archivos_excel=True):
    """
    Elimina todos los datos cargados de consultoras_tableau, metas_como_vamos, comentarios_lideres
    y los archivos Excel locales (Base para el como vamos.xlsx, Base de Datos.xlsx, etc.).
    Si vaciar_usuarios=True, resetea la lista de usuarios conservando solo las cuentas principales (superadmin/gerente).
    """
    conn = obtener_conexion_db()
    cursor = conn.cursor()

    res = {}
    try:
        cursor.execute("DELETE FROM consultoras_tableau")
        res['consultoras_tableau'] = cursor.rowcount
    except Exception:
        res['consultoras_tableau'] = 0

    try:
        cursor.execute("DELETE FROM metas_como_vamos")
        res['metas_como_vamos'] = cursor.rowcount
    except Exception:
        res['metas_como_vamos'] = 0

    try:
        cursor.execute("DELETE FROM comentarios_lideres")
        res['comentarios_lideres'] = cursor.rowcount
    except Exception:
        res['comentarios_lideres'] = 0

    if vaciar_usuarios:
        try:
            usuarios = cargar_usuarios()
            usuarios_filtrados = {u: info for u, info in usuarios.items() if info.get('rol') in ['superadmin', 'gerente']}
            guardar_usuarios(usuarios_filtrados)
            cursor.execute("DELETE FROM usuarios WHERE rol = 'lider'")
            res['usuarios_lideres_borrados'] = cursor.rowcount
        except Exception:
            res['usuarios_lideres_borrados'] = 0

    conn.commit()
    conn.close()

    # Eliminar archivos Excel locales para que no queden datos viejos en 'Cómo Vamos'
    if eliminar_archivos_excel:
        archivos_a_limpiar = [
            "Base para el como vamos.xlsx",
            "Base de Datos.xlsx",
            "Resultado_Metas_Procesadas.xlsx",
            "mi_grupo.xls",
            "activas.xlsx"
        ]
        borrados_archivos = []
        for arch in archivos_a_limpiar:
            if os.path.exists(arch):
                try:
                    os.remove(arch)
                    borrados_archivos.append(arch)
                except Exception:
                    pass
        res['archivos_excel_eliminados'] = borrados_archivos

    return res

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

