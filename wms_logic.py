"""
wms_logic.py
Lógica de cálculo del Cruce WMS (extraída de SCRIPT_WMS_V3_1.py),
separada en una función reutilizable por el backend Flask.

Esta primera entrega SOLO calcula. No genera Excel (eso es la Entrega 4).
"""
import pandas as pd


def fix_cols(df):
    df.columns = [c.encode('latin1').decode('utf-8', errors='replace')
                  if isinstance(c, str) else c for c in df.columns]
    return df


def sku_multiplier(sku):
    sku = str(sku).strip()
    if len(sku) == 15:
        try:
            return int(sku[10:12])
        except Exception:
            return 1
    return 1


def map_piso(row):
    area = str(row['AREA']).strip().upper()
    ubic = str(row['UBICACION']).strip().upper()
    if area == 'MZN04':
        return '4TO PISO'
    elif area == 'MZN03':
        return '3ER PISO'
    elif area == 'MZN02':
        return '2DO PISO'
    elif area in ('PISO', 'PARED', 'DIS', 'ACT', 'SIN'):
        return 'OMITIR'
    elif ubic.startswith('CDBUFFER-C'):
        return 'BUFFER-C'
    else:
        return '1ER PISO'


def area_asignada(row):
    etq = str(row.get('ETIQUETA', '')).strip().upper()
    if etq in ('INSUMOS', 'UNIFORMES'):
        return etq
    nivel = row.get('NIVEL_BEST', '')
    return nivel if pd.notna(nivel) and nivel != '' else 'SIN STOCK'


def puede_atender(row):
    disp = row['DISP_CONV_TOT']
    pend = row['PEND_CONVERTIDO']
    if pd.isna(disp) or disp == 0:
        return 'NO'
    if pend <= 0:
        return 'NO'
    if disp >= pend:
        return 'SÍ'
    return 'PARCIAL'


def procesar_wms(filepath):
    """
    Lee el Excel (hojas PEND, STOCK, RUTA Y PRIORIDAD) y devuelve:
      - df_full: DataFrame completo (línea por línea, pendiente vs stock)
      - stk_valid: DataFrame de stock válido (sin OMITIR), por si se necesita después
    Lanza ValueError con mensaje claro si falta alguna hoja.
    """
    try:
        xls = pd.ExcelFile(filepath)
    except Exception as e:
        raise ValueError(f"No se pudo abrir el archivo Excel: {e}")

    hojas_requeridas = ['PEND', 'STOCK', 'RUTA Y PRIORIDAD']
    faltantes = [h for h in hojas_requeridas if h not in xls.sheet_names]
    if faltantes:
        raise ValueError(
            f"Faltan hojas en el Excel: {', '.join(faltantes)}. "
            f"Hojas encontradas: {', '.join(xls.sheet_names)}"
        )

    df_pend_raw = pd.read_excel(filepath, sheet_name='PEND')
    df_stock_raw = pd.read_excel(filepath, sheet_name='STOCK')
    df_ruta_raw = pd.read_excel(filepath, sheet_name='RUTA Y PRIORIDAD')

    df_pend_raw = fix_cols(df_pend_raw)
    df_stock_raw = fix_cols(df_stock_raw)
    df_ruta_raw = fix_cols(df_ruta_raw)

    # ── PEND ──
    pend = pd.DataFrame()
    pend['N_ORDEN'] = df_pend_raw.iloc[:, 0]
    pend['ESTADO_ORDEN'] = df_pend_raw.iloc[:, 1]
    pend['SKU'] = df_pend_raw.iloc[:, 2].astype(str).str.strip()
    pend['CANT_SOL'] = pd.to_numeric(df_pend_raw.iloc[:, 3], errors='coerce').fillna(0)
    pend['CANT_ASIG'] = pd.to_numeric(df_pend_raw.iloc[:, 4], errors='coerce').fillna(0)
    pend['TDA_DESTINO'] = df_pend_raw.iloc[:, 5]
    pend['DESCRIPCION'] = df_pend_raw.iloc[:, 6]

    # ── STOCK ──
    stk = pd.DataFrame()
    stk['AREA'] = df_stock_raw.iloc[:, 0].astype(str).str.strip()
    stk['SKU'] = df_stock_raw.iloc[:, 1].astype(str).str.strip()
    stk['DESC_STK'] = df_stock_raw.iloc[:, 2]
    stk['UBICACION'] = df_stock_raw.iloc[:, 3].astype(str).str.strip()
    stk['CANT_ACTUAL'] = pd.to_numeric(df_stock_raw.iloc[:, 4], errors='coerce').fillna(0)
    stk['CANT_ASIG_STK'] = pd.to_numeric(df_stock_raw.iloc[:, 5], errors='coerce').fillna(0)

    # ── RUTA ──
    cols_ruta_requeridas = ['LIMA / PROV', 'ruta', 'tda', 'Orden', 'semana',
                             'DIA CORREO', 'NOMBR', 'ETIQUETA']
    faltan_ruta = [c for c in cols_ruta_requeridas if c not in df_ruta_raw.columns]
    if faltan_ruta:
        raise ValueError(
            f"Faltan columnas en la hoja RUTA Y PRIORIDAD: {', '.join(faltan_ruta)}"
        )

    ruta = pd.DataFrame()
    ruta['LIMA_PROV'] = df_ruta_raw['LIMA / PROV']
    ruta['RUTA'] = df_ruta_raw['ruta']
    ruta['TDA'] = df_ruta_raw['tda']
    ruta['ORDEN'] = df_ruta_raw['Orden']
    ruta['SEMANA'] = df_ruta_raw['semana'].astype(str)
    ruta['DIA_CORREO'] = df_ruta_raw['DIA CORREO'].astype(str)
    ruta['NOMBRE_TIENDA'] = df_ruta_raw['NOMBR']
    ruta['ETIQUETA'] = df_ruta_raw['ETIQUETA']

    # ── PENDIENTE ──
    pend['PENDIENTE'] = pend['CANT_SOL'] - pend['CANT_ASIG']
    pend['MULT'] = pend['SKU'].apply(sku_multiplier)
    pend['PEND_CONVERTIDO'] = pend['PENDIENTE'] * pend['MULT']

    # ── DISPONIBLE STOCK ──
    stk['DISPONIBLE'] = stk['CANT_ACTUAL'] - stk['CANT_ASIG_STK']
    stk['MULT'] = stk['SKU'].apply(sku_multiplier)
    stk['DISP_CONV'] = stk['DISPONIBLE'] * stk['MULT']

    # ── NIVELES / PISOS ──
    stk['NIVEL'] = stk.apply(map_piso, axis=1)

    PISO_ORDER = {'1ER PISO': 1, '2DO PISO': 2, '3ER PISO': 3, '4TO PISO': 4,
                  'BUFFER-C': 5, 'OMITIR': 9}
    stk['_SORT'] = stk['NIVEL'].map(PISO_ORDER).fillna(9)

    stk_valid = stk[stk['NIVEL'] != 'OMITIR'].copy()

    disp_conv_tot = stk_valid.groupby('SKU')['DISP_CONV'].sum().reset_index()
    disp_conv_tot.columns = ['SKU', 'DISP_CONV_TOT']

    nivel_disp = (stk_valid.groupby(['SKU', 'NIVEL'])['DISP_CONV']
                  .sum().reset_index().rename(columns={'DISP_CONV': 'DISP_NIVEL'}))
    if len(nivel_disp):
        idx_best = nivel_disp.groupby('SKU')['DISP_NIVEL'].idxmax()
        best_nivel = (nivel_disp.loc[idx_best, ['SKU', 'NIVEL']]
                      .rename(columns={'NIVEL': 'NIVEL_BEST'}))
    else:
        best_nivel = pd.DataFrame(columns=['SKU', 'NIVEL_BEST'])

    disp_tot = disp_conv_tot.merge(best_nivel, on='SKU', how='left')

    # ── JOIN PEND → RUTA ──
    ruta_u = ruta.drop_duplicates('ORDEN')[
        ['ORDEN', 'LIMA_PROV', 'RUTA', 'NOMBRE_TIENDA', 'SEMANA', 'DIA_CORREO', 'ETIQUETA']
    ].copy()

    pend['N_ORDEN'] = pd.to_numeric(pend['N_ORDEN'], errors='coerce')
    ruta_u['ORDEN'] = pd.to_numeric(ruta_u['ORDEN'], errors='coerce')

    df_full = pend.merge(ruta_u, left_on='N_ORDEN', right_on='ORDEN', how='left')
    df_full['SEMANA_DIA'] = df_full['SEMANA'].fillna('') + ' ' + df_full['DIA_CORREO'].fillna('')

    # ── JOIN → STOCK ──
    df_full = df_full.merge(disp_tot, on='SKU', how='left')

    # ── ÁREA ASIGNADA ──
    df_full['AREA_ASIGNADA'] = df_full.apply(area_asignada, axis=1)

    # ── PUEDE ATENDER ──
    df_full['PUEDE_ATENDER'] = df_full.apply(puede_atender, axis=1)

    return df_full, stk_valid


def resumen_kpis(df_full):
    """Resumen ligero para guardar en el historial (Entrega 1: solo lo esencial)."""
    total_lin = int(len(df_full))
    total_pconv = int(df_full['PEND_CONVERTIDO'].sum())
    n_si = int((df_full['PUEDE_ATENDER'] == 'SÍ').sum())
    n_parcial = int((df_full['PUEDE_ATENDER'] == 'PARCIAL').sum())
    n_no = int((df_full['PUEDE_ATENDER'] == 'NO').sum())
    return {
        "total_lineas": total_lin,
        "pend_conv_total": total_pconv,
        "si": n_si,
        "parcial": n_parcial,
        "no": n_no,
    }
