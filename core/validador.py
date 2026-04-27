"""Validación cruzada de KPIs entre texto Word y tabla Excel madre por compañía."""

import re
import time

import state
from config import CONFIG_COMPANIAS, CONFIG_HOJAS_ADICIONALES, CONFIG_CELDAS_DESVIACIONES, CONFIG_KPI_EXCLUIDOS, CONFIG_KPI_PREFIJOS_EXCLUIDOS, CONFIG_SUBSECCIONES_CONTEXTO

# ── Resultado estructurado (para panel HTML) ──────────────────────────────────
_resultados: list = []

def get_resultados():
    """Devuelve los resultados de la última validación como lista serializable."""
    return list(_resultados)

# ── Helpers numéricos ─────────────────────────────────────────────────────────

def _a_float(s):
    """
    Convierte string numérico a float manejando:
    - separador de miles con coma:  "450,123"   → 450123.0
    - decimal con coma:             "92,5"       → 92.5
    - decimal con punto:            "92.5"       → 92.5
    - miles + decimal punto:        "1,234,567.89" → 1234567.89
    """
    s = str(s).strip().replace(' ', '').lstrip("+")
    if re.match(r'^-?\d{1,3}(,\d{3})+$', s):
        s = s.replace(",", "")
    elif re.match(r'^-?\d{1,3}(,\d{3})+\.\d+$', s):
        s = s.replace(",", "")
    else:
        s = s.replace(",", ".")
    partes = s.split(".")
    if len(partes) > 2:
        s = "".join(partes[:-1]) + "." + partes[-1]
    try:
        return float(s)
    except ValueError:
        return None


def _normalizar_excel(valor):
    """
    Devuelve posibles floats de un valor de celda Excel.
    Para fracciones (|v| < 1), también incluye v*100 (porcentaje almacenado como fracción).
    """
    if valor is None or isinstance(valor, bool):
        return []
    if isinstance(valor, (int, float)):
        v = float(valor)
        candidatos = [round(v, 4)]
        if 0 < abs(v) < 10:
            candidatos.append(round(v * 100, 4))
        return candidatos
    if isinstance(valor, str):
        f = _a_float(re.sub(r"[^\d.,+\-]", "", valor))
        if f is not None:
            return [round(f, 4)]
    return []


def _tol_para(raw_str):
    """
    Tolerancia basada en la precisión con que el Word escribe el número:
      - sin decimal (entero ej. 238)      → ±0.6
      - miles con coma (ej. 4,506 = 4506) → ±0.6  (la coma NO es decimal)
      - 1 decimal    (ej. 7.4)            → ±0.06
      - 2 decimales  (ej. 1.08)           → ±0.006
    """
    s = raw_str.strip().lstrip("+").lstrip("-")
    # Coma como separador de miles: N,NNN  o  N,NNN,NNN → tolerancia de entero
    if re.match(r'^\d{1,3}(,\d{3})+$', s):
        return 0.6
    m = re.search(r'[.,](\d+)', raw_str)
    if m:
        return 0.6 / (10 ** len(m.group(1)))
    return 0.6


def _encontrar_en_fila(v_abs, nums_fila, tol=None, signed_val=None):
    """
    Busca v_abs en nums_fila con tolerancia.
    - Modo normal (KPIs): compara en valor absoluto en ambos lados, porque
      el Excel puede almacenar negativos para desviaciones negativas.
    - Modo con signo (acumulados): recibe signed_val y compara directamente
      signed_val vs el valor del Excel, para detectar discrepancias de signo.
    Devuelve (encontrado: bool, valor_más_cercano: float|None).
    """
    if not nums_fila:
        return False, None
    if tol is None:
        tol = 0.6
    if signed_val is not None:
        cercano = min(nums_fila, key=lambda v: abs(v - signed_val))
        return abs(cercano - signed_val) <= tol, round(cercano, 4)
    cercano = min(nums_fila, key=lambda v: abs(abs(v) - v_abs))
    return abs(abs(cercano) - v_abs) <= tol, round(cercano, 4)


# ── Extracción de números desde líneas ────────────────────────────────────────

_PAT_NUMERO = re.compile(r'[+\-] ?\d[\d.,]*|[+\-]?\d[\d.,]*')

def _numeros_de_linea(linea):
    """
    Extrae todos los valores numéricos distintos de una línea.
    Devuelve lista de (raw_str, float_abs_redondeado).
    Descarta: cero, años (1900-2100), y duplicados en valor absoluto.
    """
    # limpiar_texto_global inserta NBSP (U+00A0) entre el signo y el dígito para
    # evitar saltos de línea en Word (ej. "-3.6" → " - 3.6").
    # Si no se elimina, _PAT_NUMERO no detecta el signo y el número queda positivo.
    linea = linea.replace(' ', '')
    linea = linea.replace(' ', '')
    vistos = set()
    resultado = []
    for m in _PAT_NUMERO.finditer(linea):
        raw = m.group(0)
        v = _a_float(raw)
        if v is None or v == 0:
            continue
        v_abs = round(abs(v), 4)
        if 1900 <= v_abs <= 2100 and v_abs == int(v_abs):
            continue
        if v_abs not in vistos:
            vistos.add(v_abs)
            resultado.append((raw, v_abs))
    return resultado


# ── Normalización de texto para matching ──────────────────────────────────────

_TILDES = str.maketrans(
    'áéíóúàèìòùäëïöüñÁÉÍÓÚÀÈÌÒÙÄËÏÖÜÑ',
    'aeiouaeiouaeiounAEIOUAEIOUAEIOUN'
)

def _norm(texto):
    """Normaliza etiqueta: minúsculas, sin tildes, sin espacios dobles."""
    return re.sub(r'\s+', ' ', texto.translate(_TILDES).lower().strip())


# ── Extracción de etiqueta desde línea Word ───────────────────────────────────

_BULLETS = set('·•‣▸▹►-–—*')

def _extraer_label(linea):
    """
    Extrae la parte de texto (etiqueta KPI) al inicio de una línea, antes del primer número.
    Devuelve None si:
      - La línea empieza con viñeta o símbolo
      - La etiqueta resultante es muy corta (< 3 chars)
    """
    if not linea:
        return None
    if linea[0] in _BULLETS:
        return None
    m = re.search(r'\d', linea)
    if m and m.start() > 3:
        label = re.sub(r'[\s:(;\-+]+$', '', linea[:m.start()]).strip()
        if len(label) >= 3:
            return label
    if not m and len(linea.strip()) >= 3:
        return linea.strip()
    return None


# ── Lectura de Excel con etiquetas por fila ───────────────────────────────────

def _extender_a_col_a(rango):
    """Si el rango no empieza en columna A, lo extiende para incluirla (etiquetas)."""
    m = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', rango)
    if not m:
        return rango
    col_ini, row_ini, col_fin, row_fin = m.groups()
    if col_ini == 'A':
        return rango
    return f"A{row_ini}:{col_fin}{row_fin}"


def _leer_excel_por_etiqueta(wb_com, nombre_hoja, rango):
    """
    Lee el rango fila a fila desde COM.
    Para cada fila, la primera celda de texto no vacía es la etiqueta;
    el resto de celdas son los valores numéricos de esa fila.

    Devuelve ({norm_label: (label_orig, set_floats)}, error_str|None).
    Extiende el rango automáticamente para incluir la columna A si es necesario.
    """
    rango_ext = _extender_a_col_a(rango)
    try:
        ws = wb_com.Worksheets(nombre_hoja)
        valores_raw = ws.Range(rango_ext).Value
    except Exception as e:
        return None, str(e)

    tabla = {}
    if not valores_raw:
        return tabla, None

    for fila in valores_raw:
        if not fila:
            continue
        label_orig = None
        nums = set()
        for celda in fila:
            if label_orig is None:
                if isinstance(celda, str) and celda.strip():
                    # Descartar separadores (filas de encabezado visual)
                    limpia = celda.strip().strip('-').strip('=').strip('_').strip()
                    if limpia:
                        label_orig = celda.strip()
            else:
                for n in _normalizar_excel(celda):
                    nums.add(n)
        if label_orig and nums:
            tabla[_norm(label_orig)] = (label_orig, nums, None)

    return tabla, None


# ── Lectura por celdas exactas (CONFIG_CELDAS_DESVIACIONES) ──────────────────

_RPC_REJECTED = -2147418111  # RPC_E_CALL_REJECTED: Excel ocupado, reintentable

def _com_call(fn, reintentos=5, pausa=1.5):
    """Ejecuta fn() reintentando si Excel rechaza la llamada COM por estar ocupado.
    pywintypes.com_error guarda el HRESULT en args[0], no en el atributo .hresult."""
    for intento in range(reintentos):
        try:
            return fn()
        except Exception as e:
            args = getattr(e, 'args', ())
            hresult = args[0] if args and isinstance(args[0], int) else getattr(e, 'hresult', None)
            if intento < reintentos - 1 and hresult == _RPC_REJECTED:
                time.sleep(pausa)
                continue
            raise


def _leer_celdas_exactas(wb_com, nombre_hoja, celdas_config):
    """
    Lee las celdas exactas definidas en CONFIG_CELDAS_DESVIACIONES.
    Cada KPI tiene (celda_dif, celda_pct) o (celda_dif, celda_pct, celda_status).
    Devuelve ({norm_label: (label_orig, set_floats, status_or_None)}, error_str|None).
    Auto-detecta el estado escaneando la fila si no se configura celda_status.
    """
    try:
        ws = _com_call(lambda: wb_com.Worksheets(nombre_hoja))
    except Exception as e:
        return None, str(e)

    _cache_filas = {}  # row_num → status_text normalizado

    tabla = {}
    for label, cells in celdas_config.items():
        celda_dif = cells[0]
        celda_pct = cells[1] if len(cells) > 1 else None
        celda_status_cfg = cells[2] if len(cells) > 2 else None

        nums = set()
        for celda_a1 in filter(None, (celda_dif, celda_pct)):
            try:
                v = _com_call(lambda c=celda_a1: ws.Range(c).Value)
                for n in _normalizar_excel(v):
                    nums.add(n)
            except Exception:
                pass

        status_excel = None
        if celda_status_cfg:
            try:
                v = _com_call(lambda c=celda_status_cfg: ws.Range(c).Value)
                if isinstance(v, str) and _PAT_STATUS.search(v):
                    status_excel = _norm(v.strip())
            except Exception:
                pass
        elif celda_dif:
            row_match = re.search(r'\d+', celda_dif)
            if row_match:
                row_num = row_match.group()
                if row_num not in _cache_filas:
                    try:
                        rng = _com_call(lambda r=row_num: ws.Range(f"A{r}:Z{r}").Value)
                        found = None
                        if rng:
                            fila_vals = rng[0] if (isinstance(rng, (tuple, list))
                                                   and isinstance(rng[0], (tuple, list))) else rng
                            for cell_val in fila_vals:
                                if isinstance(cell_val, str) and _PAT_STATUS.search(cell_val):
                                    found = _norm(cell_val.strip())
                                    break
                        _cache_filas[row_num] = found
                    except Exception:
                        _cache_filas[row_num] = None
                status_excel = _cache_filas[row_num]

        if nums:
            tabla[_norm(label)] = (label, nums, status_excel)

    return tabla, None


# ── Lectura dinámica de la sección Principales Desviaciones ───────────────────

def _leer_desviaciones_dinamico(wb_com, nombre_hoja, rango_explicito=None):
    """
    Lee la sección 'Principales Desviaciones' de la hoja y devuelve
    ({norm_label: (label_orig, set_floats)}, error_str|None).

    - Si rango_explicito es dado (ej. "C37:J58"), lee ese rango tal cual,
      SIN extender a la col A, para que el primer texto de cada fila sea
      la etiqueta del KPI (no el nombre de sección de columnas previas).
    - Si es None, busca dinámicamente la cabecera 'Principales Desviaciones'
      e identifica las columnas Dif y Var% por sus encabezados.
    """
    try:
        ws = wb_com.Worksheets(nombre_hoja)
        used = ws.UsedRange
        max_r = used.Row + used.Rows.Count - 1
        max_c = used.Column + used.Columns.Count - 1
    except Exception as e:
        return None, str(e)

    # ── Camino rápido: rango explícito de config ──────────────────────────────
    if rango_explicito:
        try:
            valores_raw = ws.Range(rango_explicito).Value
        except Exception as e:
            return None, str(e)

        if not valores_raw:
            return {}, None

        # 1. Buscar índices de columna "Dif", "Var" y estado en las primeras 4 filas
        idx_dif = None
        idx_pct = None
        idx_status = None
        _PAT_STATUS_HDR = re.compile(
            r'(condici[oó]n|estado|vs[\s_]pm|evaluaci[oó]n)', re.IGNORECASE
        )
        for fila in valores_raw[:4]:
            if not fila:
                continue
            for i, celda in enumerate(fila):
                if celda and isinstance(celda, str):
                    cs = celda.strip().lower()
                    if idx_dif is None and cs.startswith("dif"):
                        idx_dif = i
                    if idx_pct is None and ("var" in cs or cs == "%"):
                        idx_pct = i
                    if idx_status is None and _PAT_STATUS_HDR.search(cs):
                        idx_status = i
            if idx_dif is not None and idx_pct is not None:
                break

        # Si no se halló columna de estado por encabezado, detectar por contenido
        if idx_status is None:
            for fila in valores_raw[4:12]:
                if not fila:
                    continue
                for i, celda in enumerate(fila):
                    if isinstance(celda, str) and _PAT_STATUS.search(celda):
                        idx_status = i
                        break
                if idx_status is not None:
                    break

        # 2. Leer filas de datos leyendo SÓLO las columnas Dif, Var% y estado
        tabla = {}
        for fila in valores_raw:
            if not fila:
                continue
            # Primera celda de texto = etiqueta del KPI
            label_orig = None
            for celda in fila:
                if isinstance(celda, str) and celda.strip():
                    limpia = celda.strip().strip('-').strip('=').strip('_').strip()
                    if limpia:
                        label_orig = celda.strip()
                        break
            if not label_orig:
                continue

            nums = set()
            cols = [c for c in (idx_dif, idx_pct) if c is not None]
            for idx in cols:
                if idx < len(fila):
                    for n in _normalizar_excel(fila[idx]):
                        nums.add(n)

            status_excel = None
            if idx_status is not None and idx_status < len(fila):
                cell_s = fila[idx_status]
                if isinstance(cell_s, str) and _PAT_STATUS.search(cell_s):
                    status_excel = _norm(cell_s.strip())

            if nums:
                tabla[_norm(label_orig)] = (label_orig, nums, status_excel)

        return tabla, None

    # ── Búsqueda dinámica (fallback para compañías sin rango_desviaciones) ────
    fila_inicio = None
    col_label = None
    for r in range(1, max_r + 1):
        for c in range(1, min(max_c + 1, 20)):
            v = ws.Cells(r, c).Value
            if v and "principales desviaciones" in str(v).strip().lower():
                fila_inicio = r
                col_label = c
                break
        if fila_inicio:
            break

    if not fila_inicio:
        return {}, None

    # Identificar columnas Dif, Var% y estado por encabezado
    col_dif = None
    col_pct = None
    col_status = None
    _PAT_STATUS_HDR2 = re.compile(
        r'(condici[oó]n|estado|vs[\s_]pm|evaluaci[oó]n)', re.IGNORECASE
    )
    for r in range(fila_inicio, min(fila_inicio + 4, max_r + 1)):
        for c in range(col_label, min(col_label + 14, max_c + 1)):
            v = ws.Cells(r, c).Value
            if v:
                vs = str(v).strip().lower()
                if col_dif is None and vs.startswith("dif"):
                    col_dif = c
                if col_pct is None and ("var" in vs or vs == "%"):
                    col_pct = c
                if col_status is None and _PAT_STATUS_HDR2.search(vs):
                    col_status = c
        if col_dif and col_pct:
            break

    if not col_dif:
        col_dif = col_label + 2
    if not col_pct:
        col_pct = col_label + 4

    tabla = {}
    filas_vacias = 0
    for r in range(fila_inicio + 1, min(fila_inicio + 40, max_r + 1)):
        v_label = ws.Cells(r, col_label).Value
        if not v_label or not isinstance(v_label, str) or not v_label.strip():
            filas_vacias += 1
            if filas_vacias >= 3:
                break
            continue
        filas_vacias = 0

        label_norm = _norm(v_label.strip())
        if not label_norm or len(label_norm) < 2:
            continue

        nums = set()
        for col in (col_dif, col_pct):
            for n in _normalizar_excel(ws.Cells(r, col).Value):
                nums.add(n)

        status_excel = None
        if col_status:
            cell_s = ws.Cells(r, col_status).Value
            if isinstance(cell_s, str) and _PAT_STATUS.search(cell_s):
                status_excel = _norm(cell_s.strip())

        if nums:
            tabla[label_norm] = (v_label.strip(), nums, status_excel)

    return tabla, None


# ── Búsqueda dinámica de celdas Acumulado al mes/año ─────────────────────────

_PAT_PCT_TEXTO = re.compile(r'[+\-]?\d+(?:[.,]\d+)?%')

def _agregar_acumulados_desde_excel(wb_com, nombre_hoja, tabla):
    """
    Lee la hoja completa (UsedRange) en una sola llamada COM, busca las celdas
    de texto que empiezan con 'Acumulado al mes' y 'Acumulado al año', extrae
    los porcentajes y los agrega a la tabla con claves normalizadas.
    """
    try:
        ws     = wb_com.Worksheets(nombre_hoja)
        valores = ws.UsedRange.Value
    except Exception:
        return
    if not valores:
        return

    pendientes = {"acumulado al mes": None, "acumulado al ano": None}

    for fila in valores:
        if not fila:
            continue
        for celda in fila:
            if not isinstance(celda, str):
                continue
            vn = _norm(celda)
            for clave in list(pendientes.keys()):
                if pendientes[clave] is not None:
                    continue
                if vn.startswith(clave):
                    # Extraer todos los números en orden (con signo), igual que en Word
                    nums_lista = []
                    seen_abs = set()
                    for raw, v_abs in _numeros_de_linea(celda):
                        v_signed = -v_abs if raw.lstrip().startswith('-') else v_abs
                        if v_abs not in seen_abs:
                            seen_abs.add(v_abs)
                            nums_lista.append(v_signed)
                    if nums_lista:
                        tabla[clave] = (celda.strip()[:80], nums_lista, None)
                        pendientes[clave] = True
        if all(v is not None for v in pendientes.values()):
            break


# ── Búsqueda de fila Excel por nombre ─────────────────────────────────────────

def _buscar_fila(label_word_norm, tabla):
    """
    Busca la fila Excel cuya etiqueta mejor corresponde a la etiqueta del Word.
    Estrategia (en orden de prioridad):
      1. Coincidencia exacta (normalizada)
      2. Substring: una etiqueta está contenida en la otra
      3. Palabras: mismo conjunto de palabras (orden distinto) o una es
         subconjunto de la otra — maneja casos como
         'Cu Fino Filtrado Pagable' ↔ 'Cu Fino Pagable Filtrado'
         y 'Molibdeno' ↔ 'Molibdeno fino pagable'
    En caso de múltiples candidatos se prefiere el de mayor solapamiento
    y, en empate, el label Excel más largo (más específico).
    Devuelve (label_orig_excel, set_floats, status_or_None) o (None, None, None).
    """
    if label_word_norm in tabla:
        return tabla[label_word_norm]

    # ── Estrategia 2: substring ───────────────────────────────────────────────
    candidatos = [
        (excel_norm, orig, nums, status)
        for excel_norm, (orig, nums, status) in tabla.items()
        if label_word_norm in excel_norm or excel_norm in label_word_norm
    ]
    if candidatos:
        candidatos.sort(key=lambda x: len(x[0]), reverse=True)
        return candidatos[0][1], candidatos[0][2], candidatos[0][3]

    # ── Estrategia 3: solapamiento por conjunto de palabras ───────────────────
    words_w = set(label_word_norm.split())
    candidatos_ws = []
    for excel_norm, (orig, nums, status) in tabla.items():
        words_e = set(excel_norm.split())
        intersec = words_w & words_e
        union    = words_w | words_e
        if not intersec:
            continue
        ratio = len(intersec) / len(union)
        # Acepta: mismo conjunto (reordenado) o uno es subconjunto del otro
        if words_w == words_e or words_w <= words_e or words_e <= words_w:
            candidatos_ws.append((ratio, len(excel_norm), orig, nums, status))

    if not candidatos_ws:
        return None, None, None

    candidatos_ws.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidatos_ws[0][2], candidatos_ws[0][3], candidatos_ws[0][4]


# ── Captura de líneas de producción ──────────────────────────────────────────

_SECCIONES_PRODUCCION = [
    "Principales Desviaciones", "Mina", "Planta", "Concentradora",
    "Cátodos", "Sulfuros", "Tren", "Camión", "Detalle por fases",
    "Planta Desaladora",
]
_FIN_SECCIONES = ["Accidentabilidad", "Reportabilidad", "Medio Ambiente", "Asuntos Públicos"]


def _capturar_lineas(texto, secciones_inicio, secciones_fin):
    """Devuelve las líneas que caen entre las secciones de inicio y las de fin.

    Los subtítulos de contexto definidos en CONFIG_SUBSECCIONES_CONTEXTO (ej.
    'Planta Hidro MET') se preservan como líneas aunque empiecen con una palabra
    que coincida con un marcador de sección (ej. 'Planta').
    """
    _subsec_headers = {
        _norm(h)
        for comp_map in CONFIG_SUBSECCIONES_CONTEXTO.values()
        for h in comp_map
    }
    lineas = []
    capturar = False
    for linea in texto.split("\n"):
        l = linea.strip()
        if not l:
            continue
        # Subtítulos de contexto: preservarlos como líneas, nunca como marcadores
        if _norm(l) in _subsec_headers:
            if capturar:
                lineas.append(l)
            continue
        if any(l.startswith(s) for s in secciones_inicio):
            capturar = True
            continue
        if any(l.startswith(f) for f in secciones_fin):
            capturar = False
            continue
        if capturar:
            lineas.append(l)
    return lineas


def _capturar_lineas_seccion(texto, nombre_seccion):
    """Devuelve líneas de una sección específica (para hojas adicionales)."""
    todas = _SECCIONES_PRODUCCION + list(CONFIG_HOJAS_ADICIONALES.keys())
    fin = _FIN_SECCIONES + [s for s in todas if s != nombre_seccion]
    lineas = []
    capturar = False
    for linea in texto.split("\n"):
        l = linea.strip()
        if not l:
            continue
        if l.startswith(nombre_seccion):
            capturar = True
            continue
        if capturar and any(l.startswith(f) for f in fin):
            break
        if capturar:
            lineas.append(l)
    return lineas


# ── Truncar línea en el indicador de estado ───────────────────────────────────

_PAT_STATUS = re.compile(
    r'\b(sobre\s+pm|bajo\s+pm|en\s+l[ií]nea)\b',
    re.IGNORECASE
)

def _truncar_en_status(linea):
    """Devuelve solo la parte antes de 'sobre PM / bajo PM / en línea'."""
    m = _PAT_STATUS.search(linea)
    return linea[:m.start()] if m else linea


def _extraer_status_word(linea):
    """Extrae el estado Word normalizado ('bajo pm', 'sobre pm', 'en linea') o None."""
    m = _PAT_STATUS.search(linea)
    return _norm(m.group(0)) if m else None


# ── Comparar y reportar ───────────────────────────────────────────────────────

def _comparar_y_reportar(clave, label_sec, lineas, tabla_excel):
    """
    Para cada línea de producción:
      1. Extrae la etiqueta KPI (texto antes del primer número).
      2. Busca la celda Excel correspondiente por nombre.
      3. Compara SÓLO los números de la parte de desviación (antes de
         'sobre PM / bajo PM / en línea') con los valores de Excel,
         usando tolerancia adaptativa según la precisión escrita en el Word.
    Además recolecta resultados estructurados en _resultados (para panel HTML).
    """
    print(f"\n  {'─'*64}")
    print(f"  {clave} — {label_sec}")
    print(f"  {'─'*64}")

    n_warn = 0
    n_sin_fila = 0
    lineas_revisadas = 0
    _kpis = []   # resultados estructurados de esta sección
    contexto_suffix = None   # sufijo de subsección activo (ej. "MET", "OXE")
    _subsecciones = {_norm(k): v for k, v in CONFIG_SUBSECCIONES_CONTEXTO.get(clave, {}).items()}
    excluidos         = {_norm(e) for e in CONFIG_KPI_EXCLUIDOS.get(clave, set())}
    prefijos_excluidos = {_norm(p) for p in CONFIG_KPI_PREFIJOS_EXCLUIDOS.get(clave, set())}

    for linea in lineas:
        # Detectar subtítulos de subsección (ej. "Planta Hidro MET")
        linea_norm = _norm(linea.strip())
        if linea_norm in _subsecciones:
            contexto_suffix = _subsecciones[linea_norm]
            print(f"\n    [contexto: {contexto_suffix}]")
            continue

        # Extraer números solo del segmento de desviación (antes del status)
        linea_dev = _truncar_en_status(linea)
        linea_dev = re.sub(r' ', '', linea_dev)
        numeros = _numeros_de_linea(linea_dev)
        if not numeros:
            continue
        lineas_revisadas += 1
        linea_corta = linea if len(linea) <= 78 else linea[:75] + "..."
        print(f"\n    {linea_corta}")

        label_word = _extraer_label(linea)
        kpi = {"linea": linea_corta, "label": label_word,
               "excel_label": None, "estado": "ok", "valores": []}

        if label_word:
            label_norm = _norm(label_word)
            # Verificar si el KPI está excluido (por substring o por prefijo)
            if any(e in label_norm for e in excluidos) or \
               any(label_norm.startswith(p) for p in prefijos_excluidos):
                print(f"      ↳ Excluido de validación")
                continue
            if re.match(r'^acumulado al mes', label_norm):
                label_norm = "acumulado al mes"
            elif re.match(r'^acumulado al an', label_norm):
                label_norm = "acumulado al ano"
            # Buscar con sufijo de contexto primero, luego sin él
            excel_label, nums_fila, status_excel = None, None, None
            if contexto_suffix:
                excel_label, nums_fila, status_excel = _buscar_fila(
                    label_norm + " " + contexto_suffix.lower(), tabla_excel
                )
            if not excel_label:
                excel_label, nums_fila, status_excel = _buscar_fila(label_norm, tabla_excel)
            es_acumulado = label_norm in ("acumulado al mes", "acumulado al ano")
        else:
            excel_label, nums_fila, status_excel = None, None, None
            es_acumulado = False

        if excel_label:
            kpi["excel_label"] = excel_label
            print(f"      ↳ Celda Excel: '{excel_label}'")
            if es_acumulado and isinstance(nums_fila, list):
                # Comparación posicional: Word[i] ↔ Excel[i] en el mismo orden.
                # Se compara en valor absoluto: el signo lo da el texto cualitativo
                # ("mayor/menor producción"), no los números en sí.
                for i, (raw, v_abs_w) in enumerate(numeros):
                    tol = _tol_para(raw)
                    if i < len(nums_fila):
                        v_abs_e = abs(nums_fila[i])
                        ok = abs(v_abs_w - v_abs_e) <= tol
                        cercano = round(nums_fila[i], 4)
                    else:
                        ok = False
                        cercano = None
                    marca = "✓" if ok else "✗"
                    dif = round(abs(v_abs_w - abs(cercano)), 4) if not ok and cercano is not None else None
                    if ok:
                        print(f"      {raw:>14}  →  {marca}  Excel = {cercano}")
                    else:
                        print(f"      {raw:>14}  →  {marca}  Excel = {cercano}  (dif: {dif})")
                        n_warn += 1
                        kpi["estado"] = "revisar"
                    kpi["valores"].append({"word": raw, "excel": cercano, "ok": ok, "dif": dif})
            else:
                for raw, v_abs in numeros:
                    tol = _tol_para(raw)
                    ok, cercano = _encontrar_en_fila(v_abs, nums_fila, tol=tol)
                    marca = "✓" if ok else "✗"
                    if ok:
                        print(f"      {raw:>14}  →  {marca}  Excel = {cercano}")
                    else:
                        print(f"      {raw:>14}  →  {marca}  Excel = {cercano}  (dif: {abs(v_abs - cercano):.4f})")
                        n_warn += 1
                        kpi["estado"] = "revisar"
                    kpi["valores"].append({
                        "word": raw, "excel": cercano, "ok": ok,
                        "dif": round(abs(v_abs - abs(cercano)), 4) if not ok and cercano is not None else None,
                    })

            # Validar indicador de estado (bajo PM / sobre PM / en línea)
            status_word = _extraer_status_word(linea)
            if status_word and status_excel:
                status_ok = status_word == status_excel
                marca_s = "✓" if status_ok else "✗"
                if status_ok:
                    print(f"      {'estado':>14}  →  {marca_s}  '{status_word}'")
                else:
                    print(f"      {'estado':>14}  →  {marca_s}  Word='{status_word}'  Excel='{status_excel}'")
                    n_warn += 1
                    kpi["estado"] = "revisar"
                kpi["status"] = {"word": status_word, "excel": status_excel, "ok": status_ok}
            elif status_word and not status_excel:
                print(f"      {'estado':>14}  →  ?  Word='{status_word}' (sin estado en Excel)")
                kpi["status"] = {"word": status_word, "excel": None, "ok": None}
        elif label_word:
            print(f"      ⚠ Sin celda Excel para: '{label_word}'")
            n_sin_fila += 1
            kpi["estado"] = "sin_celda"
        else:
            print(f"      (línea de viñeta — sin etiqueta para match)")
            kpi["estado"] = "sin_label"

        _kpis.append(kpi)

    print()
    if lineas_revisadas == 0:
        print(f"  - Sin líneas con valores numéricos")
        _resultados.append({"clave": clave, "nombre": label_sec, "kpis": [],
                             "n_ok": 0, "n_warn": 0, "n_sin_celda": 0, "estado": "ok", "error": None})
        return 1, 0

    n_ok = sum(1 for k in _kpis if k["estado"] == "ok")
    _resultados.append({
        "clave":      clave,
        "nombre":     label_sec,
        "kpis":       _kpis,
        "n_ok":       n_ok,
        "n_warn":     n_warn,
        "n_sin_celda": n_sin_fila,
        "estado":     "warn" if n_warn > 0 else "sin_celda" if n_sin_fila > 0 else "ok",
        "error":      None,
    })

    if n_warn == 0 and n_sin_fila == 0:
        print(f"  ✓ Sin diferencias  ({lineas_revisadas} línea(s) revisada(s))")
        return 1, 0
    if n_warn > 0:
        print(f"  ! {n_warn} valor(es) con diferencias, {n_sin_fila} etiqueta(s) sin fila Excel")
        state.errores.append(f"[REVISAR] {clave}: {n_warn} diferencia(s) en validación KPI")
    else:
        print(f"  ! {n_sin_fila} etiqueta(s) sin fila Excel correspondiente — revisar nombres")
    return 0, n_warn


# ── Validador principal ───────────────────────────────────────────────────────

def validar_kpis_vs_excel(informes, wb_com):
    """
    Compara los valores numéricos de las secciones de producción del Word
    contra los valores del rango correspondiente en el Excel madre,
    usando matching por nombre de KPI (etiqueta de fila).

    informes : dict  {clave: texto_word}
    wb_com   : workbook COM ya abierto
    """
    global _resultados
    _resultados = []
    print("\n══ Validación KPIs Word vs Excel ═══════════════════════════════")
    total_ok = 0
    total_warn = 0

    # 1. Validación por compañía (hoja = clave)
    for clave, cfg in CONFIG_COMPANIAS.items():
        texto = informes.get(clave, "")
        if not texto:
            continue

        lineas = _capturar_lineas(texto, _SECCIONES_PRODUCCION, _FIN_SECCIONES)
        if not lineas:
            print(f"\n  {clave}: sin líneas de producción detectadas (omitido)")
            continue

        celdas_cfg = CONFIG_CELDAS_DESVIACIONES.get(clave)
        if celdas_cfg:
            tabla_excel, err = _leer_celdas_exactas(wb_com, clave, celdas_cfg)
        else:
            tabla_excel, err = _leer_desviaciones_dinamico(wb_com, clave, cfg.get("rango_desviaciones"))

        if tabla_excel is None:
            es_rpc = _RPC_REJECTED in str(err)
            aviso = " — Excel ocupado o bloqueado: cierra ventanas de Excel abiertas y reintenta." if es_rpc else ""
            print(f"\n  ! {clave}: no se pudo leer hoja '{clave}'{aviso}\n    ({err})")
            state.errores.append(f"[REVISAR] Validación {clave}: {err}")
            total_warn += 1
            continue
        if not tabla_excel:
            print(f"\n  ! {clave}: sin KPIs configurados o sección no encontrada en hoja '{clave}'")
            total_warn += 1
            continue

        _agregar_acumulados_desde_excel(wb_com, clave, tabla_excel)

        ok, warn = _comparar_y_reportar(clave, cfg["nombre"], lineas, tabla_excel)
        total_ok += ok
        total_warn += warn

    # 2. Hojas adicionales (ej: Gestión Hídrica)
    for nombre_seccion, cfg_ad in CONFIG_HOJAS_ADICIONALES.items():
        texto = informes.get(cfg_ad["compania_fuente"], "")
        if not texto:
            continue

        lineas = _capturar_lineas_seccion(texto, nombre_seccion)
        if not lineas:
            print(f"\n  {nombre_seccion}: sin líneas detectadas (omitido)")
            continue

        tabla_excel, err = _leer_excel_por_etiqueta(wb_com, cfg_ad["hoja"], cfg_ad["rango"])
        if tabla_excel is None:
            print(f"\n  ! {nombre_seccion}: no se pudo leer hoja '{cfg_ad['hoja']}' ({err})")
            state.errores.append(f"[REVISAR] Validación {nombre_seccion}: {err}")
            total_warn += 1
            continue
        if not tabla_excel:
            print(f"\n  ! {nombre_seccion}: hoja leída pero sin filas con etiqueta — verificar rango")
            total_warn += 1
            continue

        ok, warn = _comparar_y_reportar(nombre_seccion, cfg_ad["hoja"], lineas, tabla_excel)
        total_ok += ok
        total_warn += warn

    print(f"\n{'═'*65}")
    if total_warn == 0:
        print(f"  ✓ Validación completa: todas las secciones OK")
    else:
        print(f"  ! Validación completa: {total_warn} valor(es) con diferencias — revisar antes de publicar")
    print(f"{'═'*65}\n")
