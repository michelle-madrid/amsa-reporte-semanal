"""
Revisor gramatical para informes por compañía.

Detecta:
  - Concordancia número adjetivo-sustantivo  (ej: "mejores rendimiento")
  - Concordancia género artículo-sustantivo  (ej: "el situación")
  - Frases o palabras duplicadas consecutivas
  - Errores reportados por LanguageTool API en español

Uso:
    python revisar_gramatica.py [ruta_al_docx]
    Sin argumento → abre selector de archivo.

Requisitos:
    pip install requests python-docx
"""

import sys
import re
import time
from pathlib import Path

import requests


# ── Selector de archivo ──────────────────────────────────────────────────────
def _seleccionar_archivo():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        ruta = filedialog.askopenfilename(
            title="Selecciona el informe Word a revisar",
            filetypes=[("Documentos Word", "*.docx"), ("Todos", "*.*")],
        )
        root.destroy()
        return ruta or None
    except Exception:
        return None


# ── Extracción de párrafos desde Word ───────────────────────────────────────
def extraer_parrafos(ruta_docx):
    from docx import Document
    doc = Document(ruta_docx)
    return [(p.text.strip(), p.style.name) for p in doc.paragraphs if p.text.strip()]


# ── Detección de compañía ────────────────────────────────────────────────────
COMPANIAS = {
    "Los Pelambres": "MLP",
    "Centinela":     "CEN",
    "Antucoya":      "ANT",
    "Zaldívar":      "CMZ",
    "FCAB":          "FCAB",
}

def detectar_compania(texto):
    for nombre, clave in COMPANIAS.items():
        if nombre.lower() in texto.lower():
            return f"{clave} – {nombre}"
    return None


# ════════════════════════════════════════════════════════════════════════════
# VOCABULARIO DE DOMINIO — nunca se reportan como error
# ════════════════════════════════════════════════════════════════════════════
PALABRAS_DOMINIO = {
    # Nombres propios / lugares
    "pelambres", "cuncumén", "cuncumen", "quelén", "quelen",
    "antucoya", "centinela", "zaldívar", "zaldivar", "fcab",
    "antofagasta", "atacama",
    # Proyectos y siglas operacionales
    "evu", "amsa", "amsm",
    "ocas", "oca",
    "siad", "str", "stc",
    # Unidades y abreviaciones técnicas (en minúscula para comparación)
    "kt", "ktm", "ktpd", "mt", "mtpd",
    "pm", "pa", "uebd", "ueb",
    "tmf", "tcu", "tmo", "tpd",
    "lps",
    # Términos mineros válidos
    "remanejo", "depositación", "deposición",
    "desaladora", "batimetría", "batimetria",
    "tranque", "ripios", "estéril", "esteril",
    "lastre", "oversize", "undersize",
    "chancado", "molienda", "flotación", "lixiviación",
    "pls", "mineralización", "yacimiento", "yacimientos",
    "sulfuros", "óxidos", "cátodo", "cátodos",
    "electroobtención", "electroextracción",
    # Términos técnicos inglés de uso corriente en minería
    "stockpile", "stock", "pile", "built", "as",   # "planos As Built", "stock pile"
    "overcut", "undercut", "drawpoint",
    # Formatos de archivo / extensiones usadas en el rubro
    "kmz", "kml", "dwg", "shp",
    # Adjetivos/sustantivos válidos que LT confunde con verbos
    "protocolares", "protocolar",
    # Otros que el texto puede incluir
    "faena", "faenas", "mampuesto", "resultados",
}

# Patrón para códigos de fase: F9SE, F10NW, F11, FASE2, etc.
_RE_CODIGO_FASE   = re.compile(r'\b[A-Z]\d+[A-Za-z]*\b')
# Abreviaciones todo-mayúscula de 2-5 letras (PM, EVU, UEBD, SSO…)
_RE_SIGLA         = re.compile(r'\b[A-Z]{2,5}\b')

def _es_dominio(palabra):
    """True si la palabra debe ignorarse (dominio técnico / sigla / código de fase)."""
    p = palabra.strip(".,;:()\"'").lower()
    if p in PALABRAS_DOMINIO:
        return True
    if _RE_CODIGO_FASE.fullmatch(palabra.strip(".,;:()")):
        return True
    if _RE_SIGLA.fullmatch(palabra.strip(".,;:()")):
        return True
    return False


def _parrafo_es_titulo(texto):
    """Heurística: párrafo muy corto o sin verbo conjugado → probablemente título."""
    palabras = texto.split()
    if len(palabras) <= 5:
        return True
    # Si no contiene ninguna vocal acentuada o terminación verbal típica, es título
    return False


# ════════════════════════════════════════════════════════════════════════════
# REGLAS PROPIAS — concordancia (regex conservador)
# ════════════════════════════════════════════════════════════════════════════

# ── Plural adjetivo + sustantivo singular ────────────────────────────────────
# Solo adjetivos comparativos/supletivos que NO son sustantivos por sí mismos
_ADJ_PLURAL = (
    # comparativos irregulares plurales
    "mejores", "peores", "mayores", "menores", "superiores", "inferiores",
    # adjetivos plurales frecuentes en informes
    "nuevos", "nuevas", "buenos", "buenas", "malos", "malas",
    "altos", "altas", "bajos", "bajas",
    "positivos", "positivas", "negativos", "negativas",
    "elevados", "elevadas", "reducidos", "reducidas",
    "significativos", "significativas",
    "principales", "anteriores", "posteriores", "actuales",
    "distintos", "distintas", "diversos", "diversas", "varios", "varias",
    "continuos", "continuas", "frecuentes",
    "totales", "parciales", "externos", "externas", "internos", "internas",
    "adicionales", "múltiples",
)

def _parece_singular(palabra):
    """True si la palabra parece un sustantivo en singular (heurística)."""
    p = palabra.lower().rstrip(".,;:")
    if _es_dominio(p):
        return False  # no juzgamos palabras de dominio
    # Invariables conocidos
    invariables = {"análisis", "crisis", "tesis", "síntesis", "paréntesis",
                   "virus", "atlas", "alias", "campus", "lunes", "martes",
                   "miércoles", "jueves", "viernes", "status"}
    if p in invariables:
        return False
    return not p.endswith("s")


def _detectar_concordancia_numero(texto):
    """Detecta adjetivo plural + sustantivo singular (ej: "mejores rendimiento")."""
    if _parrafo_es_titulo(texto):
        return []
    errores = []
    patron = re.compile(
        r'\b(' + '|'.join(re.escape(a) for a in _ADJ_PLURAL) + r')\s+([A-Za-záéíóúüñÁÉÍÓÚÜÑ]{4,})\b',
        re.IGNORECASE | re.UNICODE,
    )
    for m in patron.finditer(texto):
        adj, sust = m.group(1), m.group(2)
        if _es_dominio(sust):
            continue
        if _parece_singular(sust):
            errores.append((
                "CONCORDANCIA-NÚMERO",
                f'Adjetivo plural con sustantivo aparentemente singular: "{adj}" + "{sust}"',
                f'  Fragmento: "{m.group()}"  →  ¿"{adj[:-2] if adj.endswith("es") else adj[:-1]}" o "{sust}s"?',
            ))
    return errores


# ── Artículo con género incorrecto ───────────────────────────────────────────
# Terminaciones INEQUÍVOCAMENTE femeninas (no terminan en -a)
# OJO: "-ión" genérico NO va aquí porque es ambiguo:
#   camión, avión, gorrión → masculinos
#   extracción, producción → femeninos, pero ya cubiertos por "-ción"
_TERM_FEMENINAS = ("ción", "sión", "dad", "tad", "tud", "umbre")

# Terminaciones INEQUÍVOCAMENTE masculinas (no terminan en -o)
# OJO: NO incluimos "-ón" porque matchea -ción/-sión que son femeninos
_TERM_MASCULINAS = ("aje", "ismo", "miento", "mento")


def _detectar_articulo_genero(texto):
    """Detecta artículo con género incorrecto de forma conservadora."""
    if _parrafo_es_titulo(texto):
        return []
    errores = []

    # "el/un" + sustantivo con terminación inequívocamente FEMENINA
    for m in re.finditer(r'\b(el|un)\s+([A-Za-záéíóúüñÁÉÍÓÚÜÑ]{4,})\b',
                          texto, re.IGNORECASE | re.UNICODE):
        art, sust = m.group(1), m.group(2).lower()
        if _es_dominio(sust):
            continue
        # Excepción legítima: "el agua", "el área", "el alma" (femeninos con 'a' tónica)
        if re.match(r'^[aá]', sust):
            continue
        if any(sust.endswith(t) for t in _TERM_FEMENINAS):
            errores.append((
                "CONCORDANCIA-GÉNERO",
                f'Artículo "{art}" con sustantivo de terminación femenina "{sust}"',
                f'  Fragmento: "{m.group()}"  →  ¿"la {sust}"?',
            ))

    # "la/una" + sustantivo con terminación inequívocamente MASCULINA
    for m in re.finditer(r'\b(la|una)\s+([A-Za-záéíóúüñÁÉÍÓÚÜÑ]{4,})\b',
                          texto, re.IGNORECASE | re.UNICODE):
        art, sust = m.group(1), m.group(2).lower()
        if _es_dominio(sust):
            continue
        if any(sust.endswith(t) for t in _TERM_MASCULINAS):
            errores.append((
                "CONCORDANCIA-GÉNERO",
                f'Artículo "{art}" con sustantivo de terminación masculina "{sust}"',
                f'  Fragmento: "{m.group()}"  →  ¿"el {sust}"?',
            ))

    return errores


# ── Palabras o frases duplicadas ─────────────────────────────────────────────
def _detectar_duplicados(texto):
    errores = []
    re_pal = re.compile(r'\b(\w{3,})\s+\1\b', re.IGNORECASE | re.UNICODE)
    re_fra = re.compile(r'(.{12,60})[,.]?\s+\1',  re.IGNORECASE | re.UNICODE)
    for m in re_pal.finditer(texto):
        if not _es_dominio(m.group(1)):
            errores.append(("DUPLICADO", f'Palabra repetida: "{m.group()}"', ""))
    for m in re_fra.finditer(texto):
        errores.append(("DUPLICADO", f'Frase repetida: "{m.group()[:80]}…"', ""))
    return errores


# ════════════════════════════════════════════════════════════════════════════
# LanguageTool API — filtro mínimo + exclusión de dominio
# ════════════════════════════════════════════════════════════════════════════
LT_URL = "https://api.languagetool.org/v2/check"

REGLAS_LT_EXCLUIR = {
    "WHITESPACE_RULE", "SENTENCE_WHITESPACE", "PUNCTUATION_PARAGRAPH_END",
    "COMMA_PARENTHESIS_WHITESPACE", "EN_QUOTES", "UNPAIRED_BRACKETS",
    "DASH_RULE", "UPPERCASE_SENTENCE_START", "WORD_CONTAINS_UNDERSCORE",
    "CURRENCY",
    "MORFOLOGIK_RULE_ES",       # spellcheck puro → mucho ruido con términos técnicos
    "NUMERO_ORDINAL",           # formatos numéricos
    "ES_SPLIT_WORDS",           # separación de palabras compuestas
    "ENERO_01",                 # sugiere quitar ceros en fechas (06 → 6): no aplica
}
CATEGORIAS_LT_EXCLUIR = {"TYPOGRAPHY", "CASING", "STYLE", "TYPOS"}


def _es_relevante_lt(regla_id, categoria, fragmento):
    if regla_id in REGLAS_LT_EXCLUIR:
        return False
    if categoria.upper() in CATEGORIAS_LT_EXCLUIR:
        return False
    # Si CUALQUIER palabra del fragmento es de dominio → ignorar el match completo
    # (cubre casos como "Resultados batimetría" donde "batimetría" es dominio)
    palabras = re.findall(r'\b\w+\b', fragmento)
    if any(_es_dominio(p) for p in palabras):
        return False
    return True


def _consultar_lt(texto, reintentos=3):
    payload = {"text": texto, "language": "es", "enabledOnly": "false"}
    for intento in range(1, reintentos + 1):
        try:
            resp = requests.post(LT_URL, data=payload, timeout=30)
            if resp.status_code == 200:
                return resp.json().get("matches", [])
            if resp.status_code == 429:
                espera = 7 * intento
                print(f"    [rate-limit] esperando {espera}s…")
                time.sleep(espera)
            else:
                print(f"    [HTTP {resp.status_code}]")
                return []
        except requests.RequestException as e:
            print(f"    [red] {e}")
            time.sleep(5)
    return []


# ════════════════════════════════════════════════════════════════════════════
# Motor principal
# ════════════════════════════════════════════════════════════════════════════
PAUSA = 3.2   # segundos entre llamadas a la API (límite: 20 req/min)


def _flush():
    """Fuerza el flush de stdout para que el servidor web capture la línea de inmediato."""
    try:
        sys.stdout.flush()
    except Exception:
        pass


def revisar_gramatica(ruta_docx, faenas=None):
    """Revisa el informe Word.

    faenas : lista de claves a revisar (ej. ["MLP", "CMZ"]).
             None o lista vacía → revisa todas.
    """
    faenas_set = set(faenas) if faenas else None

    print(f"── Revisor gramatical: {Path(ruta_docx).name}")
    if faenas_set:
        print(f"  Secciones a revisar: {', '.join(sorted(faenas_set))}")
    _flush()

    parrafos  = extraer_parrafos(ruta_docx)
    a_revisar = [(i, t, e) for i, (t, e) in enumerate(parrafos, 1) if len(t) >= 25]
    total_p   = len(a_revisar)
    print(f"  Párrafos a revisar: {total_p}  (~{total_p * PAUSA / 60:.1f} min)")
    _flush()

    compania_actual      = "—"
    compania_clave_actual = None   # clave corta: "MLP", "CMZ", …
    total_errores        = 0
    primera_llamada      = True
    idx_revision         = 0   # contador de párrafos revisados (excluye encabezados de sección)

    for num, texto, estilo in a_revisar:

        nueva = detectar_compania(texto)
        if nueva:
            if nueva != compania_actual:
                compania_actual = nueva
                # Extraer clave del formato "MLP – Los Pelambres"
                compania_clave_actual = nueva.split(" – ")[0].strip()
                # Mostrar cabecera solo si vamos a revisar esta sección
                if faenas_set is None or compania_clave_actual in faenas_set:
                    print(f"── {compania_actual}")
                    _flush()
            continue

        # Saltar párrafos de secciones no seleccionadas
        if faenas_set is not None and compania_clave_actual not in faenas_set:
            continue

        idx_revision += 1
        # Línea de progreso visible antes de la llamada a la API
        preview_corto = texto[:60].replace('\n', ' ')
        print(f"  [{idx_revision}/{total_p}] {preview_corto}…")
        _flush()

        errores_parrafo = []

        # ── 1. Reglas propias (sin red, instantáneas) ──
        errores_parrafo += _detectar_duplicados(texto)
        errores_parrafo += _detectar_concordancia_numero(texto)
        errores_parrafo += _detectar_articulo_genero(texto)

        # ── 2. LanguageTool API ──
        if not primera_llamada:
            time.sleep(PAUSA)
        primera_llamada = False

        matches = _consultar_lt(texto)
        for m in matches:
            regla_id  = m.get("rule", {}).get("id", "")
            categoria = m.get("rule", {}).get("category", {}).get("id", "")
            cat_name  = m.get("rule", {}).get("category", {}).get("name", categoria)
            mensaje   = m.get("message", "")
            offset    = m.get("offset", 0)
            longitud  = m.get("length", 0)
            sugs      = [r["value"] for r in m.get("replacements", [])[:3]]
            fragmento = texto[offset: offset + longitud]

            if not _es_relevante_lt(regla_id, categoria, fragmento):
                continue

            if regla_id == "A_PARTICIPIO":
                resto = texto[offset + longitud:].lstrip()
                m_sig = re.match(r'\w+', resto)
                if m_sig:
                    sig = m_sig.group(0).lower()
                    _PARTICIPIOS_IRREG = {
                        "abierto", "vuelto", "puesto", "hecho", "dicho", "escrito",
                        "roto", "muerto", "visto", "resuelto", "disuelto",
                    }
                    es_participio = (
                        re.search(r'(ado|ada|idos|idas|ido|ida)s?$', sig)
                        or sig in _PARTICIPIOS_IRREG
                    )
                    if not es_participio:
                        continue

            sugerencias = ", ".join(sugs) if sugs else "—"
            ctx_ini = max(0, offset - 40)
            ctx_fin = min(len(texto), offset + longitud + 40)
            prefijo = ("…" if ctx_ini > 0 else "") + texto[ctx_ini:offset]
            sufijo  = texto[offset + longitud:ctx_fin] + ("…" if ctx_fin < len(texto) else "")
            contexto = f'{prefijo}[{fragmento}]{sufijo}'
            errores_parrafo.append((
                f"{regla_id} [{cat_name}]",
                mensaje,
                f'  "{contexto}"  →  {sugerencias}',
            ))

        if errores_parrafo:
            total_errores += len(errores_parrafo)
            for regla, msg, detalle in errores_parrafo:
                print(f"  ! [{regla}] {msg}")
                print(f"    {detalle}")
            _flush()
        else:
            # Sobreescribir la línea de progreso con ✓ para indicar que está limpio
            print(f"  ✓ Párrafo {idx_revision} sin observaciones")
            _flush()

    print(f"── Revisión completada: {total_errores} observación(es) en {idx_revision} párrafos")
    _flush()


# ── Entrada ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) > 1:
        ruta = sys.argv[1]
    else:
        print("Abriendo selector de archivo...")
        ruta = _seleccionar_archivo()

    if not ruta or not Path(ruta).is_file():
        print("[ERROR] No se seleccionó ningún archivo válido.")
        sys.exit(1)

    revisar_gramatica(ruta)
