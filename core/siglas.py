"""Detección de siglas usadas sin contexto en los informes por compañía.

Una sigla se considera CON contexto si:
  1. Está definida en el propio documento, con su expansión entre paréntesis
     (ej. "...Accidente Sin Tiempo Perdido (ASTP)..."), en cualquier parte del
     texto — basta una vez para darla por conocida en todo el informe; o
  2. Está en el glosario amplio de siglas conocidas (SIGLAS_CONOCIDAS).

Todo lo demás (sigla en mayúsculas de 2-5 letras que no se define ni es
conocida) se reporta como "sigla sin contexto", con la línea donde aparece.

La detección de definiciones inline usa una variante simplificada del algoritmo
de Schwartz-Hearst: las iniciales de las palabras previas al paréntesis deben
deletrear la sigla (de derecha a izquierda, tolerando pocos saltos).
"""

import re

# ── Glosario amplio de siglas conocidas (no se reportan) ─────────────────────
# Se listan en MAYÚSCULA. El valor es la expansión (solo documentación; no se
# muestra). Donde la expansión exacta no es segura, se deja "" — igual cuenta
# como conocida para no generar ruido.
SIGLAS_CONOCIDAS = {
    # ── Compañías / faenas ────────────────────────────────────────────────
    "AMSA": "Antofagasta Minerals S.A.",
    "AMSM": "",
    "MLP": "Minera Los Pelambres",
    "CEN": "Centinela",
    "ANT": "Antucoya",
    "CMZ": "Zaldívar",
    "FCAB": "Ferrocarril de Antofagasta a Bolivia",
    # ── Planificación / control de gestión ────────────────────────────────
    "PM": "Plan Mensual",
    "PPTO": "Presupuesto",
    "PA": "Pala",
    "UEB": "",
    "UEBD": "",
    "EVU": "",
    "SIAD": "",
    "STR": "",
    "STC": "",
    "OCA": "",
    "OCAS": "",
    # ── Seguridad / accidentabilidad ──────────────────────────────────────
    "AAP": "Accidente de Alto Potencial",
    "ACTP": "Accidente Con Tiempo Perdido",
    "ASTP": "Accidente Sin Tiempo Perdido",
    "SSO": "Seguridad y Salud Ocupacional",
    # ── Mina / procesos / metalurgia ──────────────────────────────────────
    "OXE": "Óxido Encuentro",
    "ESS": "Esperanza Sur",
    "MET": "",
    "ROM": "Run of Mine",
    "HL": "Heap Leach",
    "DL": "Dump Leach",
    "PLS": "",
    "DF": "Disponibilidad Física",
    "CAEX": "Camión de Extracción",
    # ── Unidades / símbolos químicos (el patrón los captura como siglas) ──
    "KT": "kilotoneladas",
    "KTM": "",
    "KTPD": "kilotoneladas por día",
    "MT": "millones de toneladas",
    "MTPD": "",
    "TMF": "toneladas métricas finas",
    "TCU": "toneladas de cobre",
    "TMO": "",
    "TPD": "toneladas por día",
    "LPS": "litros por segundo",
    "CU": "cobre",
    "MO": "molibdeno",
    "AU": "oro",
    "AG": "plata",
    "FE": "hierro",
}

# Palabras de 2-5 letras que pueden aparecer en mayúscula sin ser siglas
# (conectores, direcciones). Se excluyen para no generar falsos positivos.
_PALABRAS_MAYUS_COMUNES = {
    "DE", "DEL", "LA", "EL", "EN", "POR", "CON", "SIN", "LOS", "LAS",
    "UN", "UNA", "SE", "NO", "SUR", "NTE", "ESTE", "OESTE", "ESTA",
    "AL", "SUS", "SON", "MAS", "YA", "FE",
}

# ── Normalización de tildes para comparar iniciales ──────────────────────────
_TILDES = str.maketrans(
    "áéíóúàèìòùäëïöüñÁÉÍÓÚÀÈÌÒÙÄËÏÖÜÑ",
    "aeiouaeiouaeiounAEIOUAEIOUAEIOUN",
)

def _sin_tildes(s):
    return s.translate(_TILDES)


# ── Patrones ─────────────────────────────────────────────────────────────────
# Candidata a sigla: 2-5 mayúsculas, con 's' minúscula opcional de plural
# (ej. "PAs"). Códigos de equipo/fase (PA10, F07, C29) no entran porque llevan
# dígitos y rompen el \b. group(1) = base de la sigla.
_RE_SIGLA_CAND = re.compile(r"\b([A-Z]{2,5})(s)?\b")

# Definición "expansión (SIGLA)": hasta 7 palabras previas al paréntesis.
_RE_DEF_PAREN = re.compile(
    r"((?:[^\W\d_][\w]*)(?:\s+[\w]+){0,7})\s*\(([A-Z]{2,5})s?\)",
    re.UNICODE,
)
# Definición inversa "SIGLA (expansión)".
_RE_DEF_PAREN_INV = re.compile(r"\b([A-Z]{2,5})s?\s*\(([^)]{3,80})\)")

# Tag de equipo/código: la sigla va seguida de "-<dígito>" (ej. HV-44141,
# VP-403-136). No es una sigla conceptual, sino un identificador.
_RE_TAG_NUM = re.compile(r"[-–]\d")


def _definida_inline(frase, sigla):
    """True si las iniciales de las palabras de `frase` deletrean `sigla`,
    alineando de derecha a izquierda y tolerando como mucho len(sigla) saltos
    (palabras intermedias que no aportan inicial, ej. "de", "y")."""
    sl = _sin_tildes(sigla).lower()
    palabras = re.findall(r"[^\W\d_]+", _sin_tildes(frase), re.UNICODE)
    iniciales = [p[0].lower() for p in palabras if p]
    if not iniciales:
        return False
    i = len(iniciales) - 1
    j = len(sl) - 1
    saltos = 0
    while i >= 0 and j >= 0:
        if iniciales[i] == sl[j]:
            i -= 1
            j -= 1
        else:
            i -= 1
            saltos += 1
            if saltos > len(sl):
                break
    return j < 0


def _siglas_definidas(texto):
    """Conjunto de siglas (MAYÚSCULA) definidas inline en cualquier parte del texto."""
    definidas = set()
    for m in _RE_DEF_PAREN.finditer(texto):
        frase, sigla = m.group(1), m.group(2)
        if _definida_inline(frase, sigla):
            definidas.add(sigla.upper())
    for m in _RE_DEF_PAREN_INV.finditer(texto):
        sigla, exp = m.group(1), m.group(2)
        if _definida_inline(exp, sigla):
            definidas.add(sigla.upper())
    return definidas


def detectar_siglas_sin_contexto(informes):
    """Recibe {clave: texto_word} y devuelve una lista de secciones (mismo formato
    que _resultados del validador) con las siglas sin contexto por compañía.

    Cada sección lleva tipo='siglas' para que el panel la renderice aparte.
    n_warn se deja en 0 para no contaminar el contador de diferencias de KPIs.
    """
    secciones = []
    for clave, texto in informes.items():
        if not texto:
            continue
        definidas = _siglas_definidas(texto)
        vistas = {}  # sigla -> primera línea de contexto
        for linea in texto.split("\n"):
            for m in _RE_SIGLA_CAND.finditer(linea):
                sig = m.group(1).upper()
                if (sig in SIGLAS_CONOCIDAS or sig in definidas
                        or sig in _PALABRAS_MAYUS_COMUNES):
                    continue
                # Hashtag de campaña de seguridad (#CAP, #HAP, #YDN…)
                if m.start() > 0 and linea[m.start() - 1] == "#":
                    continue
                # Tag de equipo/código (HV-44141, VP-403-136…)
                if _RE_TAG_NUM.match(linea, m.end()):
                    continue
                if sig not in vistas:
                    vistas[sig] = linea.strip()

        if not vistas:
            continue

        kpis = []
        for sig, linea in sorted(vistas.items()):
            detalle = linea if len(linea) <= 160 else linea[:157] + "…"
            kpis.append({
                "linea": linea,
                "label": sig,
                "excel_label": None,
                "estado": "revisar",
                "valores": [],
                "detalle": detalle,
            })

        secciones.append({
            "clave": clave,
            "nombre": "Siglas sin contexto",
            "tipo": "siglas",
            "kpis": kpis,
            "n_ok": 0,
            "n_warn": 0,
            "n_sin_celda": 0,
            "n_siglas": len(kpis),
            "estado": "warn",
            "error": None,
        })

    return secciones
