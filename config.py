"""Configuración y constantes del generador de informe semanal."""

import os
from pathlib import Path

# Ignora warnings de formato condicional de openpyxl que no afectan el flujo.
import warnings
warnings.filterwarnings("ignore", message="Conditional Formatting extension is not supported")

# Activa o desactiva el modo debug (rutas auto-construidas con fallback a selector).
MODO_DEBUG = True

# Ruta base donde están las carpetas anuales del informe semanal.
RUTA_BASE_SEMANAS = r"N:\01 Reporting\09 Informe Semanal"

# Parte de la ruta sin la unidad de disco (para permitir reemplazar N: por otra unidad).
_RUTA_BASE_SIN_DISCO = r"\01 Reporting\09 Informe Semanal"

# Abreviaturas de meses para construir nombres de carpetas y archivos.
_MESES_ABR = {
    "01": "ene", "02": "feb", "03": "mar", "04": "abr",
    "05": "may", "06": "jun", "07": "jul", "08": "ago",
    "09": "sep", "10": "oct", "11": "nov", "12": "dic",
}
_MESES_NOMBRE = {
    "01": "Enero",    "02": "Febrero",   "03": "Marzo",     "04": "Abril",
    "05": "Mayo",     "06": "Junio",     "07": "Julio",     "08": "Agosto",
    "09": "Septiembre", "10": "Octubre", "11": "Noviembre", "12": "Diciembre",
}

def _build_raiz(base, year, mes, num_semana, di, abr_ini, df, abr_fin):
    """Construye la ruta raíz para un mes dado."""
    carpeta_mes    = f"{int(mes)} - {_MESES_NOMBRE[mes]}"
    carpeta_semana = f"{num_semana}_Semana- {di} {abr_ini} al {df} {abr_fin}"
    return base / str(year) / carpeta_mes / carpeta_semana


def construir_rutas_semana(num_semana, dia_inicio, mes_inicio, dia_fin, mes_fin, year, disco=None):
    """Devuelve las rutas esperadas para la semana dada según la estructura de carpetas estándar.

    Si se indica `disco` (ej. "N:" o "Z:"), se usa esa unidad en lugar de la definida en
    RUTA_BASE_SEMANAS. Si no se indica, se usa RUTA_BASE_SEMANAS completa.

    Cuando la semana cruza dos meses (mes_inicio != mes_fin) se generan candidatos para ambas
    carpetas de mes; la primera que exista en disco es la que se usa.
    """
    mi = str(mes_inicio).zfill(2)
    mf = str(mes_fin).zfill(2)
    di = str(dia_inicio).zfill(2)
    df = str(dia_fin).zfill(2)

    abr_ini = _MESES_ABR[mi]
    abr_fin = _MESES_ABR[mf]

    if disco:
        disco_norm = disco.strip().rstrip("/\\")
        if not disco_norm.endswith(":"):
            disco_norm += ":"
        base = Path(disco_norm + _RUTA_BASE_SIN_DISCO)
    else:
        base = Path(RUTA_BASE_SEMANAS)

    # Candidatos: primero mes_fin (convención habitual), luego mes_inicio si son distintos
    candidatos = [_build_raiz(base, year, mf, num_semana, di, abr_ini, df, abr_fin)]
    if mi != mf:
        candidatos.append(_build_raiz(base, year, mi, num_semana, di, abr_ini, df, abr_fin))

    # Usar la primera carpeta que exista; si ninguna existe, usar la primera candidata
    raiz = next((c for c in candidatos if c.is_dir()), candidatos[0])

    sso_dir = raiz / "06 -SSO"

    # Carpeta Gestión Hídrica: la primera subcarpeta que empiece con "07"
    gh_dir = None
    if raiz.is_dir():
        candidates = [f for f in raiz.iterdir() if f.is_dir() and f.name.startswith("07")]
        if candidates:
            gh_dir = candidates[0]
    if gh_dir is None:
        gh_dir = raiz / "07 -Gestión Hídrica"  # fallback para construcción de ruta

    return {
        "raiz": raiz,
        "excel_madre":           raiz / f"Semana {num_semana} -  {di} {abr_ini} al {df} {abr_fin}.xlsx",
        "excel_indicadores_dir": sso_dir,
        "carpeta_destino":       str(raiz),
        "nombre_archivo":        "Informe_Automatizado",
        "informes_dirs": {
            "MLP":  raiz / "01 -MLP",
            "CEN":  raiz / "02 -CEN",
            "ANT":  raiz / "03 -ANT",
            "CMZ":  raiz / "04 -CMZ",
            "FCAB": raiz / "05 -FCAB",
        },
        "excels_adicionales_dirs": {
            "SSO":             sso_dir,
            "Gestión Hídrica": gh_dir,
        },
    }

# Controla si se incluye la página de estado de fases de desarrollo.
INCLUIR_ESTADO_FASES_DESARROLLO = False

# Define el orden oficial de las faenas dentro del informe.
ORDEN_OFICIAL = ["MLP", "CEN", "ANT", "CMZ", "FCAB"]

# Define la configuración base por compañía para exportar sus tablas.
CONFIG_COMPANIAS = {
    "MLP": {"nombre": "Los Pelambres", "rango": "B3:AD33", "alto": 7.69,
            "rango_desviaciones": "C37:J58"},
    "ANT": {"nombre": "Antucoya", "rango": "A3:AC45", "alto": 10.13},
    "CEN": {"nombre": "Centinela", "rango": "A3:AC85", "alto": 21.41},
    "CMZ": {"nombre": "Zaldívar", "rango": "A3:AC35", "alto": 6.85},
    "FCAB": {"nombre": "FCAB", "rango": "A3:V19", "alto": 3.21},
}

# Celdas exactas en el Excel madre donde leer la desviación (dif unidades, dif %) por KPI.
# Columna E = desviación en unidades, Columna G = desviación porcentual.
# Agregar entradas para ANT, CEN, CMZ, FCAB cuando se confirmen sus filas.
CONFIG_CELDAS_DESVIACIONES = {
    "MLP": {
        # ── Mina ──────────────────────────────────────────────────────────────
        "Movimiento Mina":          ("E39", "G39"),
        "Extracción":               ("E40", "G40"),
        "Extracción Lastre":        ("E41", "G41"),
        "Extracción Mineral":       ("E42", "G42"),
        "Remanejo":                 ("E43", "G43"),
        # ── Concentradora (fila 44 = cabecera de sección) ─────────────────────
        "Mineral Procesado":        ("E45", "G45"),
        "Ley Cu":                   ("E46", "G46"),
        "Recuperación Cu":          ("E47", "G47"),
        "Cu Fino Producido":        ("E48", "G48"),
        "Concentrado Producido":    ("E49", "G49"),
        "Concentrado Filtrado":     ("E50", "G50"),
        "Cu Fino Pagable Filtrado": ("E51", "G51"),
        "Molibdeno fino pagable":   ("E52", "G52"),
        "Arenas Depositadas":       ("E53", "G53"),
        "Arenas Compactadas":       ("E54", "G54"),
    },
    "ANT": {
        # ── Mina ──────────────────────────────────────────────────────────────
        "Movimiento Mina":              ("D50", "F50"),
        "Extracción Mina":              ("D51", "F51"),
        "Extracción Mineral":           ("D52", "F52"),
        "Extracción Lastre":            ("D53", "F53"),
        "Remanejo":                     ("D54", "F54"),
        # ── Mina (Fases) — encabezados de sección ─────────────────────────────
        "Extracción de Mineral":        ("D56", "F56"),
        "Extracción de Lastre":         ("D62", "F62"),
        # ── Planta ────────────────────────────────────────────────────────────
        "Mineral Apilado":              ("D68", "F68"),
        "Mineral Beneficiado":          ("D69", "F69"),
        "Ley Cu":                       ("D70", "F70"),
        "Recuperación Cu Beneficiado":  ("D71", "F71"),
        "Descarga de Ripios":           ("D72", "F72"),
        "Cobre Fino Producido":         ("D73", "F73"),
    },
    "CEN": {
        # ── Mina ──────────────────────────────────────────────────────────────
        "Movimiento Mina":                          ("D91",  "F91"),
        "Movimiento en Rajo Tesoro":                ("D92",  "F92"),
        "Movimiento en Rajo Esperanza":             ("D96",  "F96"),
        "Movimiento en Rajo Óxido Encuentro":       ("D100", "F100"),
        "Movimiento en Rajo Esperanza Sur:":        ("D102", "F102"),
        "Movimiento en Rajo Encuentro Sulfuros":    ("D107", "F107"),
        # ── Sulfuros ──────────────────────────────────────────────────────────
        "Mineral Procesado":                        ("D112", "F112"),
        "Ley Cu":                                   ("D113", "F113"),
        "Recuperación Cu":                          ("D114", "F114"),
        "Cu Fino Producido":                        ("D115", "F115"),
        "Concentrado Filtrado":                     ("D116", "F116"),
        "Cu Fino Pagable Filtrada":                 ("D117", "F117"),
        "Recuperación Au":                          ("D119", "F119"),
        "Au Fino Pagable Filtrada":                 ("D120", "F120"),
        "Molibdeno Fino Pagable":                   ("D121", "F121"),
        # ── Cátodos ───────────────────────────────────────────────────────────
        "Producción Total de Cátodos de Cu":        ("D123", "F123"),
    },
    "CMZ": {
        # ── Mina ──────────────────────────────────────────────────────────────
        "Movimiento Mina":          ("D41", "F41"),
        "Extracción":               ("D42", "F42"),
        "Extracción Mineral":       ("D48", "F48"),
        "Extracción Lastre":        ("D49", "F49"),
        "Remanejo":                 ("D50", "F50"),
        # ── Planta ────────────────────────────────────────────────────────────
        "Mineral Apilado HL":       ("D52", "F52"),
        "Mineral Beneficiado HL":   ("D53", "F53"),
        "Ley Apilado HL TCu":       ("D54", "F54"),
        "Mineral Apilado DL":       ("D55", "F55"),
        "Mineral Beneficiado DL":   ("D56", "F56"),
        "Ley Apilado DL TCu":       ("D57", "F57"),
        "Remanejo Ripios":          ("D58", "F58"),
        "PLS":                      ("D59", "F59"),
        "Cobre Fino Producido":     ("D60", "F60"),
    },
}

# Define el orden esperado de subtítulos para las principales desviaciones por compañía.
ORDEN_PRINCIPALES_DESVIACIONES = {
    "MLP": {
        "Principales Desviaciones": ["?"],
        "Mina": ["Movimiento Mina", "Total Extracción", "Extracción", "Remanejo"],
        "Concentradora": [""],
        "Planta Desaladora": ["?"],
        "Gestión Hídrica": [""],
    },
    "CEN": {
        "Principales Desviaciones": ["?"],
        "Mina": [
            "Movimiento Mina",
            "Movimiento en Rajo Tesoro",
            "Movimiento en Rajo Esperanza",
            "Movimiento en Rajo Óxido Encuentro",
            "Movimiento en Rajo Esperanza Sur:",
            "Movimiento en Rajo Encuentro Sulfuros",
        ],
        "Sulfuros": [""],
        "Cátodos": ["Planta Hidro MET", "Planta Hidro OXE"],
    },
    "ANT": {
    "Principales Desviaciones": ["?"],
    "Mina": [
      "Movimiento Mina",
      "Extracción Mina",
      "Extracción de Mineral",
      "Extracción de lastre",
      "Remanejo",
      "Extracción a desarrollo",
    ],
    "Planta": [""],
    },
    "CMZ": {
        "Principales Desviaciones": ["?"],
        "Mina": ["Movimiento Mina", "Extracción", "Extracción Mineral", "Extracción Lastre", "Remanejo"],
        "Planta": [""],
    },
    "FCAB": {
        "Principales Desviaciones": ["?"],
        "Tren": ["#Transporte Total de Tren", "Transporte de ácido", "Transporte de Cobre", "Transporte de Concentrados"],
        "Camión": ["Transporte Total de Camión"],
    },
}

NIVEL_BASE_POR_SECCION = {
    "Principales Desviaciones": 2,
    "Mina": 2,
    "Detalle por fases": 2,
    "Planta": 1,
    "Sulfuros": 1,
    "Cátodos": 1,
    "Concentradora": 1,
    "Planta Desaladora": 2,
    "Gestión Hídrica": 1,
    "Tren": 2,
    "Camión": 2,
}

NIVEL_POR_COMPANIA_SECCION_SUBTITULO = {
    "MLP": {"Mina": {"Movimiento Mina": 1}},
    "CEN": {"Mina": {"Movimiento Mina": 1}},
    "ANT": {"Mina": {"Movimiento Mina": 1}},
    "CMZ": {"Mina": {"Movimiento Mina": 1}},
}


# Hojas adicionales del Excel madre que se validan por separado (no son compañías).
# compania_fuente: clave cuyo Word contiene el texto de esa sección.
CONFIG_HOJAS_ADICIONALES = {
    "Gestión Hídrica": {
        "hoja": "Gestión Hídrica",
        "rango": "A3:W20",
        "compania_fuente": "MLP",
    }
}

# Guarda la ruta de la plantilla Word usada para construir el informe final.
BASE_DIR = Path(__file__).resolve().parent
RUTA_PLANTILLA = BASE_DIR / "Template Viñetas Python.docx"

# Guarda el marcador del encabezado de las tablas SSO de respaldo.
SSO_MARCADOR_TABLA = "id del incidente"
