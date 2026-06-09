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


def construir_rutas_semana(num_semana, dia_inicio, mes_inicio, dia_fin, mes_fin, year, disco=None, carpeta_personalizada=None):
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

    if carpeta_personalizada:
        # El usuario eligió "No" al disco compartido: su carpeta se usa tal cual como raíz.
        raiz = Path(carpeta_personalizada)
    else:
        if disco:
            disco_norm = disco.strip().rstrip("/\\")
            if not disco_norm.endswith(":"):
                disco_norm += ":"
            base = Path(disco_norm + _RUTA_BASE_SIN_DISCO)
        else:
            base = Path(RUTA_BASE_SEMANAS)

        # Candidatos: primero mes_fin, luego mes_inicio si son distintos
        candidatos = [_build_raiz(base, year, mf, num_semana, di, abr_ini, df, abr_fin)]
        if mi != mf:
            candidatos.append(_build_raiz(base, year, mi, num_semana, di, abr_ini, df, abr_fin))
        raiz = next((c for c in candidatos if c.is_dir()), None)
        # Si el nombre exacto no existe, la carpeta puede estar nombrada de forma
        # ligeramente distinta (ej. "...04 junio" en vez de "...04 jun", o espacios
        # diferentes). Se localiza por número de semana dentro de las carpetas de
        # mes candidatas. Esto NO cambia el fallback de archivos: la selección de
        # Word/Excel por faena sigue cayendo a la semana anterior cuando el archivo
        # actual no existe; solo permite reconocer la carpeta y ofrecer renombrar.
        if raiz is None:
            for c in candidatos:
                mes_dir = c.parent
                if mes_dir.is_dir():
                    match = next((f for f in sorted(mes_dir.iterdir())
                                  if f.is_dir() and f.name.startswith(f"{num_semana}_Semana")), None)
                    if match:
                        raiz = match
                        break
        if raiz is None:
            raiz = candidatos[0]

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
# Columna E = desviación en unidades, F/G = desviación porcentual (varía por compañía).
# Pendiente confirmar celdas para ANT, CEN, CMZ.
CONFIG_CELDAS_DESVIACIONES = {
    "MLP": {
        # ── Mina ──────────────────────────────────────────────────────────────
        "Movimiento Mina":          ("E40", "G40", "H40"),
        "Extracción":               ("E41", "G41", "H41"),
        "Extracción Lastre":        ("E42", "G42", "H42"),
        "Extracción Estéril":       ("E42", "G42", "H42"),
        "Extracción Mineral":       ("E43", "G43", "H43"),
        "Remanejo":                 ("E44", "G44", "H44"),
        # ── Concentradora (fila 45 = cabecera de sección) ─────────────────────
        "Mineral Procesado":        ("E46", "G46", "H46"),
        "Ley Cu":                   ("E47", "G47", "H47"),
        "Recuperación Cu":          ("E48", "G48", "H48"),
        "Cu Fino Producido":        ("E49", "G49", "H49"),
        "Concentrado Producido":    ("E50", "G50", "H50"),
        "Concentrado Filtrado":     ("E51", "G51", "H51"),
        "Cu Fino Filtrado Pagable": ("E52", "G52", "H52"),
        "Molibdeno":               ("E53", "G53", "H53"),
        "Arenas Depositadas":       ("E54", "G54", "H54"),
        "Arenas Compactadas":       ("E55", "G55", "H55"),
    },
    "ANT": {
        # ── Mina ──────────────────────────────────────────────────────────────
        "Movimiento Mina":              ("D50", "F50", "G50"),
        "Extracción Mina":              ("D51", "F51", "G51"),
        # Nota: "Extracción Mineral" (D52) y "Extracción Lastre" (D53) son la MISMA
        # cifra que "Extracción de Mineral" (D56) y "Extracción de Lastre" (D62) —
        # duplicados en el Excel. El Word siempre usa la forma "de Mineral/Lastre",
        # así que se omiten aquí para que la verificación inversa no las marque como
        # no revisadas (el matcher por solapamiento cubre igual la forma sin "de").
        "Remanejo":                     ("D54", "F54", "G54"),
        # ── Mina (Fases) — encabezados de sección ─────────────────────────────
        "Extracción de Mineral":        ("D56", "F56", "G56"),
        "Extracción de Lastre":         ("D62", "F62", "G62"),
        # ── Detalle por fases — Extracción de Mineral (Word: F05..F08) ─────────
        "F05 mineral":                  ("D57", "F57", "G57"),
        "F06 mineral":                  ("D58", "F58", "G58"),
        "F07 mineral":                  ("D59", "F59", "G59"),
        "F08 mineral":                  ("D60", "F60", "G60"),
        # ── Detalle por fases — Extracción de Lastre (Word: F05..F08) ──────────
        "F05 lastre":                   ("D63", "F63", "G63"),
        "F06 lastre":                   ("D64", "F64", "G64"),
        "F07 lastre":                   ("D65", "F65", "G65"),
        "F08 lastre":                   ("D66", "F66", "G66"),
        # ── Planta ────────────────────────────────────────────────────────────
        "Mineral Apilado":              ("D68", "F68", "G68"),
        "Mineral Beneficiado":          ("D69", "F69", "G69"),
        "Ley Cu":                       ("F70", "F70", "G70"),
        "Recuperación Cu":              ("F71", "F71", "G71"),
        "Descarga de Ripios":           ("D72", "F72", "G72"),
        "Cobre Fino Producido":         ("D73", "F73", "G73"),
    },
    "CEN": {
        # ── Mina ──────────────────────────────────────────────────────────────
        "Movimiento Mina":                          ("D91",  "F91",  "G91"),
        "Movimiento en Rajo Tesoro":                ("D92",  "F92",  "G92"),
        "Movimiento en Rajo Esperanza":             ("D96",  "F96",  "G96"),
        "Movimiento en Rajo Óxido Encuentro":       ("D100", "F100", "G100"),
        "Movimiento en Rajo Esperanza Sur:":        ("D102", "F102", "G102"),
        "Movimiento en Rajo Encuentro Sulfuros":    ("D107", "F107", "G107"),
        # ── Sulfuros ──────────────────────────────────────────────────────────
        "Mineral Procesado":                        ("D112", "F112", "G112"),
        "Ley Cu":                                   ("D113", "F113", "G113"),
        "Recuperación Cu":                          ("D114", "F114", "G114"),
        "Cu Fino Producido":                        ("D115", "F115", "G115"),
        "Concentrado Filtrado":                     ("D116", "F116", "G116"),
        "Cu Fino Pagable Filtrado":                 ("D117", "F117", "G117"),
        "Ley Au":                                   ("F118", "F118", "G118"),
        "Recuperación Au":                          ("D119", "F119", "G119"),
        "Au Fino Pagable Filtrado":                 ("D120", "F120", "G120"),
        "Mo Fino Pagable Filtrado":                 ("D121", "F121", "G121"),
        # ── Cátodos ───────────────────────────────────────────────────────────
        "Producción Total de Cátodos de Cu":        ("D123", "F123", "G123"),
        # ── Cátodos — Planta Hidro MET ────────────────────────────────────────
        "Producción de Cátodos de Cu MET":          ("D125", "F125", "G125"),
        "Mineral Apilado MET":                      ("D126", "F126", "G126"),
        "Mineral Apilado":                          ("D126", "F126"),
        "Mineral Beneficiado MET":                  ("D127", "F127", "G127"),
        "Ley de Cu MET":                            ("F128", "F128", "G128"),
        "Producción de ROM":                        ("D130", "F130", "G130"),
        # ── Cátodos — Planta Hidro OXE ────────────────────────────────────────
        "Producción de Cátodos de Cu OXE":          ("D132", "F132", "G132"),
        "Mineral Apilado OXE":                      ("D133", "F133", "G133"),
        "Mineral Beneficiado OXE":                  ("D134", "F134", "G134"),
        "Ley de Cu OXE":                            ("F135", "F135", "G135"),
    },
    "CMZ": {
        # ── Mina ──────────────────────────────────────────────────────────────
        "Movimiento Mina":          ("D41", "F41", "G41"),
        "Extracción":               ("D42", "F42", "G42"),
        "Extracción Mineral":       ("D48", "F48", "G48"),
        "Extracción Lastre":        ("D49", "F49", "G49"),
        "Remanejo":                 ("D50", "F50", "G50"),
        # ── Planta ────────────────────────────────────────────────────────────
        "Mineral Apilado HL":       ("D52", "F52", "G52"),
        "Mineral Beneficiado HL":   ("D53", "F53", "G53"),
        "Ley Apilado HL TCu":       ("D54", "F54", "G54"),
        "Mineral Apilado DL":       ("D55", "F55", "G55"),
        "Mineral Beneficiado DL":   ("D56", "F56", "G56"),
        "Ley Apilado DL TCu":       ("D57", "F57", "G57"),
        "Remanejo Ripios":          ("D58", "F58", "G58"),
        "PLS":                      ("D59", "F59", "G59"),
        "Cobre Fino Producido":     ("D60", "F60", "G60"),
    },
    "FCAB": {
        # ── Tren ──────────────────────────────────────────────────────────────
        "Transporte de ácido":          ("E25", "F25", "G25"),
        "Transporte de Cobre":          ("E26", "F26", "G26"),
        "Transporte de Concentrados":   ("E27", "F27", "G27"),
        "Transporte Total de Tren":     ("E29", "F29", "G29"),
        # ── Camión ────────────────────────────────────────────────────────────
        "Transporte Total de Camión":   ("E34", "F34", "G34"),
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

# KPIs cuyo label empieza con alguno de estos prefijos deben ignorarse en la validación.
# Comparación normalizada (sin tildes, minúsculas).
CONFIG_KPI_PREFIJOS_EXCLUIDOS = {
    "FCAB": {"minera"},
    "ANT":  {"se proyecta", "nota"},
}

# KPIs requeridos por compañía: si no aparecen en la validación, se registra error.
CONFIG_KPI_REQUERIDOS = {
    "ANT": ["Movimiento Mina"],
}

# KPIs para los que solo se valida la desviación % (primer valor numérico del Word).
# Aplica a cualquier compañía. Normalizado sin tildes/minúsculas al usar.
CONFIG_KPI_SOLO_DESVIACION = {"Recuperación Cu", "Recuperacion Cu"}

# La desviación en UNIDADES se compara redondeando ambos lados a ENTERO por
# defecto. Los KPIs listados aquí son la excepción: se comparan con 1 decimal.
# Comparación por igualdad exacta del valor normalizado (sin tildes, minúsculas).
CONFIG_KPI_UNID_1_DECIMAL = {"au fino pagable filtrado", "pls"}

# Después de validar este KPI (normalizado), detener la validación para esa compañía.
# Usado para compañías que tienen texto libre después de los KPIs formales (ej. MLP Planta Desaladora).
CONFIG_KPI_FIN_VALIDACION = {
    "MLP": "acumulado al ano",
}

# KPIs que deben ignorarse en la validación por compañía (el label es exactamente
# como aparece en el Word, comparación normalizada — sin tildes, minúsculas).
CONFIG_KPI_EXCLUIDOS = {
    "ANT": {"Pala", "Cargador", "Extracción a desarrollo"},
    "CEN": {"Fase", "Remanejo"},
    "CMZ": {"Fase"},
    "MLP": {
        "En Planta Desaladora",
        "En términos de capacidad de impulsión",
        "Agua Continental Consumida",
        "Recirculación Mauro a Planta",
        "Intensidad de Uso de Agua en Planta",
    },
}

# Subtítulos que marcan un cambio de contexto dentro de la sección de una compañía.
# Cuando el validador encuentra una de estas líneas actualiza el sufijo de contexto,
# que se añade al label al buscar en CONFIG_CELDAS_DESVIACIONES.
# Formato: { "CLAVE": { "Texto del subtítulo": "SUFIJO" } }
CONFIG_SUBSECCIONES_CONTEXTO = {
    "CEN": {
        "Planta Hidro MET": "MET",
        "Planta Hidro OXE": "OXE",
    },
}

# Etiquetas de KPI que, ADEMÁS de validarse normalmente, fijan un contexto para
# las líneas siguientes. A diferencia de CONFIG_SUBSECCIONES_CONTEXTO (subtítulos
# puros que se omiten), estas líneas sí son KPIs que se comparan. Se usa para las
# fases de ANT: tras "Extracción de Mineral" vienen las fases F05..F08 de mineral
# (D57..D60) y tras "Extracción de Lastre" las de lastre (D63..D66). El sufijo se
# añade al label de la fase al buscar su celda (ej. "F06" + "lastre" → "F06 lastre").
# Formato: { "CLAVE": { "prefijo_label_normalizado": "SUFIJO" } }
CONFIG_CONTEXTO_POR_LABEL = {
    "ANT": {
        "extraccion de mineral": "mineral",
        "extraccion de lastre":  "lastre",
    },
}

# Guarda la ruta de la plantilla Word usada para construir el informe final.
BASE_DIR = Path(__file__).resolve().parent
RUTA_PLANTILLA = BASE_DIR / "Template Viñetas Python.docx"

# Guarda el marcador del encabezado de las tablas SSO de respaldo.
SSO_MARCADOR_TABLA = "id del incidente"
