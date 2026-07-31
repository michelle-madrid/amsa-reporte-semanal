"""Palabras clave con las que se identifican los archivos Excel y Word en disco.

Los valores por defecto viven en PATRONES_DEFAULT; el usuario puede editarlos desde
el panel (tarjeta "Palabras clave de archivos"), que los guarda en `patrones.json`
en la raíz del proyecto. Ese archivo pisa a los defaults clave por clave, así que
basta con borrarlo para volver al comportamiento original.

Reglas de coincidencia (las mismas para Excel y Word):
  - Se compara contra el NOMBRE del archivo, sin distinguir mayúsculas ni tildes.
  - Basta con que el nombre CONTENGA la palabra clave (no tiene que empezar con ella).
  - Se pueden dar varias alternativas separadas por coma: cualquiera que calce sirve.
  - Palabra clave vacía = sin filtro: se acepta cualquier archivo de esa extensión
    (y como siempre, solo se auto-detecta cuando queda exactamente un candidato).
"""

import json
import threading
import unicodedata
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
RUTA_PATRONES = BASE_DIR / "patrones.json"

# Valores de fábrica. Cada bloque corresponde a una columna/sección del panel.
PATRONES_DEFAULT = {
    # Excel vinculado de cada faena + los dos Excel adicionales.
    "excel": {
        "MLP":             "mlp semana",
        "CEN":             "informe semanal",
        "ANT":             "informe semana",
        "CMZ":             "proyectado",
        "FCAB":            "amsa",
        "SSO":             "eventos seguridad",
        "Gestión Hídrica": "seguimiento",
    },
    # Word fuente de cada faena. Vacío = cualquier .docx de la carpeta.
    "word": {
        "MLP":  "",
        "CEN":  "",
        "ANT":  "",
        "CMZ":  "",
        "FCAB": "",
    },
    # Archivos que no cuelgan de una faena.
    "otros": {
        "indicadores_sso": "bdatos",   # base de datos SSO en 06 -SSO
        "excel_madre":     "semana",   # búsqueda alternativa del Excel Base
    },
    # Fragmentos de nombre cuyos vínculos externos NO deben actualizarse.
    "ignorar_vinculos": "cd mina",
}

_lock = threading.Lock()
_cache = None
_cache_mtime = None


def _norm(texto):
    """Minúsculas, sin tildes y con espacios colapsados."""
    t = unicodedata.normalize("NFD", str(texto))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return " ".join(t.lower().split())


def alternativas(patron):
    """Divide una palabra clave en sus alternativas normalizadas (separadas por coma)."""
    if not patron:
        return []
    return [a for a in (_norm(p) for p in str(patron).split(",")) if a]


def coincide(nombre, patron):
    """True si `nombre` (nombre de archivo) calza con la palabra clave.

    Sin palabra clave no hay filtro: devuelve True para cualquier nombre."""
    alts = alternativas(patron)
    if not alts:
        return True
    nombre_norm = _norm(nombre)
    return any(a in nombre_norm for a in alts)


def _merge(defaults, guardado):
    """Combina los defaults con lo guardado, respetando la estructura conocida."""
    salida = {}
    for clave, valor in defaults.items():
        if isinstance(valor, dict):
            sub = dict(valor)
            guardado_sub = guardado.get(clave) if isinstance(guardado, dict) else None
            if isinstance(guardado_sub, dict):
                for k, v in guardado_sub.items():
                    if k in sub and isinstance(v, str):
                        sub[k] = v.strip()
            salida[clave] = sub
        else:
            g = guardado.get(clave) if isinstance(guardado, dict) else None
            salida[clave] = g.strip() if isinstance(g, str) else valor
    return salida


def cargar_patrones(forzar=False):
    """Devuelve los patrones vigentes (defaults + patrones.json).

    Relee el archivo cuando cambia su fecha de modificación, así el servidor toma
    los cambios del panel sin reiniciarse."""
    global _cache, _cache_mtime
    with _lock:
        try:
            mtime = RUTA_PATRONES.stat().st_mtime if RUTA_PATRONES.is_file() else None
        except OSError:
            mtime = None
        if not forzar and _cache is not None and mtime == _cache_mtime:
            return _cache

        guardado = {}
        if mtime is not None:
            try:
                guardado = json.loads(RUTA_PATRONES.read_text(encoding="utf-8"))
            except (OSError, ValueError) as e:
                print(f"  ! patrones.json ilegible ({e}) → se usan los valores por defecto")
                guardado = {}

        _cache = _merge(PATRONES_DEFAULT, guardado)
        _cache_mtime = mtime
        return _cache


def guardar_patrones(nuevos):
    """Escribe patrones.json con los valores recibidos y devuelve el resultado combinado."""
    combinado = _merge(PATRONES_DEFAULT, nuevos or {})
    RUTA_PATRONES.write_text(
        json.dumps(combinado, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return cargar_patrones(forzar=True)


def restaurar_patrones():
    """Elimina patrones.json para volver a los valores de fábrica."""
    try:
        RUTA_PATRONES.unlink()
    except FileNotFoundError:
        pass
    return cargar_patrones(forzar=True)


def patron_excel(clave):
    """Palabra clave del Excel de una faena (o de SSO / Gestión Hídrica)."""
    return cargar_patrones()["excel"].get(clave, "")


def patron_word(clave):
    """Palabra clave del Word de una faena."""
    return cargar_patrones()["word"].get(clave, "")


def patron_otro(clave):
    """Palabra clave de los archivos que no dependen de una faena."""
    return cargar_patrones()["otros"].get(clave, "")


def fragmentos_ignorados():
    """Fragmentos de nombre cuyos vínculos externos deben omitirse."""
    return alternativas(cargar_patrones().get("ignorar_vinculos", ""))


def debe_ignorar_vinculo(nombre_archivo):
    """True si el vínculo externo apunta a un archivo que no debe actualizarse.

    Sin fragmentos configurados no se ignora nada (a diferencia de `coincide`)."""
    frags = fragmentos_ignorados()
    if not frags:
        return False
    nombre_norm = _norm(nombre_archivo)
    return any(f in nombre_norm for f in frags)


def archivos_que_calzan(carpeta, extension, patron):
    """Lista ordenada de archivos de `carpeta` con esa extensión que calzan con el patrón.

    Ignora los temporales de Office (`~$...`)."""
    carpeta = Path(carpeta) if carpeta else None
    if not carpeta or not carpeta.is_dir():
        return []
    ext = extension.lower()
    return sorted(
        f for f in carpeta.iterdir()
        if f.is_file()
        and f.suffix.lower() == ext
        and not f.name.startswith("~$")
        and coincide(f.name, patron)
    )
