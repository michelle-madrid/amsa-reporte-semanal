"""Funciones de extracción de texto desde Word y desde bloques internos del informe."""

from docx import Document
import re

# Lee un documento Word y devuelve su texto completo.
def extraer_texto_word(ruta_word):
    try:
        doc = Document(ruta_word)
        texto = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        return texto
    except Exception as e:
        errores.append(f"[ERROR] No se pudo leer el informe {ruta_word}: {e}")
        return ""

# Extrae un bloque de texto entre un título inicial y uno final.
def extraer_bloque(texto, inicio, finales=()):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        l = linea.strip()
        if not capturar:
            if inicio in l:
                capturar = True
                continue
        else:
            if finales and any(l.startswith(f) or l == f for f in finales):
                break
            if l:
                seccion.append(l)
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_accidentabilidad(texto):
    return extraer_bloque(texto, inicio="Accidentabilidad", finales=("Reportabilidad",))

# Extrae información específica desde el texto o archivo de origen.
def extraer_reportabilidad(texto):
    return extraer_bloque(
        texto,
        inicio="Reportabilidad",
        finales=(
            "Medio Ambiente",
            "Gestión SSO",
            "Salud Ocupacional y Gestión Vial",
            "Producción Semana",
        ),
    )

# Extrae información específica desde el texto o archivo de origen.
def extraer_medio_ambiente(texto):
    seccion = []
    capturar = False
    finales = ("Asuntos Públicos", "Gestión SSO", "Producción Semana")
    for linea in texto.split("\n"):
        l = linea.strip()
        if not capturar:
            if "Medio Ambiente" in l:
                capturar = True
                continue
        else:
            if any(l.startswith(f) for f in finales):
                break
            if l:
                l_limpia = re.sub(r"^[•\-\·\s]*", "", l)
                seccion.append(l_limpia)
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_asuntos_publicos(texto):
    return extraer_bloque(texto, inicio="Asuntos Públicos", finales=("Producción Semana",))

# Extrae información específica desde el texto o archivo de origen.
def extraer_gestion_sso(texto):
    return extraer_bloque(
        texto,
        inicio="Gestión SSO",
        finales=(
            "Salud Ocupacional y Gestión Vial",
            "Producción Semana",
            "Medio Ambiente",
        ),
    )

# Extrae información específica desde el texto o archivo de origen.
def extraer_salud_ocupacional(texto):
    return extraer_bloque(texto, inicio="Salud Ocupacional y Gestión Vial", finales=("Medio Ambiente",))

# Extrae información específica desde el texto o archivo de origen.
def extraer_principales_desviaciones(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        if linea.startswith("Principales Desviaciones"):
            capturar = True
            continue
        if capturar:
            if linea.startswith("Mina") or linea.startswith("Tren"):
                break
            seccion.append(linea.strip())
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_mina(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        if linea.startswith("Mina"):
            capturar = True
            continue
        if capturar:
            if linea.startswith("Concentradora") or linea.startswith("Sulfuros") or linea.startswith("Detalle por fases") or linea.startswith("Planta"):
                break
            seccion.append(linea.strip())
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_concentradora(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        if linea.startswith("Concentradora"):
            capturar = True
            continue
        if capturar:
            if linea.startswith("Planta Desaladora"):
                break
            seccion.append(linea.strip())
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_sulfuros(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        if linea.startswith("Sulfuros"):
            capturar = True
            continue
        if capturar:
            if linea.startswith("Cátodos"):
                break
            seccion.append(linea.strip())
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_cátodos(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        if linea.startswith("Cátodos"):
            capturar = True
            continue
        if capturar:
            seccion.append(linea.strip())
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_detalle_fases(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        if linea.startswith("Detalle por fases"):
            capturar = True
            continue
        if capturar:
            if linea.startswith("Planta") or linea.startswith("Planta:"):
                break
            seccion.append(linea.strip())
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_planta(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        linea_limpia = linea.strip()
        if linea_limpia.startswith("Planta:") or linea_limpia.startswith("Planta"):
            capturar = True
            continue
        if capturar:
            seccion.append(linea_limpia)
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_planta_desaladora(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        if linea.strip() == "Planta Desaladora":
            capturar = True
            continue
        if capturar:
            if linea.startswith("Gestión Hídrica"):
                break
            seccion.append(linea.strip())
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_gestión_hídrica(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        if linea.strip() == "Gestión Hídrica":
            capturar = True
            continue
        if capturar:
            seccion.append(linea.strip())
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_tren(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        if linea.strip() == "Tren" or linea.strip() == "Tren:":
            capturar = True
            continue
        if capturar:
            if linea.startswith("Camión") or linea.startswith("Camión:"):
                break
            seccion.append(linea.strip())
    return seccion

# Extrae información específica desde el texto o archivo de origen.
def extraer_camión(texto):
    seccion = []
    capturar = False
    for linea in texto.split("\n"):
        if linea.strip() == "Camión" or linea.strip() == "Camión:":
            capturar = True
            continue
        if capturar:
            seccion.append(linea.strip())
    return seccion
