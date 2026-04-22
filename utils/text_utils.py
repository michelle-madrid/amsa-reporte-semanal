"""Utilidades de limpieza, clasificación y normalización de texto."""

import re
import unicodedata
from state import errores

# Normaliza texto para facilitar su procesamiento posterior.
def normalizar_texto_clave(texto):
  texto = texto.strip()
  texto = re.sub(r"^[•·\-\s]+", "", texto)
  texto = unicodedata.normalize("NFKD", texto)
  texto = "".join(c for c in texto if not unicodedata.combining(c))
  return texto.lower()

# Implementa una parte específica de la lógica del informe.
def construir_texto_semana(dia_inicio, mes_inicio, dia_fin, mes_fin, year):
  meses = {
    "01": "enero",
    "02": "febrero",
    "03": "marzo",
    "04": "abril",
    "05": "mayo",
    "06": "junio",
    "07": "julio",
    "08": "agosto",
    "09": "septiembre",
    "10": "octubre",
    "11": "noviembre",
    "12": "diciembre",
  }

  dia_inicio_fmt = str(int(dia_inicio))          # "7", no "07"
  dia_fin_fmt    = str(int(dia_fin))
  mes_inicio_fmt = str(mes_inicio).zfill(2)      # "01" → necesario para el dict
  mes_fin_fmt    = str(mes_fin).zfill(2)

  nombre_mes_inicio = meses[mes_inicio_fmt]
  nombre_mes_fin = meses[mes_fin_fmt]

  return f"Semana del {dia_inicio_fmt} de {nombre_mes_inicio} al {dia_fin_fmt} de {nombre_mes_fin} {year}"

# Clasifica un subtítulo según las reglas de la faena.
def clasificar_subtitulo_ant(texto):
  t = normalizar_texto_clave(texto)

  if "desarrollo" in t and "extraccion" in t:
    return "Extracción a desarrollo"

  if "mineral" in t and "extraccion" in t:
    return "Extracción de Mineral"

  if "lastre" in t and "extraccion" in t:
    return "Extracción de lastre"

  if "remanejo" in t:
    return "Remanejo"

  if "extraccion mina" in t:
    return "Extracción Mina"

  return None

# Clasifica un subtítulo según las reglas de la faena.
def clasificar_subtitulo_cmz(texto):
  t = normalizar_texto_clave(texto)

  if "movimiento mina" in t:
    return "Movimiento Mina"

  if "extraccion mineral" in t:
    return "Extracción Mineral"

  if "extraccion lastre" in t or "extraccion esteril" in t:
    return "Extracción Lastre"

  if "remanejo" in t:
    return "Remanejo"

  if re.match(r"^fase\s+\S+", texto.strip(), flags=re.IGNORECASE):
    return "Fase"

  if "extraccion" in t:
    return "Extracción"

  return None

# Clasifica un subtítulo según las reglas de la faena.
def clasificar_subtitulo_cmz_planta(texto):
  t = normalizar_texto_clave(texto)

  if "mineral apilado hl" in t:
    return "Mineral Apilado HL"

  if "mineral beneficiado hl" in t:
    return "Mineral Beneficiado HL"

  if "ley apilado hl tcu" in t or "ley apilado hl" in t:
    return "Ley Apilado HL TCu"

  if "mineral apilado dl" in t:
    return "Mineral Apilado DL"

  if "mineral beneficiado dl" in t:
    return "Mineral Beneficiado DL"

  if "ley apilado dl tcu" in t or "ley apilado dl" in t:
    return "Ley Apilado DL TCu"

  if "remanejo ripios" in t:
    return "Remanejo Ripios"

  if re.match(r'^pls\b', t):
    return "PLS"

  if "cobre fino producido" in t:
    return "Cobre Fino Producido"

  return None

# Normaliza texto para facilitar su procesamiento posterior.
def normalizar_linea_ant(texto):
  original = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", texto).strip()
  clave = normalizar_texto_clave(original)

  # --- Movimiento Mina ---
  if "movimiento mina" in clave:
    cuerpo = re.sub(r"(?i)^.*?movimiento mina[:\s-]*", "", original).strip(" :.-")
    return f"Movimiento Mina: {cuerpo}" if cuerpo else "Movimiento Mina:"

  # --- Extracción Mina ---
  if "extraccion mina" in clave:
    cuerpo = re.sub(r"(?i)^.*?extracci[oó]n mina[:\s-]*", "", original).strip(" :.-")
    return f"Extracción Mina: {cuerpo}" if cuerpo else "Extracción Mina:"

  # --- Extracción de Mineral (incluye mayor/menor) ---
  if "extraccion de mineral" in clave or "extraccion mineral" in clave:
    cuerpo = re.sub(r"(?i)^.*?extracci[oó]n(\s+de)?\s+mineral[:\s-]*", "", original).strip(" :.-")
    return f"Extracción de Mineral: {cuerpo}" if cuerpo else "Extracción de Mineral:"

  if "mayor extraccion de mineral" in clave or "menor extraccion de mineral" in clave:
    cuerpo = re.sub(r"(?i)^.*?extracci[oó]n\s+de\s+mineral[:\s-]*", "", original).strip(" :.-")
    return f"Extracción de Mineral: {cuerpo}" if cuerpo else "Extracción de Mineral:"

  # --- Extracción de lastre (incluye mayor/menor) ---
  if "extraccion de lastre" in clave or "extraccion lastre" in clave:
    cuerpo = re.sub(r"(?i)^.*?extracci[oó]n(\s+de)?\s+lastre[:\s-]*", "", original).strip(" :.-")
    return f"Extracción de lastre: {cuerpo}" if cuerpo else "Extracción de lastre:"

  if "mayor extraccion de lastre" in clave or "menor extraccion de lastre" in clave:
    cuerpo = re.sub(r"(?i)^.*?extracci[oó]n\s+de\s+lastre[:\s-]*", "", original).strip(" :.-")
    return f"Extracción de lastre: {cuerpo}" if cuerpo else "Extracción de lastre:"

  # --- 🔥 Extracción a desarrollo (CLAVE) ---
  if "extraccion a desarrollo" in clave:
    cuerpo = re.sub(r"(?i)^.*?extracci[oó]n a desarrollo[:\s-]*", "", original).strip(" :.-")
    return f"Extracción a desarrollo: {cuerpo}" if cuerpo else "Extracción a desarrollo:"

  # --- Remanejo ---
  if "remanejo" in clave:
    cuerpo = re.sub(r"(?i)^.*?remanejo[:\s-]*", "", original).strip(" :.-")
    return f"Remanejo: {cuerpo}" if cuerpo else "Remanejo:"

  # --- fallback ---
  return original

# Normaliza texto para facilitar su procesamiento posterior.
def normalizar_linea_cmz_planta(texto):
  original = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", texto).strip()
  etiqueta = clasificar_subtitulo_cmz_planta(original)

  if not etiqueta:
    return original

  patrones = {
    "Mineral Apilado HL": r"(?i)^.*?mineral\s+apilado\s+hl[:\s-]*",
    "Mineral Beneficiado HL": r"(?i)^.*?mineral\s+beneficiado\s+hl[:\s-]*",
    "Ley Apilado HL TCu": r"(?i)^.*?ley\s+apilado\s+hl(?:\s+tcu)?[:\s-]*",
    "Mineral Apilado DL": r"(?i)^.*?mineral\s+apilado\s+dl[:\s-]*",
    "Mineral Beneficiado DL": r"(?i)^.*?mineral\s+beneficiado\s+dl[:\s-]*",
    "Ley Apilado DL TCu": r"(?i)^.*?ley\s+apilado\s+dl(?:\s+tcu)?[:\s-]*",
    "Remanejo Ripios": r"(?i)^.*?remanejo\s+ripios[:\s-]*",
    "PLS": r"(?i)^.*?\bpls\b[:\s-]*",
    "Cobre Fino Producido": r"(?i)^.*?cobre\s+fino\s+producido[:\s-]*",
  }

  texto_limpio = re.sub(patrones[etiqueta], "", original).strip(" :.-")
  resultado = f"{etiqueta}: {texto_limpio}" if texto_limpio else f"{etiqueta}:"
  resultado = limpiar_texto_global(resultado)

  if original != resultado:
    errores.append(f"CMZ - Planta normalizado: '{original}' → '{resultado}'")

  return resultado

# Normaliza texto para facilitar su procesamiento posterior.
def normalizar_linea_cmz(texto):
  original = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", texto).strip()
  etiqueta = clasificar_subtitulo_cmz(original)

  if not etiqueta:
    return original

  if etiqueta == "Movimiento Mina":
    texto_limpio = re.sub(r"(?i)^.*?movimiento mina[:\s-]*", "", original).strip(" :.-")
    resultado = f"Movimiento Mina: {texto_limpio}" if texto_limpio else "Movimiento Mina:"

  elif etiqueta == "Extracción":
    texto_limpio = re.sub(r"(?i)^.*?(total\s+extracci[oó]n|extracci[oó]n)[:\s-]*", "", original).strip(" :.-")
    resultado = f"Extracción: {texto_limpio}" if texto_limpio else "Extracción:"

  elif etiqueta == "Extracción Mineral":
    texto_limpio = re.sub(r"(?i)^.*?extracci[oó]n\s+mineral[:\s-]*", "", original).strip(" :.-")
    resultado = f"Extracción Mineral: {texto_limpio}" if texto_limpio else "Extracción Mineral:"

  elif etiqueta == "Extracción Lastre":
    texto_limpio = re.sub(r"(?i)^.*?(extracci[oó]n\s+lastre|extracci[oó]n\s+est[eé]ril)[:\s-]*", "", original).strip(" :.-")
    resultado = f"Extracción Lastre: {texto_limpio}" if texto_limpio else "Extracción Lastre:"

  elif etiqueta == "Remanejo":
    texto_limpio = re.sub(r"(?i)^.*?remanejo[:\s-]*", "", original).strip(" :.-")
    resultado = f"Remanejo: {texto_limpio}" if texto_limpio else "Remanejo:"

  elif etiqueta == "Fase":
    match = re.match(r"(?i)^(fase\s+\S+(?:\s+y\s+\S+)?)\s*(.*)$", original)
    if match:
      encabezado = match.group(1).strip()
      resto = match.group(2).strip(" :.-")
      resultado = f"{encabezado}: {resto}" if resto else f"{encabezado}:"
    else:
      resultado = original

  else:
    resultado = original

  resultado = limpiar_texto_global(resultado)

  if original != resultado:
    errores.append(f"CMZ - Mina normalizado: '{original}' → '{resultado}'")

  return resultado


# Normaliza texto para facilitar su procesamiento posterior.
def _quitar_dos_puntos_inicio(texto):
  return re.sub(r"^([^:]+):\s*", r"\1 ", texto)

# Convierte separadores decimales de coma a punto, distinguiendo de separadores de miles.
# Reglas:
#   X,YYY  (exactamente 3 dígitos) → separador de miles → no tocar
#   X,Y    (1-2 dígitos)           → decimal            → reemplazar por punto
#   X,YYYY (4+ dígitos)            → inusual            → imprimir para revisión
def normalizar_decimales(texto):
  if not texto:
    return texto

  def reemplazar(match):
    entero = match.group(1)
    decimal = match.group(2)
    n = len(decimal)
    original = match.group(0)

    if n == 3:
      # separador de miles → no tocar
      return original
    elif n <= 2:
      # decimal con 1-2 dígitos → convertir
      return f"{entero}.{decimal}"
    else:
      # 4+ dígitos → inusual, imprimir para revisión
      return original

  return re.sub(r'(\d+),(\d+)(?!\d)', reemplazar, texto)

# Normaliza texto globalmente para homogeneizar formato y capitalización.
def limpiar_texto_global(texto):
  if not texto:
    return texto

  texto = str(texto).strip()

  # normalizar tabulaciones a espacio simple
  texto = re.sub(r'\t+', ' ', texto)

  # colapsar doble (o más) dos puntos antes de normalizar espaciado
  texto = re.sub(r"::+", ":", texto)
  # asegurar un solo espacio después de ":"
  texto = re.sub(r":(?!\s)", ": ", texto)
  texto = re.sub(r":\s{2,}", ": ", texto)

  # corregir formato de hora: "15: 00" → "15:00" (debe ir después de normalizar espacios)
  texto = re.sub(r'\b(\d{1,2}):\s+(\d{2})\b', r'\1:\2', texto)

  # normalizar separadores decimales de coma a punto
  texto = normalizar_decimales(texto)

  # normalizar "plan mensual" (cualquier capitalización) → "Plan Mensual"
  texto = re.sub(r'(?i)\bplan\s+mensual\b', 'Plan Mensual', texto)

  # normalizar "alto potencial (aap)" (cualquier capitalización) → "Alto Potencial (AAP)"
  texto = re.sub(r'(?i)\balto\s+potencial\s*\(\s*aap\s*\)', 'Alto Potencial (AAP)', texto)

  # eliminar ceros a la izquierda en números enteros (ej: 07 → 7), sin tocar decimales ni horas
  # (?<!:) evita borrar el cero de minutos en horas como 15:00 o 08:30
  texto = re.sub(r'(?<!:)\b0+(\d+)\b', r'\1', texto)

  # pasar a minúscula Bajo / Sobre / En línea cuando NO van al inicio real del texto
  def bajar_estado(match):
    return match.group(0).lower()

  primera_palabra_match = re.match(r"^\S+", texto)

  if primera_palabra_match:
    fin_primera = primera_palabra_match.end()
    inicio = texto[:fin_primera]
    resto = texto[fin_primera:]

    resto = re.sub(r"\bBajo\b", bajar_estado, resto)
    resto = re.sub(r"\bSobre\b", bajar_estado, resto)
    resto = re.sub(r"\bEn línea\b", bajar_estado, resto)

    texto = inicio + resto
  else:
    texto = re.sub(r"\bBajo\b", bajar_estado, texto)
    texto = re.sub(r"\bSobre\b", bajar_estado, texto)
    texto = re.sub(r"\bEn línea\b", bajar_estado, texto)

  # asegurar que las líneas de contenido terminen con . o :
  # Solo aplica a líneas con contenido real (tienen ": " = título + cuerpo), no a títulos sueltos
  texto = texto.rstrip()
  if texto and texto[-1] == ",":
    texto = texto[:-1] + "."
  elif texto and texto[-1] not in (".", ":", ";") and ": " in texto:
    texto = texto + "."

  return texto

# En líneas de Ley, elimina el valor absoluto del paréntesis dejando solo la desviación %.
# Ejemplo: (+0.080%; +27.1%) → (+27.1%)
def limpiar_parentesis_ley(texto):
  return re.sub(r'\(([^;)]+);\s*([^)]+)\)', r'(\2)', texto)

# Ajusta etiquetas específicas de MLP para dejarlas en el formato esperado.
def limpiar_texto_mlp(texto):
    reemplazos = {
        "Total Extracción": "Extracción",
        "Extracción Estéril": "Extracción Lastre",
        "Extracción Mineral": "Extracción Mineral",
    }
    nuevo_texto = texto
    for original, nuevo in reemplazos.items():
        if original in texto:
            nuevo_texto = texto.replace(original, nuevo)
    return nuevo_texto
