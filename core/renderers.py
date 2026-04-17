"""Renderizadores y procesadores por faena para construir el informe Word."""

import os
import re
import unicodedata

from config import CONFIG_COMPANIAS, INCLUIR_ESTADO_FASES_DESARROLLO, ORDEN_PRINCIPALES_DESVIACIONES, NIVEL_BASE_POR_SECCION, NIVEL_POR_COMPANIA_SECCION_SUBTITULO
from state import errores
from utils.text_utils import *
from utils.text_utils import _quitar_dos_puntos_inicio
from utils.word_utils import *
from utils.excel_utils import exportar_imagen_excel
from core.extractores import *


# Orquesta la construcción del bloque de una faena usando su procesador correspondiente.
def construir_bloque_faena(doc, clave, texto_word, excel_madre, orden_secciones=None):
    """Orquesta la creación de la sección de una compañía (delega en la función especialista)."""
    proc = PROCESADORES_FAENA.get(clave)
    if proc:
        proc(doc, texto_word, excel_madre)

# Renderiza contenido específico dentro del documento Word.
def mlp_render_medio_ambiente(doc, lineas):
  subtitulo_actual = None
  dentro_de_fecha = False
  patron_fecha = re.compile(r"^\d{1,2}\sde\s\w+\sde\s\d{4}")

  for linea in lineas:
    texto = linea.strip()
    if not texto:
      continue

    texto_limpio = re.sub(r"^[•·\-\s]+", "", texto).strip()
    texto_limpio = limpiar_texto_global(texto_limpio)

    if texto_limpio.startswith("Fuente:") or texto_limpio.startswith("Nota:"):
      p = doc.add_paragraph(style="Normal AMSA")
      p.paragraph_format.left_indent = Cm(1.27)
      p.paragraph_format.first_line_indent = Cm(0)
      p.paragraph_format.line_spacing = 1.0
      p.paragraph_format.space_before = Pt(0)
      p.paragraph_format.space_after = Pt(6)
      run = p.add_run(texto_limpio)
      run.font.name = "Arial"
      run.font.size = Pt(11)
      subtitulo_actual = None
      dentro_de_fecha = False
      continue

    es_subtitulo = (
      ":" not in texto_limpio
      and not texto_limpio.endswith(".")
      and len(texto_limpio) <= 60
    )

    if es_subtitulo:
      agregar_viñeta(doc, texto_limpio, nivel=2, espacio_despues=6)
      subtitulo_actual = texto_limpio
      dentro_de_fecha = False
      continue

    if texto_limpio.startswith("Calidad del aire"):
      dentro_de_fecha = False
      nivel_calidad = 3 if subtitulo_actual else 2
      agregar_viñeta_sin_negrita(doc, texto_limpio, nivel=nivel_calidad, espacio_despues=6)
      continue

    match_fecha = patron_fecha.match(texto_limpio)

    if match_fecha:
      # Línea con fecha: círculo del template + fecha en negrita
      dentro_de_fecha = True
      nivel_fecha = 3 if subtitulo_actual else 2
      agregar_viñeta_fecha_inicial(doc, texto_limpio, nivel=nivel_fecha, espacio_despues=6)

    elif dentro_de_fecha:
      # Párrafo de continuación tras una fecha: alineado con el texto del nivel
      p = doc.add_paragraph(style="Normal AMSA")
      p.paragraph_format.line_spacing = 1.0
      p.paragraph_format.space_before = Pt(0)
      p.paragraph_format.space_after = Pt(6)
      left = Cm(3.0) if subtitulo_actual else Cm(1.9)
      p.paragraph_format.left_indent = left
      p.paragraph_format.first_line_indent = Cm(0)

      run = p.add_run(texto_limpio)
      run.bold = False
      run.font.name = "Arial"
      run.font.size = Pt(11)

    else:
      # Línea sin fecha y fuera de bloque de fecha: círculo del template
      if subtitulo_actual:
        agregar_viñeta(doc, texto_limpio, nivel=3, espacio_despues=6)
      else:
        agregar_viñeta(doc, texto_limpio, nivel=2, espacio_despues=6)

# Inserta la sección de estado de fases de desarrollo con imagen y criterios.
def agregar_estado_fases_desarrollo(doc, excel_madre):
  doc.add_page_break()
  agregar_titulo(doc, "Estado de Fases de Desarrollo para medir adhesión al plan minero:", nivel=2)
  img_fases = exportar_imagen_excel(excel_madre, "Triger - D°Mina", "B2:S21", "estado_fases.png")
  agregar_imagen(doc, img_fases, 19, 3.12, "")

  agregar_texto(doc, "Criterios:", bold=True, color=(0x59, 0x66, 0x66))
  agregar_viñeta_color(doc, "Cumplimiento mayor o igual al 100%.", color_punto=(0x00, 0x80, 0x00))
  agregar_viñeta_color(doc, "Cumplimiento entre el rango [86%-99%].", color_punto=(0xFF, 0xFF, 0x00))
  agregar_viñeta_color(doc, "Cumplimiento menor o igual al 85%.", color_punto=(0xFF, 0x00, 0x00))

# ── Validación de consistencia cualitativa en 'Acumulado al mes/año' ─────────

_PAT_PCT_ACUM = re.compile(r'[+\-]?\d+(?:[.,]\d+)?%')
_UMBRAL_EN_LINEA_ACUM = 5.0  # |%| ≤ umbral → "en línea"; mayor magnitud → sobre/bajo PM

def _validar_clasificacion_acumulados(lineas_acum, clave):
    """
    Para cada línea 'Acumulado al mes/año', revisa que la etiqueta cualitativa
    ('en línea', 'mayor producción', 'menor producción') sea coherente con el
    porcentaje declarado en esa misma cláusula.
    """
    for linea in lineas_acum:
        t = normalizar_texto_clave(linea)
        tipo = "mes" if "acumulado al mes" in t else "año"

        # Separar por " y " para analizar cada afirmación de forma independiente
        clausulas = re.split(r"\s+y\s+", linea)
        for clausula in clausulas:
            tc = normalizar_texto_clave(clausula)
            pcts = _PAT_PCT_ACUM.findall(clausula)
            if not pcts:
                continue

            es_en_linea = "en linea" in tc
            es_mayor    = "mayor produccion" in tc or "mayor al" in tc
            es_menor    = "menor produccion" in tc or "baja produccion" in tc or "bajo pm" in tc

            for pct_str in pcts:
                try:
                    pct_val = float(pct_str.replace('%', '').replace(',', '.'))
                except ValueError:
                    continue

                if es_en_linea and abs(pct_val) > _UMBRAL_EN_LINEA_ACUM:
                    clasif = "SOBRE PM" if pct_val > 0 else "BAJO PM"
                    msg = (f"{clave} - Acumulado al {tipo}: dice 'en línea' "
                           f"pero % = {pct_str} → debería ser '{clasif}'")
                    print(f"[REVISAR] {msg}")
                    print(f"          → {linea.strip()}")
                    errores.append(msg)

                elif es_mayor and pct_val < 0:
                    msg = (f"{clave} - Acumulado al {tipo}: dice 'mayor producción' "
                           f"pero % = {pct_str} (negativo → revisar)")
                    print(f"[REVISAR] {msg}")
                    print(f"          → {linea.strip()}")
                    errores.append(msg)

                elif es_menor and pct_val > 0:
                    msg = (f"{clave} - Acumulado al {tipo}: dice 'menor producción' "
                           f"pero % = {pct_str} (positivo → revisar)")
                    print(f"[REVISAR] {msg}")
                    print(f"          → {linea.strip()}")
                    errores.append(msg)


# Valida que existan las líneas de acumulado mensual y anual en las principales desviaciones.
def validar_acumulados_principales_desviaciones(texto_compania, clave, es_seleccionada=True):
  extractores_por_compania = {
    "MLP": [
      extraer_mina,
      extraer_concentradora,
      extraer_planta_desaladora,
      extraer_gestión_hídrica,
    ],
    "CEN": [
      extraer_mina,
      extraer_sulfuros,
      extraer_cátodos,
    ],
    "ANT": [
      extraer_mina,
      extraer_detalle_fases,
      extraer_planta,
    ],
    "CMZ": [
      extraer_mina,
      extraer_planta,
    ],
    "FCAB": [
      extraer_tren,
      extraer_camión,
    ],
  }

  extractores = extractores_por_compania.get(clave, [])
  lineas = []

  for extractor in extractores:
    try:
      bloque = extractor(texto_compania)
      if bloque:
        lineas.extend(bloque)
    except Exception:
      pass

  lineas_limpias = [
    re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", l.strip()).lower()
    for l in lineas
    if l and l.strip()
  ]

  tiene_acum_mes = any("acumulado al mes" in l for l in lineas_limpias)
  tiene_acum_anio = any("acumulado al año" in l for l in lineas_limpias)

  if not tiene_acum_mes and es_seleccionada:
    print(f"[REVISAR] {clave} - Principales Desviaciones: no viene 'Acumulado al mes'")
    errores.append(f"{clave} - Principales Desviaciones: falta 'Acumulado al mes'")

  if not tiene_acum_anio and es_seleccionada:
    print(f"[REVISAR] {clave} - Principales Desviaciones: no viene 'Acumulado al año'")
    errores.append(f"{clave} - Principales Desviaciones: falta 'Acumulado al año'")

  # Validar formato de porcentajes (1 decimal exacto) solo en subsecciones seleccionadas
  # que realmente existen para esta compañía en ORDEN_PRINCIPALES_DESVIACIONES
  _SECCIONES_VALIDAR_PCT = {"Mina", "Planta", "Sulfuros", "Cátodos", "Detalle por fases", "Camión", "Tren"}
  _EXTRACTORES_PCT = {
    "Mina": extraer_mina,
    "Planta": extraer_planta,
    "Sulfuros": extraer_sulfuros,
    "Cátodos": extraer_cátodos,
    "Detalle por fases": extraer_detalle_fases,
    "Camión": extraer_camión,
    "Tren": extraer_tren,
  }
  secciones_compania = set(ORDEN_PRINCIPALES_DESVIACIONES.get(clave, {}).keys())
  lineas_pct = []
  for nombre in secciones_compania & _SECCIONES_VALIDAR_PCT:
    try:
      bloque = _EXTRACTORES_PCT[nombre](texto_compania)
      if bloque:
        lineas_pct.extend(bloque)
    except Exception:
      pass

  patron_pct = re.compile(r'[+\-]?\d+(?:[.,](\d+))?%')
  # Patrón para texto explicativo de referencia que puede tener cualquier precisión,
  # ej.: "(0.41% Real vs 0.40% PM)" o "(12.34% Real vs 11.00% PM)".
  # También cubre variantes sin espacios o con otro orden.
  _patron_ref = re.compile(
      r'\([^)]*\d+[.,]\d+%[^)]*(?:Real|real|PM|pm|Plan|plan)[^)]*\)',
      re.IGNORECASE,
  )
  casos_sin = []   # (valor, linea)
  casos_extra = [] # (valor, linea)
  for linea in lineas_pct:
    # Quitar segmentos explicativos "Real vs PM" antes de validar formato
    linea_check = _patron_ref.sub('', linea or "")
    for m in patron_pct.finditer(linea_check):
      d = m.group(1)
      if d is None:
        casos_sin.append((m.group(0), linea.strip()))
      elif len(d) != 1:
        casos_extra.append((m.group(0), linea.strip()))

  if casos_sin:
    if len(casos_sin) < 3:
      for val, linea in casos_sin:
        print(f"[REVISAR] {clave} - Porcentaje sin decimal: '{val}' en → {linea}")
        errores.append(f"{clave} - Porcentaje sin decimal: '{val}' en → {linea}")
    else:
      print(f"[REVISAR] {clave} - Porcentajes sin decimal: {len(casos_sin)} caso(s)")
      errores.append(f"{clave} - Porcentajes sin decimal: {len(casos_sin)} caso(s)")

  if casos_extra:
    if len(casos_extra) < 3:
      for val, linea in casos_extra:
        print(f"[REVISAR] {clave} - Porcentaje con decimales de más: '{val}' en → {linea}")
        errores.append(f"{clave} - Porcentaje con decimales de más: '{val}' en → {linea}")
    else:
      print(f"[REVISAR] {clave} - Porcentajes con más de 1 decimal: {len(casos_extra)} caso(s)")
      errores.append(f"{clave} - Porcentajes con más de 1 decimal: {len(casos_extra)} caso(s)")

  # Validar coherencia cualitativa en líneas 'Acumulado al mes/año'
  _limpiar = lambda l: re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", l.strip())
  lineas_acum = [
      _limpiar(l)
      for l in lineas
      if "acumulado al mes" in normalizar_texto_clave(_limpiar(l))
      or "acumulado al ano" in normalizar_texto_clave(_limpiar(l))
  ]
  if lineas_acum:
      _validar_clasificacion_acumulados(lineas_acum, clave)

# Renderiza contenido específico dentro del documento Word.
def _render_fecha_accidente(doc, fecha, texto_resto):
  """Renderiza fecha de accidente con estilo Viñeta 3 (círculo del template), fecha en negrita."""
  p = doc.add_paragraph(style="Viñeta 3")
  p.paragraph_format.line_spacing = 1.0
  p.paragraph_format.space_before = Pt(0)
  p.paragraph_format.space_after = Pt(6)
  run_fecha = p.add_run(f"{fecha}: ")
  run_fecha.bold = True
  run_fecha.font.name = "Arial"
  run_fecha.font.size = Pt(11)
  if texto_resto:
    run_resto = p.add_run(texto_resto)
    run_resto.bold = False
    run_resto.font.name = "Arial"
    run_resto.font.size = Pt(11)

def mlp_render_accidentabilidad(doc, lineas):
  patron_fecha_embebida = re.compile(r"\(Fecha:\s*(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})\):\s*")
  patron_fecha_inicio   = re.compile(r"^(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})")

  for linea in lineas:
    linea = linea.strip()
    if not linea:
      continue

    limpia = re.sub(r"^[•○o·\s]+", "", linea).strip()

    # Formato 1: "(Fecha: DD de mes de YYYY): texto"
    match_emb = patron_fecha_embebida.search(linea)
    if match_emb:
      fecha = match_emb.group(1)
      before = linea[:match_emb.start()].strip().rstrip(" ,")
      after  = linea[match_emb.end():].strip()
      texto_resto = f"{before}: {after}" if before and after else (before or after)
      _render_fecha_accidente(doc, fecha, texto_resto)
      continue

    # Formato 2: "DD de mes de YYYY: texto" al inicio de línea
    match_ini = patron_fecha_inicio.match(limpia)
    if match_ini:
      fecha = match_ini.group(1)
      resto = limpia[len(fecha):].lstrip(": ").strip()
      _render_fecha_accidente(doc, fecha, resto)
      continue

    if re.match(r"^Nota\s*:", limpia, re.IGNORECASE):
      agregar_nota_sin_viñeta(doc, limpia)
    else:
      agregar_viñeta(doc, limpia, nivel=2, espacio_despues=6)

# Renderiza accidentabilidad para todas las compañías excepto MLP.
# Las líneas de fecha (nivel 3) absorben líneas de continuación que empiezan con "-".
def render_accidentabilidad_generica(doc, lineas):
  # Acepta "06 de abril de 2026" y "06 de abril 2026" (con o sin "de" antes del año)
  patron_fecha = re.compile(r"^\d{1,2}\s+de\s+\w+(\s+de)?\s+\d{4}")
  pendiente = None  # {'fecha': str, 'texto': str}

  def _flush():
    nonlocal pendiente
    if pendiente is None:
      return
    p = doc.add_paragraph(style="Viñeta 3")
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    run_f = p.add_run(f"{pendiente['fecha']}: ")
    run_f.bold = True
    run_f.font.name = "Arial"
    run_f.font.size = Pt(11)
    run_t = p.add_run(pendiente['texto'])
    run_t.bold = False
    run_t.font.name = "Arial"
    run_t.font.size = Pt(11)
    pendiente = None

  for linea in lineas:
    linea = linea.strip()
    if not linea:
      continue
    limpia = re.sub(r"^[•○o·\s]+", "", linea).strip()
    m = patron_fecha.match(limpia)
    if m:
      _flush()
      fecha = m.group(0)
      resto = limpia[len(fecha):].lstrip(": ").strip()
      pendiente = {'fecha': fecha, 'texto': resto}
    elif pendiente is not None and limpia.startswith("-"):
      cont = limpia.lstrip("- ").strip()
      pendiente['texto'] = (pendiente['texto'] + " " + cont).strip()
    else:
      _flush()
      if re.match(r"^Nota\s*:", limpia, re.IGNORECASE):
        agregar_nota_sin_viñeta(doc, limpia)
      else:
        agregar_viñeta(doc, limpia, nivel=2, espacio_despues=6)

  _flush()

# Renderiza contenido específico dentro del documento Word.
def _ap_render_subtitulo(doc, texto):
  p = doc.add_paragraph(style="Normal AMSA")
  p.paragraph_format.left_indent = Cm(1.27)
  p.paragraph_format.first_line_indent = Cm(0)
  p.paragraph_format.line_spacing = 1.0
  p.paragraph_format.space_before = Pt(6)
  p.paragraph_format.space_after = Pt(3)
  run = p.add_run(texto)
  run.bold = False
  run.font.name = "Arial"
  run.font.size = Pt(11)

def mlp_render_asuntos_publicos(doc, lineas):
  # 1. Limpiar: quitar marcadores explícitos y aplicar limpieza global
  textos = []
  for l in lineas:
    raw = l.strip()
    if not raw:
      continue
    t = re.sub(r'^(\d+[.)]\s*|[•·○o]\s*)', '', raw).strip()
    t = limpiar_texto_global(t)
    if t:
      textos.append(t)
  if not textos:
    return

  # 2. Helpers de clasificación
  def _es_bullet_cand(t):
    # Termina en ":" y es corto: encabezado de sección (bullet negro)
    return t.endswith(':') and len(t) <= 80 and t.count(':') == 1

  def _es_sub_cand(t):
    # Sin ":" ni ".", corto: candidato a subtítulo sin viñeta
    return ':' not in t and not t.endswith('.') and len(t) <= 70

  # 3. Clasificar con lookahead
  # Un candidato a subtítulo ES subtítulo si la siguiente línea es bullet o subtítulo
  tipos = []
  for i, t in enumerate(textos):
    if _es_bullet_cand(t):
      tipos.append('bullet')
    elif _es_sub_cand(t):
      tipos.append('sub')
    else:
      tipos.append('circulo')

  # 4. Si hay bullets, es estructura de 3 niveles; si no, 2 niveles
  hay_bullets = 'bullet' in tipos

  # 5. Renderizar — solo círculos blancos del template (Viñeta 2 / Viñeta 3), sin negritas
  for texto, tipo in zip(textos, tipos):
    if tipo == 'sub':
      _ap_render_subtitulo(doc, texto)
    elif tipo == 'bullet':
      # En estructura 3 niveles: Viñeta 2 (círculo intermedio)
      # En estructura 2 niveles: Viñeta 2 (único nivel de círculo)
      agregar_viñeta_sin_negrita(doc, texto, nivel=2, espacio_despues=6)
    else:
      # circulo: nivel 3 si hay bullets (estructura 3 niveles), nivel 2 si no
      nivel = 3 if hay_bullets else 2
      agregar_viñeta_sin_negrita(doc, texto, nivel=nivel, espacio_despues=6)

# Renderiza contenido específico dentro del documento Word.
def cen_render_catodos(doc, texto_compania, excel_madre=None):
  contenido = [linea.strip() for linea in extraer_cátodos(texto_compania) if linea.strip()]
  if not contenido:
    return

  agregar_texto(doc, "Cátodos:", bold=True, color=(0x00, 0x77, 0x8B))

  linea_total = None
  bloque_met = []
  bloque_oxe = []
  bloque_actual = None

  for linea in contenido:
    texto = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", linea).strip()

    if texto.startswith("Producción total de cátodos de Cu"):
      linea_total = texto
      bloque_actual = None
      continue

    if texto.startswith("Planta Hidro MET"):
      bloque_actual = "MET"
      continue

    if texto.startswith("Planta Hidro OXE"):
      bloque_actual = "OXE"
      continue

    if bloque_actual == "MET":
      bloque_met.append(texto)
    elif bloque_actual == "OXE":
      bloque_oxe.append(texto)

  if linea_total:
    agregar_bullet_negro_manual(
      doc,
      linea_total,
      left_indent_cm=1.27,
      bullet_indent_cm=0.85,
      espacio_despues=6,
      bold=True
    )
  else:
    print("[REVISAR] CEN - Cátodos: no se encontró la línea 'Producción total de cátodos de Cu'.")
    errores.append("[REVISAR] CEN - Cátodos: no se encontró la línea 'Producción total de cátodos de Cu'.")

  if bloque_met:
    agregar_texto_subrayado(doc, "Planta Hidro MET", left_indent_cm=0.85, espacio_despues=6, bold=True)
    for linea in bloque_met:
      texto_base = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", linea).strip()

      if (
        texto_base.startswith("Acumulado al mes")
        or texto_base.startswith("Acumulado al año")
        or texto_base.startswith("Respecto del Plan")
      ):
        agregar_linea_acumulado(doc, texto_base)
      else:
        agregar_viñeta(doc, linea, nivel=2, espacio_despues=6)

  if bloque_oxe:
    agregar_texto_subrayado(doc, "Planta Hidro OXE", left_indent_cm=0.85, espacio_despues=6, bold=True)
    for linea in bloque_oxe:
      texto_base = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", linea).strip()

      if (
        texto_base.startswith("Acumulado al mes")
        or texto_base.startswith("Acumulado al año")
        or texto_base.startswith("Respecto del Plan")
      ):
        agregar_linea_acumulado(doc, texto_base)
      else:
        agregar_viñeta(doc, linea, nivel=2, espacio_despues=6)

# Renderiza contenido específico dentro del documento Word.
def ant_render_medio_ambiente(doc, lineas):
  dentro_subgrupo = False

  for linea in lineas:
    texto = linea.strip()
    if not texto:
      continue

    if texto.startswith("Fuente:") or texto.startswith("Nota:"):
      p = doc.add_paragraph(style="Normal AMSA")
      p.paragraph_format.left_indent = Cm(1.27)
      p.paragraph_format.first_line_indent = Cm(0)
      p.paragraph_format.line_spacing = 1.0
      p.paragraph_format.space_before = Pt(0)
      p.paragraph_format.space_after = Pt(6)

      run = p.add_run(texto)
      run.font.name = "Arial"
      run.font.size = Pt(11)

      dentro_subgrupo = False
      continue

    if texto.endswith(":"):
      agregar_viñeta_plana(doc, texto, nivel=2, espacio_despues=6)
      dentro_subgrupo = True
      continue

    if dentro_subgrupo:
      agregar_viñeta_plana(doc, texto, nivel=3, espacio_despues=6)
    else:
      agregar_viñeta_plana(doc, texto, nivel=2, espacio_despues=6)

# Renderiza contenido específico dentro del documento Word.
def ant_render_mina(doc, texto_compania, excel_madre=None):
  contenido = [linea.strip() for linea in extraer_mina(texto_compania) if linea.strip()]
  if not contenido:
    return

  agregar_texto(doc, "Mina:", bold=True, color=(0x00, 0x77, 0x8B))

  movimiento_mina = None
  bloques_nivel_2 = []
  detalles_extraccion_mina = []
  otros = []

  for texto in contenido:
    t = texto.strip()
    clave = normalizar_texto_clave(t)

    if clave.startswith("movimiento mina"):
      movimiento_mina = t

    elif (
      clave.startswith("extraccion mina")
      or clave.startswith("extraccion de mineral")
      or clave.startswith("extraccion de lastre")
      or clave.startswith("remanejo")
      or clave.startswith("extraccion a desarrollo")
      or clave.startswith("mayor extraccion de mineral")
      or clave.startswith("menor extraccion de mineral")
      or clave.startswith("mayor extraccion de lastre")
      or clave.startswith("menor extraccion de lastre")
    ):
      bloques_nivel_2.append(normalizar_linea_ant(t))

    elif re.match(r"^(Pala|Cargador)", t):
      detalles_extraccion_mina.append(t)

    else:
      otros.append(t)

  if movimiento_mina:
    agregar_viñeta_full_bold(doc, movimiento_mina, nivel=1, espacio_despues=6)
  else:
    print("[REVISAR] ANT - Mina: no se encontró 'Movimiento Mina' en formato esperado.")
    errores.append("ANT - Mina: falta 'Movimiento Mina'")

  for bloque in bloques_nivel_2:
    agregar_viñeta_con_titulo(doc, bloque, nivel=2, espacio_despues=6)

    if normalizar_texto_clave(bloque).startswith("extraccion mina"):
      for detalle in detalles_extraccion_mina:
        agregar_viñeta_plana(doc, detalle, nivel=3, espacio_despues=6)

  for texto in otros:
    print(f"[REVISAR] ANT - Mina: línea no clasificada dentro de subtítulos esperados -> '{texto}'")
    errores.append(f"ANT - Mina: línea no clasificada -> '{texto}'")
    agregar_viñeta_plana(doc, texto, nivel=2, espacio_despues=6)

# Procesa una sección genérica aplicando su extractor y renderizador.
def procesar_seccion(doc, texto_compania, nombre_compania, nombre_seccion, orden_subtitulos, excel_madre=None):
    nivel_base = NIVEL_BASE_POR_SECCION.get(nombre_seccion, 2)
    overrides_seccion = NIVEL_POR_COMPANIA_SECCION_SUBTITULO.get(nombre_compania, {}).get(nombre_seccion, {})

    extractores = {
        "Principales Desviaciones": extraer_principales_desviaciones,
        "Mina": extraer_mina,
        "Detalle por fases": extraer_detalle_fases,
        "Planta": extraer_planta,
        "Sulfuros": extraer_sulfuros,
        "Cátodos": extraer_cátodos,
        "Concentradora": extraer_concentradora,
        "Planta Desaladora": extraer_planta_desaladora,
        "Gestión Hídrica": extraer_gestión_hídrica,
        "Tren": extraer_tren,
        "Camión": extraer_camión,
    }

    extractor = extractores.get(nombre_seccion)
    if not extractor:
        return

    contenido = extractor(texto_compania)
    if not contenido:
        return

    agregar_texto(doc, f"{nombre_seccion}:", bold=True, color=(0x00, 0x77, 0x8B))

    if orden_subtitulos and orden_subtitulos[0] == "?":
        for linea in contenido:
            texto_base = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", linea.strip()).strip()
            if (
                texto_base.startswith("Acumulado al mes")
                or texto_base.startswith("Acumulado al año")
                or texto_base.startswith("Respecto del Plan")
            ):
                agregar_linea_acumulado(doc, texto_base)
            else:
                agregar_texto(doc, linea)
        return

    if not orden_subtitulos or orden_subtitulos == [""]:
        for linea in contenido:
            texto_limpio = linea.strip()
            texto_base = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", texto_limpio).strip()

            if texto_base.lower().startswith("ley ") or texto_base.lower().startswith("recuperaci"):
                linea = limpiar_parentesis_ley(linea)
                texto_base = limpiar_parentesis_ley(texto_base)

            if (
                texto_base.startswith("Acumulado al mes")
                or texto_base.startswith("Acumulado al año")
                or texto_base.startswith("Respecto del Plan")
            ):
                agregar_linea_acumulado(doc, texto_base)
            else:
                if nombre_compania == "ANT" and nombre_seccion == "Planta":
                    agregar_viñeta_inicio_negrita(doc, linea, nivel=nivel_base, espacio_despues=6)
                else:
                    agregar_viñeta(doc, linea, nivel=nivel_base, espacio_despues=6)
        return

    grupos = {sub: [] for sub in orden_subtitulos}
    subtitulo_actual = None
    contenido_extra = []
    orden_subtitulos_match = sorted(orden_subtitulos, key=len, reverse=True)

    for linea in contenido:
        texto = linea.strip()
        if not texto:
            continue
        texto_norm = texto.lower()
        match = None
        for sub in orden_subtitulos_match:
            sub_norm = sub.lower()
            if texto_norm.startswith(sub_norm) or texto_norm.startswith(sub_norm + ":"):
                match = sub
                break
        if match:
            subtitulo_actual = match
            grupos[subtitulo_actual].append(texto)
        else:
            if subtitulo_actual:
                grupos[subtitulo_actual].append(texto)
            else:
                contenido_extra.append(texto)

    for texto in contenido_extra:
        agregar_viñeta(doc, texto, nivel=nivel_base, espacio_despues=6)

    for subtitulo in orden_subtitulos:
        lineas = grupos.get(subtitulo, [])
        if not lineas:
            continue

        nivel_subtitulo = overrides_seccion.get(subtitulo, nivel_base)

        primera = True
        for texto in lineas:
            texto_base = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", texto.strip()).strip()

            if texto_base.lower().startswith("ley ") or texto_base.lower().startswith("recuperaci"):
                texto = limpiar_parentesis_ley(texto)
                texto_base = limpiar_parentesis_ley(texto_base)

            if (
                texto_base.startswith("Acumulado al mes")
                or texto_base.startswith("Acumulado al año")
                or texto_base.startswith("Respecto del Plan")
            ):
                agregar_linea_acumulado(doc, texto_base)
                primera = False
                continue

            nivel_actual = nivel_subtitulo if primera else min(nivel_subtitulo + 1, 4)

            if nombre_compania == "CEN" and nombre_seccion == "Mina" and primera:
                agregar_viñeta_full_bold(doc, texto, nivel=nivel_actual, espacio_despues=6)
            elif nombre_compania == "CEN" and nombre_seccion == "Mina" and nivel_actual >= 3:
                agregar_viñeta_sin_negrita(doc, texto, nivel=nivel_actual, espacio_despues=6)
            else:
                agregar_viñeta(doc, texto, nivel=nivel_actual, espacio_despues=6)

            primera = False

# Renderiza contenido específico dentro del documento Word.
def cen_render_medio_ambiente(doc, lineas):
  subtitulo_actual = None

  for linea in lineas:
    texto = linea.strip()
    if not texto:
      continue

    texto = re.sub(r"^(o\s+|[•·\-\s]+)", "", texto).strip()
    texto = limpiar_texto_global(texto)

    if texto.startswith("Fuente:") or texto.startswith("Nota:"):
      p = doc.add_paragraph(style="Normal AMSA")
      p.paragraph_format.left_indent = Cm(1.27)
      p.paragraph_format.first_line_indent = Cm(0)
      p.paragraph_format.line_spacing = 1.0
      p.paragraph_format.space_before = Pt(0)
      p.paragraph_format.space_after = Pt(6)

      run = p.add_run(texto)
      run.font.name = "Arial"
      run.font.size = Pt(11)

      subtitulo_actual = None
      continue

    es_subtitulo = (
      ":" not in texto
      and not texto.endswith(".")
      and len(texto) <= 60
    )

    if es_subtitulo:
      agregar_circulo_blanco_manual(
        doc,
        texto.strip(),
        left_indent_cm=1.9,
        bullet_indent_cm=1.45,
        espacio_despues=6
      )
      subtitulo_actual = texto
      continue

    if subtitulo_actual:
      agregar_circulo_blanco_manual(
        doc,
        texto.strip(),
        left_indent_cm=3.0,
        bullet_indent_cm=2.55,
        espacio_despues=6
      )
    else:
      agregar_circulo_blanco_manual(
        doc,
        texto.strip(),
        left_indent_cm=1.9,
        bullet_indent_cm=1.45,
        espacio_despues=6
      )

# Renderiza contenido específico dentro del documento Word.
def mlp_render_mina(doc, texto_compania, excel_madre=None):
  contenido = [linea.strip() for linea in extraer_mina(texto_compania) if linea.strip()]
  if not contenido:
    print("[MLP][WARN] La sección 'Mina' viene vacía en el informe original.")
    return
  p_titulo = doc.add_paragraph(style="Normal AMSA")
  p_titulo.paragraph_format.space_before = Pt(6)
  p_titulo.paragraph_format.space_after = Pt(0)

  run_titulo = p_titulo.add_run("Mina:")
  run_titulo.bold = True
  run_titulo.font.color.rgb = RGBColor(0x00, 0x77, 0x8B)
  run_titulo.font.name = "Arial"
  run_titulo.font.size = Pt(11)

  movimiento_mina = None
  extraccion = None
  extraccion_lastre = None
  extraccion_mineral = None
  remanejo = None

  for texto in contenido:
    clave = normalizar_texto_clave(texto)

    if clave.startswith("movimiento mina"):
      movimiento_mina = texto
    elif clave.startswith("total extraccion") or clave.startswith("extraccion:"):
      extraccion = re.sub(r"^Total Extracción", "Extracción", texto, count=1)
    elif clave.startswith("extraccion esteril"):
      extraccion_lastre = re.sub(r"^Extracción Estéril", "Extracción Lastre", texto, count=1)
    elif clave.startswith("extraccion mineral"):
      extraccion_mineral = texto
    elif clave.startswith("remanejo"):
      remanejo = texto

  if not movimiento_mina:
    print("[REVISAR][MLP] Corregir: No se encontró la línea 'Movimiento Mina' en el informe original.")

  if movimiento_mina:
    movimiento_mina = limpiar_texto_global(movimiento_mina)
    partes = movimiento_mina.split(":", 1)
    p = doc.add_paragraph(style="Viñeta 1")
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)

    run_bold = p.add_run(partes[0].strip() + ": ")
    run_bold.bold = True
    run_bold.font.name = "Arial"
    run_bold.font.size = Pt(11)

    if len(partes) > 1:
      run_normal = p.add_run(partes[1].strip())
      run_normal.font.name = "Arial"
      run_normal.font.size = Pt(11)

  if extraccion:
    agregar_viñeta(doc, extraccion, nivel=2, espacio_despues=6)

    if extraccion_lastre:
      agregar_viñeta(doc, _quitar_dos_puntos_inicio(extraccion_lastre), nivel=3, espacio_despues=6)

    if extraccion_mineral:
      agregar_viñeta(doc, _quitar_dos_puntos_inicio(extraccion_mineral), nivel=3, espacio_despues=6)

  if remanejo:
    agregar_viñeta(doc, remanejo, nivel=2, espacio_despues=6)

# Renderiza contenido específico dentro del documento Word.
def mlp_render_planta_desaladora(doc, texto_compania, excel_madre=None):
  contenido = [linea.strip() for linea in extraer_planta_desaladora(texto_compania) if linea.strip()]
  if not contenido:
    return
  doc.add_paragraph("") 
  p = doc.add_paragraph("Planta Desaladora:", style="Normal AMSA")
  p.paragraph_format.space_before = Pt(6)
  p.paragraph_format.space_after = Pt(12)

  run = p.runs[0]
  run.bold = True
  run.font.color.rgb = RGBColor(0x00, 0x77, 0x8B)
  run.font.name = "Arial"
  run.font.size = Pt(11)

  i = 0
  while i < len(contenido):
    texto = contenido[i]

    if re.match(r"^\d{1,2}\sde\s\w+\sde\s\d{4}:", texto):
      if i + 1 < len(contenido) and contenido[i + 1].strip().startswith("Restricción:"):
        texto = texto.strip() + " " + contenido[i + 1].strip()
        i += 1

      p = doc.add_paragraph(style="Normal AMSA")
      p.paragraph_format.line_spacing = 1.0
      p.paragraph_format.space_before = Pt(6)
      p.paragraph_format.space_after = Pt(6)
      p.paragraph_format.left_indent = Cm(1.27)
      p.paragraph_format.first_line_indent = Cm(-0.42)

      run_bullet = p.add_run("○  ")
      run_bullet.font.name = "Arial"
      run_bullet.font.size = Pt(11)
      run_bullet.bold = False

      texto = limpiar_texto_global(texto)
      partes = texto.split(":", 1)

      run_bold = p.add_run(partes[0].strip() + ": ")
      run_bold.bold = True
      run_bold.font.name = "Arial"
      run_bold.font.size = Pt(11)

      if len(partes) > 1:
        run_normal = p.add_run(partes[1].strip())
        run_normal.font.name = "Arial"
        run_normal.font.size = Pt(11)

    else:
      p = doc.add_paragraph(style="Normal AMSA")
      p.paragraph_format.line_spacing = 1.0
      p.paragraph_format.space_before = Pt(0)
      p.paragraph_format.space_after = Pt(6)

      run = p.add_run(texto)
      run.font.name = "Arial"
      run.font.size = Pt(11)

    i += 1

# Renderiza contenido específico dentro del documento Word.
def mlp_render_gestion_hidrica(doc, texto_compania, excel_madre):
  contenido = [linea.strip() for linea in extraer_gestión_hídrica(texto_compania) if linea.strip()]
  if not contenido:
    return
  doc.add_paragraph("") 
  p = doc.add_paragraph("Gestión Hídrica:", style="Normal AMSA")
  p.paragraph_format.space_before = Pt(6)
  p.paragraph_format.space_after = Pt(12)

  run = p.runs[0]
  run.bold = True
  run.font.color.rgb = RGBColor(0x00, 0x77, 0x8B)
  run.font.name = "Arial"
  run.font.size = Pt(11)

  if excel_madre:
    img_hidrica_mlp = exportar_imagen_excel(
      excel_madre, "Gestión Hídrica", "A3:W20", "tabla_hidrica_mlp.png"
    )
    agregar_imagen(doc, img_hidrica_mlp, 19, 3.24, "")

    p_img = doc.add_paragraph(style="Normal AMSA")
    p_img.paragraph_format.space_before = Pt(0)
    p_img.paragraph_format.space_after = Pt(12)
  else:
    # MLP no seleccionada: usar imagen cacheada del Word previo si existe
    img_cache = os.path.join(r"C:\Temp", "tabla_hidrica_mlp.png")
    if os.path.exists(img_cache) and os.path.getsize(img_cache) > 0:
      agregar_imagen(doc, img_cache, 19, 3.24, "")
      p_img = doc.add_paragraph(style="Normal AMSA")
      p_img.paragraph_format.space_before = Pt(0)
      p_img.paragraph_format.space_after = Pt(12)

  patron_fechas = re.compile(
    r"\b(?:El día\s)?\d{1,2}\sde\s\w+\sde\s\d{4}(?:\s*-\s*\d{1,2}\sde\s\w+\sde\s\d{4})?\b",
    flags=re.IGNORECASE
  )

  for linea in contenido:
    texto = linea.strip()
    if not texto:
      continue

    p = doc.add_paragraph(style="Viñeta 1")
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(8)

    match_cabecera = re.search(r"^([^:]{2,80}):\s*", texto)
    if match_cabecera:
      cabecera = match_cabecera.group(1).strip() + ": "
      resto = texto[match_cabecera.end():].strip()

      run_c = p.add_run(cabecera)
      run_c.bold = True
      run_c.font.name = "Arial"
      run_c.font.size = Pt(11)

      pos = 0
      for match in patron_fechas.finditer(resto):
        if match.start() > pos:
          run_normal = p.add_run(resto[pos:match.start()])
          run_normal.font.name = "Arial"
          run_normal.font.size = Pt(11)

        run_fecha = p.add_run(match.group(0).lower())
        run_fecha.bold = False 
        run_fecha.font.name = "Arial"
        run_fecha.font.size = Pt(11)

        pos = match.end()

      if pos < len(resto):
        run_final = p.add_run(resto[pos:])
        run_final.font.name = "Arial"
        run_final.font.size = Pt(11)

    else:
      pos = 0
      for match in patron_fechas.finditer(texto):
        if match.start() > pos:
          run_normal = p.add_run(texto[pos:match.start()])
          run_normal.font.name = "Arial"
          run_normal.font.size = Pt(11)

        run_fecha = p.add_run(match.group(0).lower())
        run_fecha.bold = False 
        run_fecha.font.name = "Arial"
        run_fecha.font.size = Pt(11)

        pos = match.end()

      if pos < len(texto):
        run_final = p.add_run(texto[pos:])
        run_final.font.name = "Arial"
        run_final.font.size = Pt(11)

# Procesa la sección o faena indicada usando las reglas correspondientes.
def procesar_mlp(doc, texto_compania, excel_madre):
    agregar_hechos_relevantes(doc, texto_compania, compania="MLP")
    doc.add_page_break()
    agregar_produccion_semana_faena(doc, "MLP", excel_madre)
    doc.add_paragraph()
    agregar_titulo(doc, "Principales Desviaciones", nivel=2)
    validar_acumulados_principales_desviaciones(texto_compania, "MLP", es_seleccionada=excel_madre is not None)
    orden = ORDEN_PRINCIPALES_DESVIACIONES["MLP"]
    for nombre_seccion, orden_subtitulos in orden.items():
        if nombre_seccion == "Mina":
            mlp_render_mina(doc, texto_compania, excel_madre)
        elif nombre_seccion == "Planta Desaladora": 
            mlp_render_planta_desaladora(doc, texto_compania, excel_madre)
        elif nombre_seccion == "Gestión Hídrica":
            mlp_render_gestion_hidrica(doc, texto_compania, excel_madre)
        elif nombre_seccion == "Concentradora":
            doc.add_paragraph("")
            procesar_seccion(doc, texto_compania, "MLP", nombre_seccion, orden_subtitulos, excel_madre)
        else:
            procesar_seccion(doc, texto_compania, "MLP", nombre_seccion, orden_subtitulos, excel_madre)

# Implementa una parte específica de la lógica del informe.
def _procesar_faena_generica(doc, texto_compania, excel_madre, clave):
    agregar_hechos_relevantes(doc, texto_compania, compania=clave)
    doc.add_page_break()
    agregar_produccion_semana_faena(doc, clave, excel_madre)
    doc.add_paragraph("") 
    agregar_titulo(doc, "Principales Desviaciones", nivel=2)
    validar_acumulados_principales_desviaciones(texto_compania, clave, es_seleccionada=excel_madre is not None)
    orden = ORDEN_PRINCIPALES_DESVIACIONES.get(clave, {})
    for nombre_seccion, orden_subtitulos in orden.items():
        procesar_seccion(doc, texto_compania, clave, nombre_seccion, orden_subtitulos, excel_madre)

# Procesa la sección o faena indicada usando las reglas correspondientes.
def procesar_ant(doc, texto_compania, excel_madre):
  agregar_hechos_relevantes(doc, texto_compania, compania="ANT")
  doc.add_page_break()
  agregar_produccion_semana_faena(doc, "ANT", excel_madre)
  doc.add_paragraph("") 
  agregar_titulo(doc, "Principales Desviaciones", nivel=2)

  ant_render_mina(doc, texto_compania, excel_madre)
  procesar_seccion(doc, texto_compania, "ANT", "Planta", [""], excel_madre)

# Procesa la sección o faena indicada usando las reglas correspondientes.
def procesar_cen(doc, texto_compania, excel_madre):
  agregar_hechos_relevantes(doc, texto_compania, compania="CEN")
  doc.add_page_break()
  agregar_produccion_semana_faena(doc, "CEN", excel_madre)
  doc.add_paragraph("") 
  agregar_titulo(doc, "Principales Desviaciones", nivel=2)
  validar_acumulados_principales_desviaciones(texto_compania, "CEN", es_seleccionada=excel_madre is not None)
  procesar_seccion(
    doc,
    texto_compania,
    "CEN",
    "Mina",
    [
      "Movimiento Mina",
      "Movimiento en Rajo Tesoro",
      "Movimiento en Rajo Esperanza",
      "Movimiento en Rajo Óxido Encuentro",
      "Movimiento en Rajo Esperanza Sur:",
      "Movimiento en Rajo Encuentro Sulfuros",
    ],
    excel_madre
  )

  procesar_seccion(doc, texto_compania, "CEN", "Sulfuros", [""], excel_madre)
  cen_render_catodos(doc, texto_compania, excel_madre)

# Renderiza contenido específico dentro del documento Word.
def cmz_render_planta(doc, texto_compania, excel_madre=None):
  contenido = [linea.strip() for linea in extraer_planta(texto_compania) if linea.strip()]
  if not contenido:
    print("[WARNING] CMZ - Planta: no se encontró contenido.")
    errores.append("CMZ - Planta: sección vacía")
    return

  agregar_texto(doc, "Planta:", bold=True, color=(0x00, 0x77, 0x8B))

  orden_titulos = [
    "Mineral Apilado HL",
    "Mineral Beneficiado HL",
    "Ley Apilado HL TCu",
    "Mineral Apilado DL",
    "Mineral Beneficiado DL",
    "Ley Apilado DL TCu",
    "Remanejo Ripios",
    "PLS",
    "Cobre Fino Producido",
  ]

  bloques = {titulo: [] for titulo in orden_titulos}
  acumulados = []
  otros = []

  for texto in contenido:
    t = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", texto).strip()
    t_clave = normalizar_texto_clave(t)
    etiqueta = clasificar_subtitulo_cmz_planta(t)

    if "acumulado al mes" in t_clave or "acumulado al ano" in t_clave:
      acumulados.append(t)
    elif etiqueta:
      bloques[etiqueta].append(normalizar_linea_cmz_planta(t))
    else:
      otros.append(t)

  for titulo in orden_titulos:
    if bloques[titulo]:
      for linea in bloques[titulo]:
        if titulo.lower().startswith("ley ") or titulo.lower().startswith("recuperaci"):
          linea = limpiar_parentesis_ley(linea)
        agregar_viñeta_con_titulo(doc, linea, nivel=1, espacio_despues=6)
    else:
      print(f"[WARNING] CMZ - Planta: no se encontró '{titulo}' en formato esperado.")
      errores.append(f"CMZ - Planta: falta '{titulo}'")

  if acumulados:
    doc.add_paragraph("")
    for linea in acumulados:
      agregar_texto(doc, linea)

  for texto in otros:
    print(f"[WARNING] CMZ - Planta: línea no clasificada -> '{texto}'")
    errores.append(f"CMZ - Planta: línea no clasificada -> '{texto}'")
    agregar_viñeta_plana(doc, texto, nivel=1, espacio_despues=6)

# Renderiza contenido específico dentro del documento Word.
def cmz_render_mina(doc, texto_compania, excel_madre=None):
  contenido = [linea.strip() for linea in extraer_mina(texto_compania) if linea.strip()]
  if not contenido:
    print("[WARNING] CMZ - Mina: no se encontró contenido.")
    errores.append("CMZ - Mina: sección vacía")
    return

  agregar_texto(doc, "Mina:", bold=True, color=(0x00, 0x77, 0x8B))

  movimiento_mina = None
  extraccion = None
  fases = []
  extraccion_mineral = None
  extraccion_lastre = None
  remanejo = None
  otros = []

  for texto in contenido:
    t = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", texto).strip()
    clave = normalizar_texto_clave(t)

    if "movimiento mina" in clave:
      movimiento_mina = normalizar_linea_cmz(t)

    elif "extraccion mineral" in clave:
      extraccion_mineral = normalizar_linea_cmz(t)

    elif "extraccion lastre" in clave or "extraccion esteril" in clave:
      extraccion_lastre = normalizar_linea_cmz(t)

    elif "remanejo" in clave:
      remanejo = normalizar_linea_cmz(t)

    elif re.match(r"^fase\s+\S+", t, flags=re.IGNORECASE):
      fases.append(normalizar_linea_cmz(t))

    elif "extraccion" in clave:
      extraccion = normalizar_linea_cmz(t)

    else:
      otros.append(t)

  if movimiento_mina:
    agregar_viñeta_con_titulo(doc, movimiento_mina, nivel=1, espacio_despues=6)
  else:
    print("[REVISAR] CMZ - Mina: no se encontró 'Movimiento Mina' en formato esperado.")
    errores.append("CMZ - Mina: falta 'Movimiento Mina'")

  if extraccion:
    agregar_viñeta_con_titulo(doc, extraccion, nivel=2, espacio_despues=6)
  else:
    print("[REVISAR] CMZ - Mina: no se encontró 'Extracción' en formato esperado.")
    errores.append("CMZ - Mina: falta 'Extracción'")

  for fase in fases:
    agregar_viñeta_con_titulo(doc, fase, nivel=3, espacio_despues=6)

  if extraccion_mineral:
    agregar_viñeta_con_titulo(doc, extraccion_mineral, nivel=2, espacio_despues=6)
  else:
    print("[REVISAR] CMZ - Mina: no se encontró 'Extracción Mineral' en formato esperado.")
    errores.append("CMZ - Mina: falta 'Extracción Mineral'")

  if extraccion_lastre:
    agregar_viñeta_con_titulo(doc, extraccion_lastre, nivel=2, espacio_despues=6)
  else:
    print("[REVISAR] CMZ - Mina: no se encontró 'Extracción Lastre' en formato esperado.")
    errores.append("CMZ - Mina: falta 'Extracción Lastre'")

  if remanejo:
    agregar_viñeta_con_titulo(doc, remanejo, nivel=2, espacio_despues=6)
  else:
    print("[REVISAR] CMZ - Mina: no se encontró 'Remanejo' en formato esperado.")
    errores.append("CMZ - Mina: falta 'Remanejo'")

  for texto in otros:
    print(f"[REVISAR] CMZ - Mina: línea no clasificada -> '{texto}'")
    errores.append(f"CMZ - Mina: línea no clasificada -> '{texto}'")
    agregar_viñeta_plana(doc, texto, nivel=2, espacio_despues=6)

# Procesa la sección o faena indicada usando las reglas correspondientes.
def procesar_cmz(doc, texto_compania, excel_madre):
  agregar_hechos_relevantes(doc, texto_compania, compania="CMZ")
  doc.add_page_break()
  agregar_produccion_semana_faena(doc, "CMZ", excel_madre)
  doc.add_paragraph("")
  agregar_titulo(doc, "Principales Desviaciones", nivel=2)
  validar_acumulados_principales_desviaciones(texto_compania, "CMZ", es_seleccionada=excel_madre is not None)

  cmz_render_mina(doc, texto_compania, excel_madre)
  cmz_render_planta(doc, texto_compania, excel_madre)

# Renderiza contenido específico dentro del documento Word.
def fcab_render_medio_ambiente(doc, lineas):
  patron_fecha = re.compile(r"^\d{1,2}\s+de\s+\w+\s+de\s+\d{4}", re.IGNORECASE)

  for linea in lineas:
    texto = linea.strip()
    if not texto:
      continue

    texto_limpio = re.sub(r"^[•○o·\-\s]+", "", texto).strip()

    if patron_fecha.match(texto_limpio):
      p = doc.add_paragraph(style="Normal AMSA")
      p.paragraph_format.line_spacing = 1.0
      p.paragraph_format.space_before = Pt(0)
      p.paragraph_format.space_after = Pt(6)
      p.paragraph_format.left_indent = Cm(3.0)
      p.paragraph_format.first_line_indent = Cm(-0.4)

      run_bullet = p.add_run("o  ")
      run_bullet.font.name = "Arial"
      run_bullet.font.size = Pt(11)
      run_bullet.bold = False

      match = patron_fecha.match(texto_limpio)
      fecha = match.group(0)
      resto = texto_limpio[len(fecha):].lstrip(": ").strip()

      run_fecha = p.add_run(fecha + ": ")
      run_fecha.bold = True
      run_fecha.font.name = "Arial"
      run_fecha.font.size = Pt(11)

      run_resto = p.add_run(resto)
      run_resto.bold = False
      run_resto.font.name = "Arial"
      run_resto.font.size = Pt(11)

    else:
      p = doc.add_paragraph(style="Normal AMSA")
      p.paragraph_format.line_spacing = 1.0
      p.paragraph_format.space_before = Pt(0)
      p.paragraph_format.space_after = Pt(6)
      p.paragraph_format.left_indent = Cm(1.9)
      p.paragraph_format.first_line_indent = Cm(-0.4)

      run_bullet = p.add_run("o  ")
      run_bullet.font.name = "Arial"
      run_bullet.font.size = Pt(11)
      run_bullet.bold = False

      run_texto = p.add_run(texto_limpio)
      run_texto.font.name = "Arial"
      run_texto.font.size = Pt(11)
      run_texto.bold = False

# Renderiza contenido específico dentro del documento Word.
def fcab_render_tren(doc, texto_compania, excel_madre=None):
  contenido = [linea.strip() for linea in extraer_tren(texto_compania) if linea.strip()]
  if not contenido:
    print("[WARNING] FCAB - Tren: no se encontró contenido.")
    errores.append("FCAB - Tren: sección vacía")
    return

  doc.add_paragraph("")

  p = doc.add_paragraph()
  p.paragraph_format.line_spacing = 1.0
  p.paragraph_format.space_before = Pt(0)
  p.paragraph_format.space_after = Pt(6)
  p.paragraph_format.left_indent = Cm(0)
  p.paragraph_format.first_line_indent = Cm(0)

  run = p.add_run("Tren:")
  run.bold = True
  run.font.name = "Arial"
  run.font.size = Pt(11)
  run.font.color.rgb = RGBColor(0x00, 0x77, 0x8B)

  primera_linea_general = None
  transporte_total = None
  hijos_total = []
  otros_bloques = []
  bloque_actual = None

  for texto in contenido:
    t = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", texto).strip()
    clave = normalizar_texto_clave(t)

    if primera_linea_general is None and (
      "transporte total del grupo" in clave
      or "el transporte total del grupo" in clave
    ):
      primera_linea_general = t

    elif "transporte total de tren" in clave:
      transporte_total = t
      bloque_actual = "total_tren"

    elif (
      "transporte de acido" in clave
      or "transporte de cobre" in clave
      or "transporte de concentrados" in clave
    ):
      otros_bloques.append(("titulo", t))
      bloque_actual = "subbloque"

    else:
      if bloque_actual == "total_tren":
        hijos_total.append(t)
      else:
        otros_bloques.append(("detalle", t))

  if primera_linea_general:
    agregar_parrafo_fcab_alineado(doc, primera_linea_general, bold=False, espacio_antes=False)

  if transporte_total:
    agregar_parrafo_fcab_alineado(doc, transporte_total, bold=True, espacio_antes=False)
  else:
    print("[WARNING] FCAB - Tren: no se encontró 'Transporte Total de Tren'.")
    errores.append("FCAB - Tren: falta 'Transporte Total de Tren'")

  acumulados = []
  detalles_normales = []

  for linea in hijos_total:
    clave = normalizar_texto_clave(linea)
    if "acumulado al mes" in clave or "acumulado al ano" in clave:
      acumulados.append(linea)
    else:
      detalles_normales.append(linea)

  for tipo, linea in otros_bloques:
    if tipo == "titulo":
      agregar_circulo_blanco_manual(doc, linea, left_indent_cm=1.27, bullet_indent_cm=0.85, bold=True, espacio_despues=6)
    else:
      agregar_circulo_blanco_manual(doc, linea, left_indent_cm=1.9, bullet_indent_cm=1.45, espacio_despues=6)

  for linea in detalles_normales:
    agregar_circulo_blanco_manual(doc, linea, espacio_despues=6)

  if acumulados:
    doc.add_paragraph("")
    for linea in acumulados:
      agregar_parrafo_fcab_alineado(doc, linea, bold=False, espacio_antes=False)

# Renderiza contenido específico dentro del documento Word.
def fcab_render_camion(doc, texto_compania, excel_madre=None):
  contenido = [linea.strip() for linea in extraer_camión(texto_compania) if linea.strip()]
  if not contenido:
    print("[WARNING] FCAB - Camión: no se encontró contenido.")
    errores.append("FCAB - Camión: sección vacía")
    return

  doc.add_paragraph("")
  agregar_texto(doc, "Camión:", bold=True, color=(0x00, 0x77, 0x8B))

  transporte_total = None
  detalles = []
  acumulados = []
  otros = []

  for texto in contenido:
    t = re.sub(r"^[•○o·\-\s\u200b\ufeff]+", "", texto).strip()
    clave = normalizar_texto_clave(t)

    if "transporte total de camion" in clave:
      transporte_total = t
    elif "acumulado al mes" in clave or "acumulado al ano" in clave:
      acumulados.append(t)
    elif transporte_total:
      detalles.append(t)
    else:
      otros.append(t)

  if transporte_total:
    agregar_parrafo_fcab_alineado(doc, transporte_total, bold=True, espacio_antes=False)
  else:
    print("[WARNING] FCAB - Camión: no se encontró 'Transporte Total de Camión'.")
    errores.append("FCAB - Camión: falta 'Transporte Total de Camión'")

  detalles_reales = []
  for linea in detalles:
    clave = normalizar_texto_clave(linea)
    if "acumulado al mes" in clave or "acumulado al ano" in clave:
      acumulados.append(linea)
    else:
      detalles_reales.append(linea)

  for linea in detalles_reales:
    agregar_circulo_blanco_manual(doc, linea, espacio_despues=6)

  for linea in acumulados:
    agregar_parrafo_fcab_alineado(doc, linea, bold=False, espacio_antes=True)

  for linea in otros:
    print(f"[WARNING] FCAB - Camión: línea no clasificada -> '{linea}'")
    errores.append(f"FCAB - Camión: línea no clasificada -> '{linea}'")
    agregar_circulo_blanco_manual(doc, linea, espacio_despues=6)

# Procesa la sección o faena indicada usando las reglas correspondientes.
def procesar_fcab(doc, texto_compania, excel_madre):
  agregar_hechos_relevantes(doc, texto_compania, compania="FCAB")

  doc.add_page_break()
  agregar_produccion_semana_faena(doc, "FCAB", excel_madre)
  doc.add_paragraph("")

  agregar_titulo(doc, "Principales Desviaciones", nivel=2)
  validar_acumulados_principales_desviaciones(texto_compania, "FCAB", es_seleccionada=excel_madre is not None)

  fcab_render_tren(doc, texto_compania, excel_madre)
  fcab_render_camion(doc, texto_compania, excel_madre)

# Relaciona cada faena con su función procesadora principal.
PROCESADORES_FAENA = {
    "MLP": procesar_mlp,
    "ANT": procesar_ant,
    "CEN": procesar_cen,
    "CMZ": procesar_cmz,
    "FCAB": procesar_fcab,
}

SECCIONES_HECHOS = [
    {
        "titulo": "Accidentabilidad",
        "extractor": extraer_accidentabilidad,
        "regla_nivel": lambda linea: 3 if re.match(r"^\d{1,2}\s+de\s+\w+(\s+de)?\s+\d{4}", linea.strip()) else 2,
    },
    {
        "titulo": "Reportabilidad",
        "extractor": extraer_reportabilidad,
        "regla_nivel": lambda linea: 2
        if any(
            x in linea
            for x in [
                "cuasi accidente",
                "Cuasi accidente",
                "Cuasi accidentes",
                "cuasi accidentes",
                "hallazgo",
                "Hallazgo",
                "hallazgos",
                "Hallazgos",
                "YDN",
            ]
        )
        else 3,
    },
    {
        "titulo": "Gestión SSO",
        "extractor": extraer_gestion_sso,
        "regla_nivel": lambda linea: 2,
        "plain": True,
    },
    {
        "titulo": "Salud Ocupacional y Gestión Vial",
        "extractor": extraer_salud_ocupacional,
        "regla_nivel": lambda linea: 2,
    },
    {
        "titulo": "Medio Ambiente",
        "extractor": extraer_medio_ambiente,
        "regla_nivel": lambda linea: 2,
    },
    {
        "titulo": "Asuntos Públicos",
        "extractor": extraer_asuntos_publicos,
        "regla_nivel": lambda linea: 2,
    },
]

# Agrega al documento el elemento indicado por su nombre.
def agregar_hechos_relevantes(doc, texto_compania, compania=None):
  agregar_titulo(doc, "Hechos Relevantes", nivel=2)

  for idx, seccion in enumerate(SECCIONES_HECHOS):
    lineas = seccion["extractor"](texto_compania)
    if not lineas:
      continue

    p_titulo = doc.add_paragraph(style="Viñeta 1")
    p_titulo.paragraph_format.line_spacing = 1.0
    p_titulo.paragraph_format.space_before = Pt(0)
    p_titulo.paragraph_format.space_after = Pt(6)

    run_titulo = p_titulo.add_run(seccion["titulo"])
    run_titulo.bold = True
    run_titulo.underline = True
    run_titulo.font.name = "Arial"
    run_titulo.font.size = Pt(11)

    if seccion["titulo"] == "Accidentabilidad" and compania == "MLP":
      mlp_render_accidentabilidad(doc, lineas)

    elif seccion["titulo"] == "Accidentabilidad" and compania in ("FCAB", "CMZ"):
      render_accidentabilidad_generica(doc, lineas)

    elif seccion["titulo"] == "Asuntos Públicos" and compania == "MLP":
      mlp_render_asuntos_publicos(doc, lineas)

    elif seccion["titulo"] == "Medio Ambiente" and compania == "ANT":
      ant_render_medio_ambiente(doc, lineas)

    elif seccion["titulo"] == "Medio Ambiente" and compania == "MLP":
      mlp_render_medio_ambiente(doc, lineas)

    elif seccion["titulo"] == "Medio Ambiente" and compania == "CEN":
      cen_render_medio_ambiente(doc, lineas)

    elif seccion["titulo"] == "Medio Ambiente" and compania == "FCAB":
      fcab_render_medio_ambiente(doc, lineas)

    else:
      for linea in lineas:
        if "Medición calidad de aire" in linea:
          p = doc.add_paragraph(style="Normal AMSA")
          p.paragraph_format.line_spacing = 1.0
          p.paragraph_format.space_before = Pt(0)
          p.paragraph_format.space_after = Pt(6)

          run = p.add_run(linea)
          run.font.name = "Arial"
          run.font.size = Pt(11)

        else:
          nivel = seccion["regla_nivel"](linea)

          if compania == "CEN" and nivel >= 3:
            agregar_viñeta_fecha_inicial(doc, linea, nivel=nivel, espacio_despues=6)
          elif seccion.get("plain"):
            agregar_viñeta_plana(doc, linea, nivel=nivel, espacio_despues=6)
          elif compania == "FCAB" and seccion["titulo"] == "Salud Ocupacional y Gestión Vial":
            agregar_viñeta_plana(doc, linea, nivel=nivel, espacio_despues=6)
          else:
            agregar_viñeta(doc, linea, nivel=nivel, espacio_despues=6)

    if idx < len(SECCIONES_HECHOS) - 1:
      p_sep = doc.add_paragraph(style="Normal AMSA")
      p_sep.paragraph_format.space_before = Pt(0)
      p_sep.paragraph_format.space_after = Pt(6)