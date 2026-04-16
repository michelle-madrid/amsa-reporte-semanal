"""Punto de entrada para generar el informe semanal completo."""

import os

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

import state
from config import *
from utils.word_utils import *
from utils.excel_utils import *
from core.extractores import *
from core.renderers import *
from utils.excel_utils import _obtener_excel_app, _rangos_tablas_sso_backup_dinamico
from state import _workbooks_abiertos
from core.validador import validar_kpis_vs_excel


# ─────────────────────────────────────────────────────────────────────────────
# Helper compartido: construye el Word desde un dict de informes ya listo.
# ─────────────────────────────────────────────────────────────────────────────
_MSG_PENDIENTE = "[No solicitado. Presumiblemente en espera de envío información]"

def _construir_doc(
    informes,            # dict clave→texto_compania (para TODAS las faenas)
    excel_madre,
    excel_indicadores,
    dia_inicio, mes_inicio, dia_fin, mes_fin, year, num_semana,
    carpeta_destino, nombre_final,
    incluir_sso=True,
    incluir_gh=True,
    faenas_con_excel=None,  # None = todas; set/list = solo esas usan Excel para su imagen
):
    """Construye y guarda el documento Word.  No hace preguntas ni resuelve rutas."""
    escribir_fechas_excel(excel_madre, dia_inicio, mes_inicio, dia_fin, mes_fin)

    doc = Document(RUTA_PLANTILLA)
    section = doc.sections[0]
    section.top_margin    = Cm(1.27)
    section.bottom_margin = Cm(1.27)
    section.left_margin   = Cm(1.27)
    section.right_margin  = Cm(1.27)

    p = doc.add_paragraph("Informe Semanal de Operación - Antofagasta PLC", style="Título 1 AMSA")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(24)
    for run in p.runs:
        run.font.color.rgb = RGBColor(0x12, 0x6F, 0x7A)

    # ── Aviso de información incompleta ───────────────────────────────
    faenas_sin_texto = [c for c in ORDEN_OFICIAL if not informes.get(c)]
    info_incompleta  = bool(faenas_sin_texto) or not incluir_sso or not incluir_gh
    if info_incompleta:
        partes = []
        if faenas_sin_texto:
            partes.append(f"Faenas: {', '.join(faenas_sin_texto)}")
        if not incluir_sso:
            partes.append("SSO")
        if not incluir_gh:
            partes.append("Gestión Hídrica")
        aviso = f"[Información incompleta — pendiente: {'; '.join(partes)}]"
        p_aviso = doc.add_paragraph(style="Normal AMSA")
        p_aviso.paragraph_format.space_before = Pt(0)
        p_aviso.paragraph_format.space_after  = Pt(10)
        p_aviso.paragraph_format.line_spacing = 1.0
        run_aviso = p_aviso.add_run(aviso)
        run_aviso.bold = True
        run_aviso.font.name = "Arial"
        run_aviso.font.size = Pt(11)
        run_aviso.font.color.rgb = RGBColor(0xC0, 0x50, 0x00)   # naranja oscuro
        print(f"  ⚠ {aviso}")

    resumen_texto = extraer_resumen_excel(excel_madre)
    for linea in resumen_texto.split("\n"):
        linea_limpia = linea.strip()
        if linea_limpia:
            agregar_texto(doc, linea_limpia)
            if linea_limpia.endswith("."):
                doc.add_paragraph()

    for _ in range(4):
        doc.add_paragraph()

    img_resumen = exportar_imagen_excel(excel_madre, "Grupo Minero FCAB PLAN", "A3:X34", "tabla_principal.png")
    agregar_imagen(doc, img_resumen, 19, 8.8, "")

    if INCLUIR_ESTADO_FASES_DESARROLLO:
        agregar_estado_fases_desarrollo(doc, excel_madre)

    doc.add_page_break()
    agregar_titulo(doc, "Gestión Hídrica", nivel=2)
    if incluir_gh:
        img_hidrica = exportar_imagen_excel(excel_madre, "Gestión Hídrica", "A3:W20", "gestion_hidrica.png")
        agregar_imagen(doc, img_hidrica, 19, 3.24, "")
    else:
        agregar_texto(doc, _MSG_PENDIENTE, color=(128, 128, 128))
        print("  → Gestión Hídrica: En espera de envío información")

    doc.add_page_break()
    agregar_titulo(doc, "Accidentabilidad", nivel=2)
    if incluir_sso:
        img_semanal = exportar_imagen_excel(excel_indicadores, "Informe Viernes", "A29:M41", "valor_semanal.png")
        img_mensual = exportar_imagen_excel(excel_indicadores, "Informe Viernes", "A15:M27", "valor_mensual.png")
        img_anual   = exportar_imagen_excel(excel_indicadores, "Informe Viernes", "A1:M13",  "valor_anual.png")
        for img_path, texto_titulo in [
            (img_semanal, "Indicadores Valor Semanal"),
            (img_mensual, "Indicadores Valor Mensual"),
            (img_anual,   "Indicadores Valor Anual"),
        ]:
            p_titulo = doc.add_paragraph()
            p_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            run = p_titulo.add_run(texto_titulo)
            run.bold = False
            run.font.name = "Arial"
            run.font.size = Pt(11)
            doc.add_paragraph()
            agregar_imagen(doc, img_path, 19, 4.3)
            doc.add_paragraph()
    else:
        agregar_texto(doc, _MSG_PENDIENTE, color=(128, 128, 128))
        print("  → Accidentabilidad (SSO): En espera de envío información")

    for clave in ORDEN_OFICIAL:
        cfg = CONFIG_COMPANIAS[clave]
        texto_compania = informes.get(clave, "")
        doc.add_page_break()
        agregar_titulo(doc, cfg["nombre"], nivel=1)
        if not texto_compania:
            agregar_texto(doc, _MSG_PENDIENTE, color=(128, 128, 128))
            continue
        procesador = PROCESADORES_FAENA.get(clave)
        if procesador:
            # Solo regenerar imagen Excel para las faenas seleccionadas
            excel_para_faena = excel_madre if (faenas_con_excel is None or clave in faenas_con_excel) else None
            procesador(doc, texto_compania, excel_para_faena)

    # Accidentabilidad Back-up — siempre al final, después de todas las faenas
    doc.add_page_break()
    agregar_titulo(doc, "Accidentabilidad Back-up", nivel=1)
    if incluir_sso:
        def _obtener_wb_madre():
            wb = _workbooks_abiertos.get(excel_madre)
            if wb is None:
                wb = _obtener_excel_app().Workbooks.Open(excel_madre, UpdateLinks=0)
                _workbooks_abiertos[excel_madre] = wb
            else:
                # Verificar que el proxy COM sigue activo; si no, re-abrir
                try:
                    _ = wb.Name
                except Exception:
                    del _workbooks_abiertos[excel_madre]
                    wb = _obtener_excel_app().Workbooks.Open(excel_madre, UpdateLinks=0)
                    _workbooks_abiertos[excel_madre] = wb
            return wb

        wb_madre = _obtener_wb_madre()
        rangos_tablas_sso = _rangos_tablas_sso_backup_dinamico(wb_madre.Worksheets("SSO"))
        if not rangos_tablas_sso:
            print("  ! Accidentabilidad Back-up: no se encontraron tablas con datos. Verifica hoja SSO.")
        for i, rango_tabla in enumerate(rangos_tablas_sso):
            nombre_img = f"accidentabilidad_{i + 1}.png"
            wb_madre = _obtener_wb_madre()
            img_backup = exportar_imagen_sso_filtrada(excel_madre, wb_madre.Worksheets("SSO"), rango_tabla, nombre_img)
            if i > 0:
                doc.add_page_break()
            agregar_imagen(doc, img_backup, 19, None, "")
        if rangos_tablas_sso:
            print(f"  ✓ Accidentabilidad Back-up: {len(rangos_tablas_sso)} tabla(s) exportada(s)")
    else:
        agregar_texto(doc, _MSG_PENDIENTE, color=(128, 128, 128))
        print("  → Accidentabilidad Back-up (SSO): En espera de envío información")

    texto_pie = construir_texto_semana(dia_inicio, mes_inicio, dia_fin, mes_fin, year)
    agregar_pie_de_pagina(doc, texto_pie)

    ruta_guardado = os.path.join(carpeta_destino, f"{nombre_final}.docx")
    doc.save(ruta_guardado)
    cerrar_excels()
    print(f"Informe generado en: {ruta_guardado}")


# ─────────────────────────────────────────────────────────────────────────────
# Actualizar secciones de un Word existente (modo parcial)
# ─────────────────────────────────────────────────────────────────────────────
def actualizar_secciones_word(
    ruta_existente,
    faenas_actualizar,
    dia_inicio, mes_inicio, dia_fin, mes_fin, year, num_semana,
    excel_madre, excel_indicadores, carpeta_destino,
    nombre_override=None, actualizar_vinculos=False,
    informes_paths=None,   # dict clave → ruta del Word fuente de cada faena a actualizar
    incluir_sso=True,
    incluir_gh=True,
    disco=None,
):
    """
    Carga un Word existente y regenera solo las secciones de las faenas indicadas.
    Las secciones no seleccionadas se extraen del Word existente y se re-procesan.
    """
    if not Path(ruta_existente).is_file():
        print(f"[ERROR] Word existente no encontrado: {ruta_existente}")
        return

    print(f"\n── Modo Word existente: {Path(ruta_existente).name}")
    print(f"  Secciones a actualizar: {', '.join(faenas_actualizar)}")

    # ── Extraer textos del Word existente ─────────────────────────────
    _N2K = {cfg["nombre"]: k for k, cfg in CONFIG_COMPANIAS.items()}
    doc_prev = Document(ruta_existente)
    informes_previos, clave_actual, buf = {}, None, []

    def _save_buf():
        if clave_actual and buf:
            informes_previos[clave_actual] = "\n".join(buf)

    for p in doc_prev.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if t in _N2K:
            _save_buf()
            clave_actual = _N2K[t]
            buf = []
        elif clave_actual:
            buf.append(t)
    _save_buf()

    if informes_previos:
        print(f"  Secciones detectadas en Word previo: {', '.join(informes_previos)}")
    else:
        print(f"  ⚠ No se detectaron secciones de compañía en el Word existente")

    # ── Construir dict de informes: fresh para las elegidas, previo para el resto ─
    informes = {}
    for clave in ORDEN_OFICIAL:
        if clave in faenas_actualizar:
            ruta_src = (informes_paths or {}).get(clave, "")
            if ruta_src and Path(ruta_src).is_file():
                informes[clave] = extraer_texto_word(ruta_src)
                print(f"  ✓ {clave}: texto extraído de Word fuente")
            else:
                texto_prev = informes_previos.get(clave, "")
                informes[clave] = texto_prev
                if texto_prev:
                    print(f"  ⚠ {clave}: sin Word fuente → reutilizando texto del Word existente")
                else:
                    print(f"  ⚠ {clave}: sin Word fuente y sin sección previa → quedará vacío")
        else:
            informes[clave] = informes_previos.get(clave, "")
            if informes[clave]:
                print(f"  → {clave}: mantenido del Word existente")

    # ── Actualizar vínculos Excel solo para las faenas seleccionadas ──
    if actualizar_vinculos:
        rutas = construir_rutas_semana(num_semana, dia_inicio, mes_inicio, dia_fin, mes_fin, year, disco=disco)
        dirs_parcial = {k: v for k, v in rutas["informes_dirs"].items()
                        if k in faenas_actualizar}
        if dirs_parcial:
            abrir_excel_y_actualizar_vinculos(
                excel_madre, dirs_parcial, carpeta_destino=carpeta_destino,
                ordenar_sso="SSO" in faenas_actualizar,
            )
        nombre_base = os.path.splitext(os.path.basename(excel_madre))[0]
        excel_madre_act = os.path.join(carpeta_destino, f"{nombre_base}_act.xlsx")
        if os.path.exists(excel_madre_act):
            print(f"  Usando Excel actualizado: {os.path.basename(excel_madre_act)}")
            excel_madre = excel_madre_act

    nombre_final = nombre_override or os.path.splitext(os.path.basename(ruta_existente))[0]
    _construir_doc(
        informes, excel_madre, excel_indicadores,
        dia_inicio, mes_inicio, dia_fin, mes_fin, year, num_semana,
        carpeta_destino, nombre_final,
        incluir_sso=incluir_sso,
        incluir_gh=incluir_gh,
        faenas_con_excel=set(faenas_actualizar),
    )


# ─────────────────────────────────────────────────────────────────────────────
# Generar informe semanal completo a partir de los archivos de entrada.
# ─────────────────────────────────────────────────────────────────────────────
def generar_informe(nombre_override=None, incluir_sso=True, incluir_gh=True, disco=None):
    def pedir_entero(mensaje, minimo, maximo):
        while True:
            valor = input(mensaje).strip()
            if valor.isdigit() and minimo <= int(valor) <= maximo:
                return valor
            print(f"  Valor inválido. Ingrese un número entre {minimo} y {maximo}.")

    dia_inicio = pedir_entero("Ingrese el día de inicio: ", 1, 31)
    mes_inicio = pedir_entero("Ingrese el mes de inicio: ", 1, 12)
    dia_fin    = pedir_entero("Ingrese el día de término: ", 1, 31)
    mes_fin    = pedir_entero("Ingrese el mes de término: ", 1, 12)
    year       = pedir_entero("Ingrese el año: ", 2000, 2100)
    num_semana = pedir_entero("Ingrese el número de semana: ", 1, 53)
    texto_pie = f"Semana del {dia_inicio} de {mes_inicio} al {dia_fin} de {mes_fin} {year}"

    orden_oficial = list(ORDEN_OFICIAL)
    seleccion = input(
        "Indica las faenas a procesar separadas por coma (ej: MLP, ANT) o presiona ENTER para todas: "
    ).upper().replace(" ", "")

    if seleccion == "__NINGUNA__":
        faenas_activas = []
    elif seleccion:
        faenas_activas = seleccion.split(",")
    else:
        faenas_activas = orden_oficial

    def _resolver_archivo(ruta_esperada, mensaje):
        """Usa la ruta construida si existe; si no, abre selector."""
        if ruta_esperada and Path(ruta_esperada).is_file():
            print(f"  ✓ {mensaje}: {ruta_esperada.name}")
            return str(ruta_esperada)
        print(f"  ! {mensaje}: no encontrado en ruta esperada → abriendo selector")
        return seleccionar_archivo(mensaje)

    def _resolver_unico_docx(carpeta, mensaje):
        """Busca un único .docx en la carpeta; si hay exactamente uno lo usa, si no abre selector."""
        carpeta = Path(carpeta)
        if carpeta.is_dir():
            docxs = [f for f in carpeta.glob("*.docx") if not f.name.startswith("~$")]
            if len(docxs) == 1:
                print(f"  ✓ {mensaje}: {docxs[0].name}")
                return str(docxs[0])
            if len(docxs) > 1:
                print(f"  ! {mensaje}: {len(docxs)} archivos .docx encontrados → abriendo selector")
        else:
            print(f"  ! {mensaje}: carpeta no encontrada → abriendo selector")
        return seleccionar_archivo(mensaje)

    def _resolver_unico_xlsx(carpeta, mensaje):
        """Busca un .xlsx que empiece con 'BDatos' en la carpeta; si no abre selector."""
        carpeta = Path(carpeta)
        if carpeta.is_dir():
            candidatos = [f for f in carpeta.glob("BDatos*.xlsx") if not f.name.startswith("~$")]
            if len(candidatos) == 1:
                print(f"  ✓ {mensaje}: {candidatos[0].name}")
                return str(candidatos[0])
            if len(candidatos) > 1:
                print(f"  ! {mensaje}: {len(candidatos)} archivos BDatos*.xlsx encontrados → abriendo selector")
            else:
                print(f"  ! {mensaje}: no se encontró BDatos*.xlsx en {carpeta} → abriendo selector")
        else:
            print(f"  ! {mensaje}: carpeta no encontrada → abriendo selector")
        return seleccionar_archivo(mensaje)

    if MODO_DEBUG:
        rutas = construir_rutas_semana(num_semana, dia_inicio, mes_inicio, dia_fin, mes_fin, year, disco=disco)
        excel_madre       = _resolver_archivo(rutas["excel_madre"], "Excel Base")
        excel_indicadores = _resolver_unico_xlsx(rutas["excel_indicadores_dir"], "Excel de indicadores")
        carpeta_destino   = rutas["carpeta_destino"] if Path(rutas["carpeta_destino"]).is_dir() else seleccionar_carpeta()
        nombre_final      = nombre_override or rutas["nombre_archivo"]
    else:
        excel_madre = seleccionar_archivo("Excel Base")
        excel_indicadores = seleccionar_archivo("Excel de indicadores")
        carpeta_destino = seleccionar_carpeta()
        nombre_final = nombre_override or input("\nEscribe el nombre del informe final: ")

    if MODO_DEBUG:
        informes_dirs = dict(rutas["informes_dirs"])
        informes_dirs["SSO"] = rutas["excel_indicadores_dir"]
        carpeta_raiz = Path(rutas["raiz"])
        carpetas_07 = [f for f in carpeta_raiz.iterdir() if f.is_dir() and f.name.startswith("07")]
        if carpetas_07:
            informes_dirs["Gestión Hídrica"] = carpetas_07[0]
        else:
            print("  ! No se encontró subcarpeta '07...' para Gestión Hídrica")

        nombre_base = os.path.splitext(os.path.basename(excel_madre))[0]
        excel_madre_act = os.path.join(carpeta_destino, f"{nombre_base}_act.xlsx")

        actualizar = input("\n¿Deseas actualizar los vínculos del Excel? (s/n): ").strip().lower()
        if actualizar == "s":
            abrir_excel_y_actualizar_vinculos(
                excel_madre, informes_dirs, carpeta_destino=carpeta_destino,
                ordenar_sso=incluir_sso,
            )

        if os.path.exists(excel_madre_act):
            print(f"  Usando Excel actualizado: {os.path.basename(excel_madre_act)}")
            excel_madre = excel_madre_act
        else:
            print("  ! No se encontró archivo _act, usando Excel original.")
    else:
        informes_dirs = None

    informes = {}
    for clave in orden_oficial:
        if clave in faenas_activas:
            if MODO_DEBUG:
                ruta = _resolver_unico_docx(rutas["informes_dirs"][clave], f"Informe {clave}")
            else:
                ruta = seleccionar_archivo(f"Informe {clave}")
            informes[clave] = extraer_texto_word(ruta) if ruta else ""

    validar = input("\n¿Deseas validar KPIs Word vs Excel? (s/n): ").strip().lower()

    _construir_doc(
        informes, excel_madre, excel_indicadores,
        dia_inicio, mes_inicio, dia_fin, mes_fin, year, num_semana,
        carpeta_destino, nombre_final,
        incluir_sso=incluir_sso,
        incluir_gh=incluir_gh,
    )

    if validar == "s":
        wb_madre = _workbooks_abiertos.get(excel_madre)
        if wb_madre is None:
            wb_madre = _obtener_excel_app().Workbooks.Open(excel_madre, UpdateLinks=0)
            _workbooks_abiertos[excel_madre] = wb_madre
        validar_kpis_vs_excel(informes, wb_madre)

if __name__ == "__main__":
    generar_informe()
