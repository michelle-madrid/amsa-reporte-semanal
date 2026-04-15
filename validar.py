"""Validador standalone: compara KPIs del Word final contra el Excel madre."""

import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from docx import Document

from config import CONFIG_COMPANIAS
from core.validador import validar_kpis_vs_excel
from utils.excel_utils import _obtener_excel_app
from state import _workbooks_abiertos


# ── Selector de archivos ───────────────────────────────────────────────────────

def _pedir_archivo(titulo, tipos):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    ruta = filedialog.askopenfilename(title=titulo, filetypes=tipos)
    root.destroy()
    return ruta or None


# ── Separar el Word final por compañía ────────────────────────────────────────

_NOMBRE_A_CLAVE = {cfg["nombre"]: clave for clave, cfg in CONFIG_COMPANIAS.items()}


def _extraer_informes_desde_word_final(ruta_word):
    """
    Lee el Word final y devuelve un dict {clave: texto} separando por sección
    de compañía. Detecta los títulos por el texto exacto del nombre de la compañía.
    """
    doc = Document(ruta_word)
    informes = {}
    clave_actual = None
    buffer = []

    def _guardar():
        if clave_actual and buffer:
            informes[clave_actual] = "\n".join(buffer)

    for p in doc.paragraphs:
        texto = p.text.strip()
        if not texto:
            continue
        if texto in _NOMBRE_A_CLAVE:
            _guardar()
            clave_actual = _NOMBRE_A_CLAVE[texto]
            buffer = []
        elif clave_actual:
            buffer.append(texto)

    _guardar()

    if not informes:
        print("  ! No se encontraron secciones de compañía en el Word.")
        print(f"    Nombres buscados: {list(_NOMBRE_A_CLAVE.keys())}")

    return informes


# ── Abrir Excel vía COM ────────────────────────────────────────────────────────

def _abrir_excel(ruta_excel):
    ruta = str(Path(ruta_excel).resolve())
    wb = _workbooks_abiertos.get(ruta)
    if wb is None:
        print(f"  Abriendo Excel: {Path(ruta).name} ...")
        wb = _obtener_excel_app().Workbooks.Open(ruta, UpdateLinks=0)
        _workbooks_abiertos[ruta] = wb
    return wb


# ── Punto de entrada ───────────────────────────────────────────────────────────

def main():
    print("\n── Validador KPIs Word vs Excel ────────────────────────────────")

    ruta_word = _pedir_archivo(
        "Selecciona el Word final del informe",
        [("Word", "*.docx"), ("Todos", "*.*")],
    )
    if not ruta_word:
        print("  ! No se seleccionó el Word. Abortando.")
        return

    ruta_excel = _pedir_archivo(
        "Selecciona el Excel madre (actualizado)",
        [("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")],
    )
    if not ruta_excel:
        print("  ! No se seleccionó el Excel. Abortando.")
        return

    print(f"\n  Word : {Path(ruta_word).name}")
    print(f"  Excel: {Path(ruta_excel).name}")

    print("\n  Leyendo Word final...")
    informes = _extraer_informes_desde_word_final(ruta_word)
    if not informes:
        print("  ! No se pudo extraer texto. Verifica que el Word tenga el formato esperado.")
        return

    print(f"  Compañías detectadas: {', '.join(informes.keys())}")

    wb = _abrir_excel(ruta_excel)
    validar_kpis_vs_excel(informes, wb)


if __name__ == "__main__":
    main()
