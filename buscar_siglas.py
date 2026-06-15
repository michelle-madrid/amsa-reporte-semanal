"""Búsqueda rápida de siglas sin contexto en un informe Word.

Corre SOLO la detección de siglas (sin abrir Excel ni validar KPIs), por si la
validación completa se demora demasiado.

Uso:
    python buscar_siglas.py [ruta_al_docx]
    Sin argumento → abre un selector de archivo.

Salida: por cada compañía, lista las siglas que aparecen sin estar definidas
en el documento ni en el glosario conocido, con la línea donde se usan.
"""

import sys
from pathlib import Path

from docx import Document

from config import CONFIG_COMPANIAS
from core.siglas import detectar_siglas_sin_contexto


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


def _extraer_informes(ruta_docx):
    """Divide el Word en {clave: texto} usando los encabezados de compañía."""
    n2k = {cfg["nombre"]: k for k, cfg in CONFIG_COMPANIAS.items()}
    doc = Document(ruta_docx)
    informes, clave_actual, buf = {}, None, []

    def _save():
        if clave_actual and buf:
            informes[clave_actual] = "\n".join(buf)

    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        if t in n2k:
            _save()
            clave_actual = n2k[t]
            buf = []
        elif clave_actual:
            buf.append(t)
    _save()
    return informes


def main():
    ruta = sys.argv[1] if len(sys.argv) > 1 else _seleccionar_archivo()
    if not ruta:
        print("No se seleccionó archivo.")
        return 1
    ruta = str(Path(ruta))
    if not Path(ruta).is_file():
        print(f"[ERROR] Archivo no encontrado: {ruta}")
        return 1

    print(f"\n══ Búsqueda de siglas sin contexto ════════════════════════════")
    print(f"  Word: {Path(ruta).name}\n")

    informes = _extraer_informes(ruta)
    if not informes:
        print("  No se detectaron secciones de compañía en el Word.")
        return 1

    secciones = detectar_siglas_sin_contexto(informes)
    if not secciones:
        print("  ✓ No se encontraron siglas sin contexto.")
        return 0

    total = 0
    for sec in secciones:
        print(f"  {'─'*60}")
        print(f"  {sec['clave']} — {sec['n_siglas']} sigla(s) sin contexto")
        print(f"  {'─'*60}")
        for k in sec["kpis"]:
            total += 1
            print(f"    ✗ {k['label']}")
            print(f"        {k['detalle']}")
        print()

    print(f"{'═'*64}")
    print(f"  Total: {total} sigla(s) sin contexto en {len(secciones)} compañía(s)")
    print(f"{'═'*64}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
