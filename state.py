"""Estado compartido entre módulos del generador de informes."""

# Acumula mensajes de validación y errores detectados durante la ejecución.
errores = []

# Guarda la instancia reutilizable de Excel abierta por COM.
_excel_app = None

# Mantiene cacheados los workbooks abiertos para evitar reabrirlos.
_workbooks_abiertos = {}
