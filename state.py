"""Estado compartido entre módulos del generador de informes."""

# Acumula mensajes de validación y errores detectados durante la ejecución.
errores = []

# Guarda la instancia reutilizable de Excel abierta por COM.
_excel_app = None

# Mantiene cacheados los workbooks abiertos para evitar reabrirlos.
_workbooks_abiertos = {}

# Ruta del Word fuente de la faena que se está renderizando en este momento.
# La fija _construir_doc antes de llamar a cada procesador; permite a un renderer
# re-leer el Word original para recuperar metadatos que se pierden al aplanar el
# texto (ej. la numeración automática de listas en MLP Planta Desaladora).
ruta_word_actual = None
