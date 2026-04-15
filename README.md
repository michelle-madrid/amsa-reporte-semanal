# Generador de informe semanal modularizado

## Estructura
- `config.py`: constantes, rutas y configuración base.
- `state.py`: estado compartido de errores y Excel COM.
- `utils/text_utils.py`: limpieza y normalización de texto.
- `utils/word_utils.py`: helpers de formato y escritura en Word.
- `utils/excel_utils.py`: selección de archivos, Excel COM y exportación de imágenes.
- `core/extractores.py`: extracción de texto y bloques desde Word.
- `core/renderers.py`: renderizadores por faena y procesadores principales.
- `main.py`: orquestación completa del informe.

## Ejecución
```bash
python main.py
```
