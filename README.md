# Generador de Informe Semanal — Antofagasta Minerals

Herramienta web para generar automáticamente el Informe Semanal de Operación de Antofagasta PLC a partir de los archivos Word por faena y el Excel madre de la semana.

---

## Uso general

### 1. Iniciar el panel

```bash
python server.py
```

Abre `http://localhost:5000` en el navegador.

### 2. Configurar la semana

En la pestaña **Configuración** ingresar:

- Número de semana y fechas (día/mes de inicio y fin)
- Si el disco compartido tiene una letra distinta a `N:`, activar "Disco compartido" e indicar la letra (ej. `Z:`)

El sistema detecta automáticamente todos los archivos dentro de la carpeta de la semana. Si alguno no se encuentra, se puede buscar manualmente con el botón "Buscar".

> Si existe un Excel `_act` (generado al actualizar vínculos), se usa ese en lugar del original.

### 3. Generar el informe

En la pestaña **Generar** hay dos modos:

| Modo | Cuándo usarlo |
|---|---|
| **Word nuevo** | Primera generación de la semana |
| **Word existente** | Ya hay un Word preliminar y se quiere actualizar secciones puntuales |

En ambos modos se puede seleccionar qué faenas procesar y si incluir SSO / Gestión Hídrica.

Opcionalmente se pueden actualizar los vínculos del Excel madre antes de generar.

### 4. Validar KPIs

La pestaña **Validar KPIs** compara los valores numéricos del Word generado contra el Excel madre, faena por faena.

### 5. Revisar gramática

La pestaña **Gramática** analiza el Word en busca de errores de concordancia, género, palabras duplicadas y ortografía.

---

## Detalle técnico

### Estructura de carpetas esperada

```
N:\01 Reporting\09 Informe Semanal\
  {año}\
    {N} - {Mes}\
      {num}_Semana- {dd} {mmm} al {dd} {mmm}\        ← carpeta raíz de la semana
        Semana {N} -  {dd} {mmm} al {dd} {mmm}.xlsx  ← Excel madre
        Semana {N} -  {dd} {mmm} al {dd} {mmm}_act.xlsx  ← Excel actualizado (opcional)
        01 -MLP\      ← Word informe + Excel vinculado MLP
        02 -CEN\
        03 -ANT\
        04 -CMZ\
        05 -FCAB\
        06 -SSO\      ← BDatos*.xlsx (Excel indicadores SSO)
        07 -{nombre}\ ← carpeta Gestión Hídrica
```

Si la semana cruza dos meses se busca primero en la carpeta del mes de cierre y luego en la del mes de inicio.

### Estructura del documento generado

El Word se construye en este orden:

1. **Título** — "Informe Semanal de Operación - Antofagasta PLC"
2. **Aviso de información incompleta** — aparece si hay faenas sin texto o si falta SSO / Gestión Hídrica
3. **Resumen ejecutivo** — texto extraído de la hoja `Grupo Minero FCAB PLAN` del Excel madre
4. **Tabla resumen** — imagen del rango `A3:X34` de la hoja `Grupo Minero FCAB PLAN`
5. **Gestión Hídrica** — imagen del rango `A3:W20` de la hoja `Gestión Hídrica` del Excel madre
6. **Accidentabilidad** — tres tablas de indicadores desde el Excel SSO (`BDatos*.xlsx`), hoja `Informe Viernes`:
   - Valor Anual (`A1:M13`)
   - Valor Mensual (`A15:M27`)
   - Valor Semanal (`A29:M41`)
7. **Sección por faena** (en orden: MLP, CEN, ANT, CMZ, FCAB) — ver detalle más abajo
8. **Accidentabilidad Back-up** — tablas de la hoja `SSO` del Excel madre, una por página; se filtran las filas con ID de incidente = 0

Al final se escribe el pie de página con el rango de fechas de la semana.

### Secciones por faena

Cada faena ocupa su propia página y contiene:

#### MLP — Los Pelambres
- Tabla de producción semanal (Excel madre)
- **Hechos Relevantes**: Accidentabilidad · Reportabilidad · Gestión SSO · Salud Ocupacional y Gestión Vial · Medio Ambiente · Asuntos Públicos
- **Principales Desviaciones**: Mina · Concentradora · Planta Desaladora · Gestión Hídrica
- Tablas de desviaciones (imagen desde rango configurable del Excel madre)

#### CEN — Centinela
- Tabla de producción semanal
- Hechos Relevantes (ídem MLP)
- **Principales Desviaciones**: Mina · Sulfuros · Cátodos

#### ANT — Antucoya
- Tabla de producción semanal
- Hechos Relevantes
- **Principales Desviaciones**: Mina · Planta

#### CMZ — Zaldívar
- Tabla de producción semanal
- Hechos Relevantes
- **Principales Desviaciones**: Mina · Planta

#### FCAB
- Tabla de producción semanal
- Hechos Relevantes
- **Principales Desviaciones**: Tren · Camión

### Extracción de texto desde Word por faena

El texto de cada Word de faena se extrae por bloques delimitados por títulos de sección. Cada bloque se formatea según reglas propias:

- Las viñetas respetan niveles de jerarquía (1–4)
- Las fechas (`DD de Mes de AAAA`) se identifican y formatean en negrita
- Los porcentajes se validan: deben tener exactamente 1 decimal y sin espacio antes de `%`
- Se detectan y reportan líneas no clasificadas por faena

### Exportación de imágenes desde Excel

Todas las imágenes se generan vía COM (`win32com`) y se guardan como PNG en `C:\Temp\`:

- **Tablas generales**: `CopyPicture` → chart temporal → `Chart.Export` → borrar chart (sin portapapeles, sin esperas)
- **Tablas SSO backup**: igual al anterior, pero primero se ocultan las filas con ID de incidente = 0 y se restauran al terminar

### Actualización de vínculos Excel

Antes de generar se puede actualizar el Excel madre con los datos de cada faena. El proceso:

1. Abre cada Excel de faena (subcarpetas `01-MLP`, `02-CEN`, etc.)
2. Copia los datos al Excel madre vía COM
3. Guarda una copia como `{nombre}_act.xlsx` en la carpeta destino
4. El informe usa ese archivo `_act` si existe

### Validación de KPIs

Compara los valores numéricos del Word contra el Excel madre celda a celda (rangos configurados en `config.py` por compañía). Reporta diferencias en unidades y porcentaje.

### Revisión gramatical

Analiza el Word párrafo a párrafo en busca de:
- Palabras repetidas consecutivas
- Errores de concordancia en número (singular/plural)
- Errores de concordancia de género (artículo/sustantivo)
- Palabras desconocidas (ortografía)

### Módulos

| Archivo | Responsabilidad |
|---|---|
| `server.py` | Panel web Flask, endpoints de la API |
| `main.py` | Orquestación del informe (Word nuevo y modo actualización) |
| `config.py` | Rutas base, rangos Excel, orden de faenas, configuración por compañía |
| `state.py` | Estado compartido: lista de errores, instancia Excel COM, workbooks abiertos |
| `core/extractores.py` | Extracción de bloques de texto desde Word por sección |
| `core/renderers.py` | Renderizadores Word por faena y por sección |
| `core/validador.py` | Validación KPIs Word vs Excel |
| `utils/excel_utils.py` | COM Excel, exportación de imágenes, actualización de vínculos |
| `utils/word_utils.py` | Escritura y formato de párrafos, viñetas, imágenes en Word |
| `utils/text_utils.py` | Normalización y limpieza de texto |
| `revisar_gramatica.py` | Revisión ortográfica y gramatical |

### Dependencias

```bash
pip install flask python-docx openpyxl pywin32 Pillow pymupdf pyspellchecker
```

Requiere Microsoft Excel instalado en el equipo (se usa vía COM para exportar imágenes y actualizar vínculos).
