# Generador de Informe Semanal — Antofagasta Minerals

Herramienta web para generar automáticamente el Informe Semanal de Operación de Antofagasta PLC a partir de los archivos Word por faena y el Excel madre de la semana.

---

## Uso general

### 1. Instalar dependencias

```bash
pip install -r requirements.txt
```

Requiere **Microsoft Excel** instalado en el equipo (se usa vía COM para exportar imágenes y actualizar vínculos).

### 2. Iniciar el panel

```bash
python server.py
```

Abre `http://localhost:5000` en el navegador.

### 3. Configurar la semana

En la pestaña **Configuración** ingresar:

- Número de semana y fechas (día/mes de inicio y fin)
- Si el disco compartido tiene una letra distinta a `N:`, activar "Disco compartido" e indicar la letra (ej. `Z:`)

El sistema detecta automáticamente todos los archivos dentro de la carpeta de la semana. Si alguno no se encuentra, se puede buscar manualmente con el botón "Buscar".

> Si existe un Excel `_act` (generado al actualizar vínculos), se usa ese en lugar del original.

> **Semanas especiales**: cuando la semana seleccionada tiene una duración distinta a 7 días —ya sea menor o mayor, calculado desde las fechas de inicio y fin del calendario— el panel muestra un aviso rojo recordando revisar las referencias Excel de CMZ, ya que suelen cambiar el nombre de la pestaña y la configuración de columnas/filas.

### 4. Generar el informe

En la pestaña **Generar** hay dos modos:

| Modo | Cuándo usarlo |
|---|---|
| **Word nuevo** | Primera generación de la semana |
| **Word existente** | Ya hay un Word preliminar y se quiere actualizar secciones puntuales |

En ambos modos se puede seleccionar qué faenas procesar y si incluir SSO / Gestión Hídrica.

Opcionalmente se pueden actualizar los vínculos del Excel madre antes de generar.

### 5. Validar KPIs

La pestaña **Validar KPIs** compara los valores numéricos del Word generado contra el Excel madre, faena por faena. Reporta diferencias de unidad y de porcentaje, e identifica KPIs ausentes o no reconocidos.

### 6. Revisar gramática

La pestaña **Gramática** analiza el Word en busca de errores de concordancia, género, palabras duplicadas y observaciones reportadas por la API de LanguageTool en español.

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

### Palabras clave para identificar los archivos

Dentro de cada carpeta el archivo se reconoce por una palabra clave contenida en su nombre
(por ejemplo `mlp semana` para el Excel de MLP o `bdatos` para los indicadores SSO). Los valores
de fábrica están en `utils/patrones.py` (`PATRONES_DEFAULT`) y se editan desde el panel, en
**Configuración → 🔎 Palabras clave de archivos**; lo que se guarda queda en `patrones.json`
en la raíz del proyecto (borrarlo devuelve todo a los valores por defecto).

- La comparación no distingue mayúsculas ni tildes y basta con que el nombre *contenga* la palabra.
- Se admiten varias alternativas separadas por coma: calza cualquiera de ellas.
- Palabra clave vacía = sin filtro; se toma el archivo solo si hay uno de esa extensión en la carpeta.
- El Word de cada faena viene sin palabra clave: por defecto se usa el único `.docx` de la carpeta.
- `ignorar_vinculos` lista los vínculos externos que nunca se actualizan (por defecto `cd mina`).

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
8. **Accidentabilidad Back-up** — tablas de la hoja `SSO` del Excel madre, una por página; se filtran las filas con ID de incidente = 0; los registros se ordenan por fecha primero y luego por ID de incidente

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
- Líneas de acumulado leídas directamente desde el Excel madre (celdas B62 y B63 de la hoja `CMZ`), no desde el Word de faena

#### FCAB
- Tabla de producción semanal
- Hechos Relevantes
- **Principales Desviaciones**: Tren · Camión

### Extracción de texto desde Word por faena

El texto de cada Word de faena se extrae por bloques delimitados por títulos de sección (`extraer_bloque`). Cada bloque se formatea según reglas propias del renderer de su faena:

- Las viñetas respetan niveles de jerarquía (1–4) usando estilos `Viñeta N` del template `.docx`
- Las fechas con formato `DD de Mes de AAAA:` se identifican y se rinden como encabezados en negrita dentro de listas blancas
- Los marcadores de viñeta del texto fuente (`•`, `·`, `○`, `o`) se preservan para clasificar el nivel correcto antes de ser reemplazados por el estilo Word
- Se detectan y reportan líneas no clasificadas por faena

### Secciones nuevas en el Word de faena

El informe tiene una estructura fija de secciones, pero a veces el redactor agrega un título propio
para algo excepcional de la semana (por ejemplo ANT sumó *Explicaciones por contingencia climática*
por las lluvias del 17 y 18 de agosto de 2026). Ese título **se conserva**: no se aplana a una viñeta más.

- Se detecta por **formato, no por texto**: un título nuevo es un párrafo del Word fuente que no es
  viñeta de lista, no lleva sangría, viene marcado como título y no aparece en
  `TITULOS_SECCION_CONOCIDOS` (`config.py`). Se ignora el encabezado del Word de faena — todo lo que
  esté antes de la primera sección conocida —, las frases destacadas que terminan en punto y las
  viñetas de KPI con forma `Etiqueta: (-1,880 kt; -22.4%) bajo PM...`, que son texto y no título.
- Qué cuenta como "marcado como título" **lo define cada Word**: si subraya los títulos que ya
  conocemos (`Detalle por fases`, `Mina`, `Accidentabilidad`...), entonces exige subrayado también
  para los nuevos; si no subraya ninguno, se conforma con la negrita. La negrita sola no alcanza como
  criterio general porque en los Word de faena también vienen en negrita las fechas de accidentes,
  las notas al pie y las bajadas de texto. Por eso conviene mantener completa
  `TITULOS_SECCION_CONOCIDOS`: es la que calibra la detección.
- En el informe se escribe **con el mismo formato que los demás títulos de sección** (azul `#00778B`,
  negrita, al margen y con dos puntos, igual que `Detalle por fases:`), no con el formato que traía en
  el Word de faena: lo que se conserva es el texto del título, no su presentación. Las viñetas que
  cuelgan de él sí conservan la jerarquía y la negrita del Word fuente, leídas del `w:ilvl` de la
  lista o de la sangría izquierda.
- El bloque completo se escribe **al final de la sección**, después de los acumulados, para no partir
  el bloque de KPIs.
- Siempre se avisa con un `[REVISAR] {FAENA}: SECCIÓN NUEVA en el Word de faena -> '{título}'`, aunque
  el renderer de esa sección no la conserve, para revisar ubicación y formato a mano.

> Conservar el título y la jerarquía está implementado en `procesar_seccion`, tanto en las secciones
> de texto libre (las que en `ORDEN_PRINCIPALES_DESVIACIONES` tienen `[""]`) como en las que tienen
> subtítulos (ej. ANT *Detalle por fases* con `["Extracción de Mineral", "Extracción de Lastre"]`).
> Las secciones con renderer propio (`ant_render_mina`, `mlp_render_medio_ambiente`, ...) y las de
> tipo `["?"]` quedan fuera: ahí el aviso sí se emite, pero el título se sigue renderizando con las
> reglas de esa sección.

### Alineación de viñetas con círculo blanco

Las viñetas manuales con círculo blanco (`○`) usan un **tab stop XML** inyectado vía `parse_xml` sobre el elemento `<w:pPr>` del párrafo, en lugar de espacios o NBSP. Esto garantiza que el texto siempre comience en la misma posición horizontal independientemente de cómo Word justifique los caracteres previos.

```python
def _agregar_tab_stop(p, posicion):
    tab_twips = round(posicion.pt * 20)
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    p._element.get_or_add_pPr().append(
        parse_xml(f'<w:tabs {ns}><w:tab w:val="left" w:pos="{tab_twips}"/></w:tabs>')
    )
```

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

Compara los valores numéricos del Word contra el Excel madre celda a celda (rangos configurados en `config.py` por compañía). Lógica de comparación:

- **KPIs normales**: compara el valor absoluto del Word con el del Excel, con tolerancia proporcional a la precisión decimal del texto (±0.6 para enteros, ±0.06 para 1 decimal, etc.). Solo se comparan los dos primeros valores de cada línea (UNID + %) para evitar falsos positivos con números extra en el texto (fases, fechas, etc.)
- **Líneas "Acumulado al..."**: solo se compara el porcentaje (último valor); la unidad física se descarta porque puede estar expresada como "mayor/menor producción". Los números se comparan en orden de aparición por posición relativa
- **Acumulados CMZ**: se leen desde celdas fijas (B62/B63 de la hoja `CMZ`) en lugar de escanear el UsedRange completo
- **Celdas con solo guiones** (`---`): se interpretan como 100% en el Excel para evitar falsos errores en KPIs sin variación
- **Guión no-separable** (U+2011): `_numeros_de_linea` lo normaliza a guión estándar antes de parsear, ya que `limpiar_texto_global` lo usa para evitar saltos de línea
- **Caracteres NBSP**: `limpiar_texto_global` inserta NBSP (U+00A0) antes del guión no-separable para prevenir saltos de línea en Word (ej. `-3.6` → ` ‑3.6`). El validador los elimina antes de extraer números

### Revisión gramatical

Analiza el Word párrafo a párrafo combinando dos motores:

1. **Reglas propias** (sin red): concordancia en número (adjetivo plural + sustantivo singular), concordancia de género (artículo/sustantivo), palabras o frases duplicadas consecutivas
2. **LanguageTool API** (`https://api.languagetool.org/v2/check`): análisis en español, con filtros para excluir categorías ruidosas (tipografía, estilo, ortografía pura) y un vocabulario de dominio minero que nunca se reporta como error

### Módulos

| Archivo | Responsabilidad |
|---|---|
| `server.py` | Panel web Flask, endpoints de la API, captura de stdout hacia el cliente |
| `main.py` | Orquestación del informe (Word nuevo y modo actualización) |
| `config.py` | Rutas base, rangos Excel, orden de faenas, configuración por compañía |
| `state.py` | Estado compartido: lista de errores, instancia Excel COM, workbooks abiertos |
| `validar.py` | Validador standalone con selector de archivos vía tkinter |
| `revisar_gramatica.py` | Revisión ortográfica y gramatical (reglas propias + LanguageTool API) |
| `core/extractores.py` | Extracción de bloques de texto desde Word por sección |
| `core/renderers.py` | Renderizadores Word por faena y por sección |
| `core/validador.py` | Validación KPIs Word vs Excel (lógica de comparación y reporte) |
| `utils/excel_utils.py` | COM Excel, exportación de imágenes, actualización de vínculos, lectura de acumulados por compañía |
| `utils/word_utils.py` | Escritura y formato de párrafos, viñetas, imágenes en Word |
| `utils/text_utils.py` | Normalización y limpieza de texto, inserción de NBSP, detección de fechas |
| `utils/patrones.py` | Palabras clave con que se identifican los Excel/Word en disco (defaults + `patrones.json`) |

---

## Construcción técnica

### Stack tecnológico

| Capa | Tecnología |
|---|---|
| **Lenguaje** | Python 3.11+ |
| **Interfaz web** | Flask 3.x (servidor local), HTML/CSS/JS vanilla |
| **Generación Word** | python-docx 1.x — estilos, párrafos, runs, XML directo vía `lxml` |
| **Lectura/escritura Excel** | openpyxl 3.x (lectura de datos), win32com / pywin32 (COM para imágenes y vínculos) |
| **Exportación de imágenes** | win32com → `CopyPicture` + chart temporal → PNG en `C:\Temp\` |
| **Procesamiento de imágenes** | Pillow (ImageGrab en fallback) |
| **Revisión gramatical** | `requests` → LanguageTool API pública (es) + reglas propias con `re` |
| **Selector de archivos** | tkinter (filedialog, modo standalone) |
| **Concurrencia** | `threading.Lock` para el buffer de logs del panel web |
| **Entorno** | Windows 10/11, Excel instalado localmente |

### Decisiones de diseño relevantes

**Generación Word por capas**
El documento se construye empezando desde un template `.docx` (`Template Viñetas Python.docx`) que define los estilos `Viñeta 1`–`Viñeta 4`. Los renderers añaden párrafos usando esos estilos en lugar de definirlos en código, lo que garantiza consistencia visual sin lógica de fuente/tamaño dispersa.

**Separación extractor / renderer**
`core/extractores.py` solo lee texto plano desde los Word de faena (por bloques delimitados por título). `core/renderers.py` recibe esos bloques y decide la estructura visual. Esto permite cambiar la lógica de renderizado sin tocar la extracción y viceversa.

**Tab stop XML para alineación de viñetas**
Las viñetas manuales con símbolo de círculo necesitan alinear el texto siempre al mismo punto horizontal. Usar espacios o NBSP fallaba al justificar el párrafo porque Word los estira. La solución es inyectar un `<w:tab w:val="left">` en el XML del párrafo y usar `\t` como separador, garantizando alineación exacta independiente del carácter previo.

**NBSP y guión no-separable para prevenir saltos de línea**
`limpiar_texto_global` inserta NBSP (U+00A0) antes del guión no-separable (U+2011, `‑`) para que Word no corte `‑3.6 %` o `‑7 kt` en medio de un valor. Se usa guión no-separable en lugar de guión estándar para que Word no lo interprete como guión de fin de línea. El validador de KPIs elimina los NBSP y normaliza el guión no-separable a guión estándar antes de aplicar regex numéricos.

**Comparación de KPIs acumulados por posición**
Para las líneas "Acumulado al mes / año", el Excel puede almacenar `-7` mientras el Word escribe `7.0` (con el signo implícito en "menor producción"). La comparación por conjunto de valores fallaba siempre. La solución es comparar solo el porcentaje (último valor numérico), en valor absoluto, descartando la unidad física. Para CMZ, los acumulados se leen desde celdas fijas (B62/B63 de la hoja `CMZ`) porque su formato no es compatible con el escaneo dinámico genérico.

**Panel web como thin wrapper**
`server.py` redirige `stdout` a un buffer compartido (`_Tee`) que los endpoints `/api/logs` exponen como stream de texto. Toda la lógica vive en los módulos `main.py`, `core/` y `utils/`; el servidor solo orquesta llamadas e importaciones. Esto permite ejecutar cualquier módulo standalone desde la terminal sin depender del servidor.

### Dependencias

```
python-docx>=1.1    # Generación y lectura de documentos Word (.docx)
openpyxl>=3.1       # Lectura de Excel sin COM (datos y rangos)
pywin32>=306        # Automatización COM de Excel (imágenes, vínculos)
Pillow>=10.0        # Procesamiento de imágenes PNG
requests>=2.31      # Llamadas a la API de LanguageTool
flask>=3.0          # Servidor web del panel de control
```

> `tkinter` viene incluido con la instalación estándar de Python en Windows y no requiere instalación adicional.
>
> Microsoft Excel debe estar instalado localmente. El COM de Excel se inicializa con `win32com.client.DispatchEx("Excel.Application")`.
