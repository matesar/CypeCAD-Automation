# Automatización CYPECAD → Memoria de cálculo (Word/PDF)

Este repositorio contiene una utilidad para automatizar el paso de tablas y resultados desde un `.docx` exportado por CYPECAD hacia tu plantilla de memoria de cálculo en Word.

## ¿Qué resuelve?

En vez de copiar/pegar tabla por tabla:

1. Exportás la memoria/resultados desde CYPECAD a `docx`.
2. Definís un archivo de mapeo (`yaml`) con reglas simples:
   - qué texto buscar en párrafos del documento de CYPE (**opcional**),
   - qué texto buscar en el encabezado de la tabla (**opcional y recomendado**),
   - qué tabla tomar según el orden,
   - y en qué marcador de tu plantilla pegarla.
3. Ejecutás un comando y obtenés una memoria final en tu formato.

Opcionalmente, también podés convertir el resultado a PDF.

---

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Guía paso a paso desde cero (sin tener nada instalado)

> Esta guía está pensada para Windows (caso más común para CYPECAD).

### 1) Instalar Python

1. Entrá a: https://www.python.org/downloads/
2. Descargá **Python 3.11 o 3.12**.
3. Durante la instalación, marcá la opción **"Add Python to PATH"**.
4. Verificá en una terminal (CMD o PowerShell):

```powershell
python --version
pip --version
```

Si estos comandos responden con versión, está OK.

### 2) Descargar este proyecto

#### Opción A — Sin Git (más fácil)

1. Abrí en el navegador la página del repositorio.
2. Hacé clic en **Code** → **Download ZIP**.
3. Guardá el ZIP en una carpeta fácil.
4. Descomprimí el ZIP.
5. Entrá a la carpeta descomprimida.
6. Dentro de esa carpeta, abrí terminal:
   - clic en la barra de dirección del explorador,
   - escribí `powershell`,
   - presioná Enter.

Comprobación rápida (deberías ver `README.md`, `src`, `datos`):

```powershell
dir
```

### 3) Crear entorno virtual e instalar dependencias

En la carpeta del proyecto:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Si todo salió bien, ya está listo para correr.

### 5) Preparar tus 3 archivos de entrada

Dentro de la carpeta `datos/` vas a dejar:

1. `cype_export.docx` → el Word exportado por CYPECAD.
2. `plantilla_memoria.docx` → tu plantilla personal de memoria.
3. `mapeo.yaml` → reglas para saber qué tabla va en qué marcador.

Podés partir del ejemplo:

```powershell
copy datos\mapeo.example.yaml datos\mapeo.yaml
```

### 6) Preparar placeholders en tu plantilla Word

En `plantilla_memoria.docx`, escribí marcadores como texto (uno por párrafo), por ejemplo:

- `{{VIGAS_MOMENTOS}}`
- `{{VIGAS_FLECHAS}}`
- `{{COLUMNAS_CUANTIA}}`
- `{{FUNDACIONES_VERIFICACION}}`

Esos textos deben coincidir exactamente con `target_placeholder` en el YAML.

### 7) Ajustar `mapeo.yaml`

Para cada bloque, definí:

- `source_trigger_regex` (opcional): texto/regex que aparece antes de la tabla en el DOCX de CYPE.
- `table_header_regex` (opcional): regex del encabezado de tabla (primera fila) para reconocerla aunque no tenga título arriba.
- `table_offset_after_trigger`: índice de tabla válida (1, 2, 3...).
- `target_placeholder`: marcador de tu plantilla.

### 8) Ejecutar la generación de la memoria final

```powershell
python src/cype_memoria_automation.py `
  --source-docx datos/cype_export.docx `
  --template-docx datos/plantilla_memoria.docx `
  --mapping-yaml datos/mapeo.yaml `
  --output-docx salida/memoria_final.docx
```

Resultado: `salida/memoria_final.docx`

### 9) Ejecutar también con PDF automático

```powershell
python src/cype_memoria_automation.py `
  --source-docx datos/cype_export.docx `
  --template-docx datos/plantilla_memoria.docx `
  --mapping-yaml datos/mapeo.yaml `
  --output-docx salida/memoria_final.docx `
  --output-pdf
```

Resultado adicional: `salida/memoria_final.pdf`

### 10) Flujo operativo diario (resumen)

1. Calculás en CYPECAD.
2. Exportás DOCX de resultados.
3. Reemplazás `datos/cype_export.docx` por el nuevo.
4. Ejecutás el comando.
5. Te queda la memoria en tu formato, con tablas cargadas automáticamente.

---

## Linux/macOS (comandos equivalentes rápidos)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

python src/cype_memoria_automation.py \
  --source-docx datos/cype_export.docx \
  --template-docx datos/plantilla_memoria.docx \
  --mapping-yaml datos/mapeo.yaml \
  --output-docx salida/memoria_final.docx
```

## Estructura recomendada

- `datos/cype_export.docx`: memoria exportada por CYPECAD.
- `datos/plantilla_memoria.docx`: tu plantilla Word con marcadores.
- `datos/mapeo.yaml`: reglas de extracción/inserción.
- `salida/memoria_final.docx`: resultado generado.

## Marcadores en plantilla

En tu plantilla Word, insertá marcadores de texto únicos, por ejemplo:

- `{{VIGAS_MOMENTOS}}`
- `{{VIGAS_FLECHAS}}`
- `{{COLUMNAS_CUANTIA}}`
- `{{FUNDACIONES_VERIFICACION}}`

Cada marcador debe estar en un párrafo propio para simplificar el reemplazo.

## Ejemplo de `mapeo.yaml`

```yaml
sections:
  - id: vigas_momentos
    source_trigger_regex: "(?i)vigas.*momento"
    table_header_regex: "(?i)tramo|M\+|M-"
    table_offset_after_trigger: 1
    target_placeholder: "{{VIGAS_MOMENTOS}}"

  - id: vigas_flechas
    source_trigger_regex: "(?i)vigas.*flecha"
    table_header_regex: "(?i)flecha|L/"
    table_offset_after_trigger: 1
    target_placeholder: "{{VIGAS_FLECHAS}}"

  - id: fundaciones_verificacion
    table_header_regex: "(?i)zapata|tensi[oó]n|punzonamiento"
    table_offset_after_trigger: 1
    target_placeholder: "{{FUNDACIONES_VERIFICACION}}"
```

### Parámetros de cada sección

- `id`: nombre interno.
- `source_trigger_regex` (opcional): regex para detectar el párrafo disparador en el documento CYPE.
- `table_header_regex` (opcional): regex para detectar el texto del **encabezado (primera fila)** de la tabla.
- `table_offset_after_trigger`:
  - si hay `source_trigger_regex`: toma la N-ésima tabla válida después del trigger,
  - si NO hay `source_trigger_regex`: toma la N-ésima tabla válida en todo el documento.
- `target_placeholder`: marcador en plantilla de Word.

> Debe existir al menos uno de estos dos campos: `source_trigger_regex` o `table_header_regex`.

## Uso

```bash
python src/cype_memoria_automation.py \
  --source-docx datos/cype_export.docx \
  --template-docx datos/plantilla_memoria.docx \
  --mapping-yaml datos/mapeo.yaml \
  --output-docx salida/memoria_final.docx
```

### Exportar también a PDF

```bash
python src/cype_memoria_automation.py \
  --source-docx datos/cype_export.docx \
  --template-docx datos/plantilla_memoria.docx \
  --mapping-yaml datos/mapeo.yaml \
  --output-docx salida/memoria_final.docx \
  --output-pdf
```

> `--output-pdf` requiere `soffice` (LibreOffice) disponible en PATH.

## Limitaciones y buenas prácticas

- Si CYPE cambia el texto de encabezados, ajustá los regex en el YAML.
- Si el marcador no existe en la plantilla, la sección se informa como warning.
- Este flujo está pensado para memorias "tabulares" (sin análisis textual complejo), justo el caso que describiste.

## Próximos pasos sugeridos

- Agregar reglas por tipo de proyecto (hormigón/metal/fundaciones especiales).
- Incorporar una plantilla maestra por empresa.
- Integrar este script al pipeline posterior al cálculo para que salga la memoria automáticamente.

