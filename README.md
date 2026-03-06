# Automatización CYPECAD → Memoria de cálculo (Word/PDF)

Este repositorio contiene una utilidad para automatizar el paso de tablas y resultados desde un `.docx` exportado por CYPECAD hacia una plantilla de memoria de cálculo en Word.

## ¿Qué resuelve?

En vez de copiar/pegar tabla por tabla:

1. Se exporta la memoria/resultados desde CYPECAD a `docx`.
2. Se define un archivo de mapeo (`yaml`) con reglas simples:
   - qué texto buscar en párrafos del documento de CYPE,
   - qué texto buscar en el encabezado de la tabla,
   - qué tabla tomar según el orden,
   - y en qué marcador de la plantilla pegarla.
3. Se ejecuta un comando como botón y se obtiene una memoria final en formato.

---

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Guía paso a paso desde cero (sin tener nada instalado)

> Guía está pensada para Windows.

### 1) Instalar Python

1. Entrar a: https://www.python.org/downloads/
2. Descargar **Python 3.11 o 3.12**.
3. Durante la instalación, marcar la opción **"Add Python to PATH"**.
4. Verificar en una terminal (CMD o PowerShell):

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
Si tira error en activar Scripts ejecutá: ```Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned```
Despues activas Scripts y requirements.

### 5) Preparar tus 3 archivos de entrada

Dentro de la carpeta `datos/` vas a dejar:

1. `fuentes/` (carpeta) → ahí ponés **todos** los DOCX exportados por CYPECAD (5 o 6 si querés).
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
- `table_offset_after_trigger`: índice inicial de tabla válida (1, 2, 3...).
- `match_mode` (opcional): `single` (default) o `all`.
  - `single`: inserta una sola tabla (comportamiento fijo).
  - `all`: inserta todas las tablas coincidentes desde ese índice (ideal para vigas con N variable).
- `target_placeholder`: marcador de tu plantilla.

### 8) Ejecutar la generación de la memoria final

```powershell
python src/cype_memoria_automation.py `
  --source-dir datos/fuentes `
  --template-docx datos/plantilla_memoria.docx `
  --mapping-yaml datos/mapeo.yaml `
  --output-docx salida/memoria_final.docx
```

Resultado: `salida/memoria_final.docx`


### 9) Flujo operativo diario (resumen)

1. Calculás en CYPECAD.
2. Exportás DOCX de resultados.
3. Reemplazás/actualizás los DOCX dentro de `datos/fuentes/`.
4. Ejecutás el comando con el botón.
5. Te queda la memoria en tu formato, con tablas cargadas automáticamente.

---

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
    match_mode: "all"
    target_placeholder: "{{VIGAS_MOMENTOS}}
```

### Explicaci+on de parámetros de cada sección del yaml.

- `id`: nombre interno.
- `source_trigger_regex` (opcional): regex para detectar el párrafo disparador en el documento CYPE.
- `table_header_regex` (opcional): regex para detectar el texto del **encabezado (primera fila)** de la tabla.
- `table_offset_after_trigger`:
  - si hay `source_trigger_regex`: toma la N-ésima tabla válida después del trigger,
  - si NO hay `source_trigger_regex`: toma la N-ésima tabla válida en todo el documento.
- `target_placeholder`: marcador en plantilla de Word.

> Debe existir al menos uno de estos dos campos: `source_trigger_regex` o `table_header_regex`.

## Limitaciones

- Si CYPE cambia el texto de encabezados, ajustá los regex en el YAML.
- Si el marcador no existe en la plantilla, la sección se informa como warning.
- Este flujo está pensado para memorias "tabulares" (sin análisis textual complejo).

