# Automatización CYPECAD → Memoria de cálculo (Word/PDF)

Este repositorio contiene una utilidad para automatizar el paso de tablas y resultados desde un `.docx` exportado por CYPECAD hacia tu plantilla de memoria de cálculo en Word.

## ¿Qué resuelve?

En vez de copiar/pegar tabla por tabla:

1. Exportás la memoria/resultados desde CYPECAD a `docx`.
2. Definís un archivo de mapeo (`yaml`) con reglas simples:
   - qué texto buscar en el documento de CYPE,
   - qué tabla tomar luego de ese texto,
   - y en qué marcador de tu plantilla pegarla.
3. Ejecutás un comando y obtenés una memoria final en tu formato.

Opcionalmente, también podés convertir el resultado a PDF (si tenés LibreOffice instalado).

---

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
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
    table_offset_after_trigger: 1
    target_placeholder: "{{VIGAS_MOMENTOS}}"

  - id: vigas_flechas
    source_trigger_regex: "(?i)vigas.*flecha"
    table_offset_after_trigger: 1
    target_placeholder: "{{VIGAS_FLECHAS}}"

  - id: columnas_cuantia
    source_trigger_regex: "(?i)columnas.*cuant[ií]a"
    table_offset_after_trigger: 1
    target_placeholder: "{{COLUMNAS_CUANTIA}}"
```

### Parámetros de cada sección

- `id`: nombre interno.
- `source_trigger_regex`: regex para detectar el párrafo disparador en el documento CYPE.
- `table_offset_after_trigger`: qué tabla tomar después del párrafo detectado.
  - `1` = primera tabla posterior
  - `2` = segunda tabla posterior
- `target_placeholder`: marcador en plantilla de Word.

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
