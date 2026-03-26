# FormBot

Sistema empresarial para diligenciamiento automatizado de formularios preservando estructura original.

## Entregable Implementado
- Adaptador Excel funcional (`openpyxl`)
- Mapeo inteligente por `label + offset` desde YAML
- Caso de uso `fill_form`
- Pipeline ejecutable `scripts/run_pipeline.py`
- Sistema de trazabilidad Markdown en `docs/`

## Instalacion
```bash
pip install -r requirements.txt
```

## Generar plantilla de ejemplo
```bash
python scripts/create_sample_excel_template.py
```

## Ejecutar pipeline
```bash
python scripts/run_pipeline.py ^
  --template examples/templates/form_template.xlsx ^
  --mapping config/mappings/excel_demo.yaml ^
  --data examples/data/input_demo.json ^
  --output outputs/form_filled.xlsx
```

## Frontend Web
Levantar interfaz web local:
```bash
pip install -r requirements.txt
python scripts/run_frontend.py
```

Link local del frontend:
- `http://127.0.0.1:8000`

## Llenado Masivo De Formularios Excel
Para diligenciar en lote archivos `.xlsx/.xlsm` con datos ASTECO y deteccion automatica por labels:
```bash
python scripts/run_bulk_autofill.py ^
  --input-dir "C:\ruta\formularios" ^
  --output-dir outputs\bulk_autofill ^
  --data config\data\asteco_master_profile.json ^
  --field-guide outputs\analysis_proveedores\canonical_field_guide.json ^
  --allow-overwrite-existing ^
  --min-confidence 0.78
```

Salida:
- Excel diligenciados en `--output-dir`
- Reporte consolidado en `--output-dir/bulk_autofill_report.json`

## Modo precision (cero-error operativo)
Este modo prioriza exactitud sobre volumen: si detecta ambiguedad o baja confianza, bloquea la escritura en `--strict`.

Ejecucion:
```bash
python scripts/run_pipeline_precise.py ^
  --template tests/fixtures/fixture_form.xlsx ^
  --mapping tests/fixtures/pipeline_mapping.yaml ^
  --data tests/fixtures/pipeline_data_valid.json ^
  --output outputs/precision_filled.xlsx ^
  --report outputs/precision_report.json ^
  --strict ^
  --min-confidence 0.85
```

En el reporte JSON por campo se registra:
- estado (`written`, `blocked`, `skipped`)
- confianza
- razon
- label y posicion usada
- posicion destino

Claves opcionales de precision en YAML por campo:
- `aliases`: lista de etiquetas alternativas
- `type`: `email`, `number`, `date`, `phone`, `nit`, etc.
- `target_strategy`: `offset`, `infer`, `offset_or_infer`
- `confidence_threshold`: umbral 0..1 por campo
- `write_mode`: `value` (default) o `mark` para marcar opciones tipo Si/No
- `mark_symbol`: simbolo a escribir en `write_mode: mark` (por ejemplo `X`, `✔`)

Comportamiento de `write_mode: mark`:
- Detecta opciones en la misma fila del label y tambien en filas vecinas (arriba/abajo).
- Para patrones `Si/No`, marca la celda vacia adyacente segun el valor recibido.
- Para opciones textuales (por ejemplo `Proveedor`, `Cliente`, `Accionista`), busca la opcion por texto y marca su zona adyacente.

### Ejemplo CU-FOR-001 (marcacion Si/No + campos texto)
Se incluye ejemplo de mapping y data para el formulario de contrapartes:
- `config/mappings/cu_for_001_contrapartes.yaml`
- `config/data/cu_for_001_contrapartes_profile.json`

Ejecucion:
```bash
python scripts/run_pipeline_precise.py ^
  --template "C:\ruta\CU-FOR-001 V.1 Formulario Contrapartes.xlsx" ^
  --mapping config/mappings/cu_for_001_contrapartes.yaml ^
  --data config/data/cu_for_001_contrapartes_profile.json ^
  --output outputs/cu_for_001_diligenciado.xlsx ^
  --report outputs/cu_for_001_report.json ^
  --min-confidence 0.80
```

## Ejecutar pruebas
```bash
pip install -r requirements.txt
python -m pytest -q
```

Fixtures usados por integration tests:
- `tests/fixtures/fixture_form.xlsx`
- `tests/fixtures/pipeline_mapping.yaml`
- `tests/fixtures/pipeline_data_valid.json`
- `tests/fixtures/pipeline_data_missing_required.json`

## Trazabilidad
- Decisiones: `docs/decisions/decision_log.md`
- Cambios: `docs/changes/change_log.md`
- Ejecuciones: `docs/runs/execution_log.md`
- Arquitectura: `docs/architecture.md`
