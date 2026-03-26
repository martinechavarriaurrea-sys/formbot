# Fixtures Reales

Coloca aqui los archivos reales para pruebas.

## Excel base
- Cualquier nombre `.xlsx` o `.xlsm` (el fixture `excel_path` toma el primero).
- Opcionalmente define `FORMBOT_TEST_EXCEL=<nombre_archivo>` para fijar uno.

## Requeridos para integration tests
- `pipeline_mapping.yaml`
- `pipeline_data_valid.json`
- `pipeline_data_missing_required.json`

Si faltan estos archivos, los tests de integracion quedan en `SKIP` automatico.

## Fixtures incluidos por defecto en el repositorio
- `fixture_form.xlsx`
- `pipeline_mapping.yaml`
- `pipeline_data_valid.json`
- `pipeline_data_missing_required.json`
