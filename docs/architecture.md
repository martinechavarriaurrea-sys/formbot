# Arquitectura

## Objetivo
FormBot automatiza diligenciamiento de formularios empresariales preservando estructura original. No reconstruye documentos; abre la plantilla base y escribe solo valores en posiciones calculadas por reglas.

## Capas
- `domain`: contratos y modelos inmutables (`CellPosition`, `MappingRule`) + excepciones de negocio.
- `application`: caso de uso `FillFormUseCase`, coordinador del flujo de diligenciamiento.
- `infrastructure`: adaptadores concretos (Excel), parsers de entrada, y estrategia de mapeo.
- `app`: bootstrap de dependencias para pipeline.
- `shared`: utilidades transversales, logging y trazabilidad Markdown.

## Estrategia De Mapeo
Regla principal: `label + offset relativo`.

1. Se busca la etiqueta textual (`label`) en la hoja configurada o en todas.
2. Se obtiene su posicion base (`row`, `column`).
3. Se aplica offset (`row`, `col`) para calcular la celda destino.
4. Se escribe el valor en la celda destino sin alterar estilos.

Esta estrategia evita hardcode de coordenadas aisladas y permite absorber desplazamientos por celdas vacias, bloques y estructura compleja.

## Pipeline
1. Cargar documento base.
2. Cargar reglas YAML.
3. Cargar datos JSON.
4. Detectar etiqueta por regla.
5. Calcular posicion objetivo.
6. Escribir valores.
7. Guardar y registrar ejecucion en `docs/runs/execution_log.md`.

## Trazabilidad
- `docs/decisions/decision_log.md`: decisiones tecnicas y trade-offs (registro manual).
- `docs/changes/change_log.md`: cambios funcionales (registro manual y automatico para scripts principales).
- `docs/runs/execution_log.md`: historial de ejecuciones y errores del pipeline (automatico).

`TraceabilityRegistry` asegura estructura, agrega entradas Markdown y monitorea hash de scripts principales para detectar cambios automaticamente.

## Limitaciones Actuales
- El entregable implementa adaptador funcional solo para Excel (`openpyxl`).
- Adaptadores Word/PDF quedan como extension pendiente.
- La deteccion de etiquetas usa coincidencia exacta y parcial textual; no incluye OCR ni analisis semantico avanzado.
