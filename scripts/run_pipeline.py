from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from formbot.app.bootstrap import bootstrap_excel_pipeline
from formbot.domain.exceptions import FormBotError
from formbot.shared.utils import TraceabilityRegistry, configure_logging, format_trace_lines

LOGGER = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Pipeline de diligenciamiento de formularios Excel (label + offset)."
    )
    parser.add_argument("--template", required=True, type=Path, help="Ruta del template .xlsx")
    parser.add_argument("--mapping", required=True, type=Path, help="Ruta del mapping .yaml")
    parser.add_argument("--data", required=True, type=Path, help="Ruta del payload .json")
    parser.add_argument("--output", required=True, type=Path, help="Ruta del archivo de salida")
    parser.add_argument(
        "--log-level",
        default="INFO",
        help="Nivel de logging (DEBUG, INFO, WARNING, ERROR)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    configure_logging(args.log_level)
    traceability = TraceabilityRegistry(PROJECT_ROOT)
    traceability.ensure_structure()

    changed_scripts = traceability.register_script_changes([Path(__file__).resolve()])
    script_name = Path(__file__).name

    context = None
    try:
        LOGGER.info("1/7 Cargar documento y dependencias de pipeline")
        context = bootstrap_excel_pipeline(
            template_path=args.template.resolve(),
            mapping_path=args.mapping.resolve(),
            data_path=args.data.resolve(),
        )

        LOGGER.info("2/7 Analizar estructura y detectar etiquetas configuradas")
        LOGGER.info("3/7 Calcular posiciones objetivo por offset relativo")
        LOGGER.info("4/7 Insertar datos en el documento preservando estructura")
        result = context.use_case.execute(
            data=context.data,
            mapping_rules=context.mapping_rules,
            output_path=args.output.resolve(),
        )

        LOGGER.info("5/7 Validar resultado de mapeo")
        LOGGER.info("6/7 Exportar documento final")
        LOGGER.info("7/7 Registrar trazabilidad de ejecucion")

        observations = _build_success_observations(result, changed_scripts)
        traceability.append_execution(
            script_name=script_name,
            result="Exito",
            observations=observations,
        )
        LOGGER.info("Pipeline completado. Salida: %s", result.output_path)
        return 0
    except FormBotError as exc:
        message = f"{type(exc).__name__}: {exc}"
        LOGGER.error(message)
        traceability.append_execution(
            script_name=script_name,
            result="Error",
            observations=message,
        )
        return 1
    except Exception as exc:  # pragma: no cover - ruta defensiva.
        message = f"Error no controlado: {type(exc).__name__}: {exc}"
        LOGGER.exception(message)
        traceability.append_execution(
            script_name=script_name,
            result="Error",
            observations=message,
        )
        return 1
    finally:
        if context is not None:
            context.use_case.close()


def _build_success_observations(result, changed_scripts: list[Path]) -> str:
    written_fields = ", ".join(trace.field_name for trace in result.written_fields) or "Ninguno"
    changed_scripts_text = (
        ", ".join(path.name for path in changed_scripts) if changed_scripts else "Sin cambios"
    )
    return format_trace_lines(
        [
            f"Archivo generado: {result.output_path}",
            f"Campos escritos: {len(result.written_fields)} ({written_fields})",
            f"Campos opcionales omitidos: {', '.join(result.skipped_optional_fields) or 'Ninguno'}",
            f"Campos de entrada sin regla: {', '.join(result.unmapped_input_fields) or 'Ninguno'}",
            f"Scripts principales con cambio detectado: {changed_scripts_text}",
        ]
    )


if __name__ == "__main__":
    raise SystemExit(main())

