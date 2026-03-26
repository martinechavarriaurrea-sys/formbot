from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from formbot.application.precision_fill import PrecisionFillUseCase
from formbot.domain.exceptions import FormBotError
from formbot.infrastructure.parsers.json_data_provider import JsonFileDataProvider
from formbot.infrastructure.parsers.yaml_mapping_provider import YamlMappingProvider
from formbot.shared.utils import configure_logging

LOGGER = logging.getLogger(__name__)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Diligenciamiento de alta precision para Excel: "
            "si hay ambiguedad o baja confianza, bloquea en modo estricto."
        )
    )
    parser.add_argument("--template", required=True, type=Path, help="Ruta del template .xlsx/.xlsm")
    parser.add_argument("--mapping", required=True, type=Path, help="Ruta del mapping .yaml")
    parser.add_argument("--data", required=True, type=Path, help="Ruta del payload .json")
    parser.add_argument("--output", required=True, type=Path, help="Ruta del archivo de salida")
    parser.add_argument(
        "--report",
        type=Path,
        default=PROJECT_ROOT / "outputs" / "precision_report.json",
        help="Ruta del reporte JSON con decisiones por campo",
    )
    parser.add_argument(
        "--min-confidence",
        type=float,
        default=0.85,
        help="Confianza minima global para escribir (0..1)",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Activa modo estricto (si hay bloqueos, no escribe ningun campo).",
    )
    parser.add_argument(
        "--allow-overwrite-existing",
        action="store_true",
        help="Permite sobreescribir celdas no vacias (reduce seguridad).",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        help="Nivel de logging (DEBUG, INFO, WARNING, ERROR)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    configure_logging(args.log_level)

    report_path = args.report.resolve()
    report_path.parent.mkdir(parents=True, exist_ok=True)

    mapping_provider = YamlMappingProvider(file_path=args.mapping.resolve())
    data_provider = JsonFileDataProvider(file_path=args.data.resolve())
    use_case = PrecisionFillUseCase(
        template_path=args.template.resolve(),
        strict_mode=args.strict,
        min_confidence=args.min_confidence,
        allow_overwrite_existing=args.allow_overwrite_existing,
    )

    try:
        mapping_rules = mapping_provider.load()
        data = data_provider.load()
        result = use_case.execute(
            data=data,
            mapping_rules=mapping_rules,
            output_path=args.output.resolve(),
        )
        report = {
            "status": "success",
            "strict_mode": result.strict_mode,
            "min_confidence": result.min_confidence,
            "output_path": str(result.output_path),
            "written_fields": result.written_fields,
            "blocked_fields": result.blocked_fields,
            "decisions": [
                {
                    "field_name": item.field_name,
                    "status": item.status,
                    "confidence": item.confidence,
                    "reason": item.reason,
                    "label": item.label,
                    "label_position": _position_to_dict(item.label_position),
                    "target_position": _position_to_dict(item.target_position),
                    "value_preview": item.value_preview,
                }
                for item in result.decisions
            ],
        }
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        LOGGER.info("Pipeline de precision completado. Salida: %s", result.output_path)
        LOGGER.info("Reporte: %s", report_path)
        return 0
    except FormBotError as exc:
        message = f"{type(exc).__name__}: {exc}"
        report = {
            "status": "error",
            "error": message,
            "strict_mode": args.strict,
            "min_confidence": args.min_confidence,
        }
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        LOGGER.error(message)
        LOGGER.error("Reporte: %s", report_path)
        return 1
    except Exception as exc:  # pragma: no cover - ruta defensiva.
        message = f"Error no controlado: {type(exc).__name__}: {exc}"
        report = {
            "status": "error",
            "error": message,
            "strict_mode": args.strict,
            "min_confidence": args.min_confidence,
        }
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        LOGGER.exception(message)
        LOGGER.error("Reporte: %s", report_path)
        return 1
    finally:
        use_case.close()


def _position_to_dict(position) -> dict | None:
    if position is None:
        return None
    return {
        "sheet_name": position.sheet_name,
        "row": position.row,
        "column": position.column,
    }


if __name__ == "__main__":
    raise SystemExit(main())

