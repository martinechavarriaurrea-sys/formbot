from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from formbot.application.precision_fill import PrecisionDecision, PrecisionFillUseCase
from formbot.domain.exceptions import FormBotError
from formbot.domain.models import MappingRule
from formbot.shared.utils import configure_logging

LOGGER = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}

VALUE_TYPE_BY_FIELD: dict[str, str] = {
    "numero_identificacion_nit": "nit",
    "digito_verificacion": "number",
    "representante_legal_documento": "number",
    "telefono_fijo": "phone",
    "celular": "phone",
    "numero_cuenta": "number",
    "correo_electronico": "email",
    "fecha_diligenciamiento": "date",
}

EXTRA_ALIASES: dict[str, tuple[str, ...]] = {
    "correo_electronico": (
        "correo de contacto (comercial)",
        "correo de contacto (cartera)",
        "correo electronico para notificacion de pago",
        "email de contacto",
    ),
    "contacto_nombre": (
        "contacto comercial",
        "contacto pagos",
        "nombre del contacto",
    ),
    "representante_legal_nombre": (
        "nombre representante legal",
        "nombres y apellidos del representante legal",
    ),
    "representante_legal_documento": (
        "identificacion no.",
        "numero identificacion representante legal",
        "documento representante legal",
    ),
    "numero_identificacion_nit": (
        "numero de identificacion",
        "identificacion tributaria",
    ),
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Diligencia en lote formularios Excel usando deteccion por etiquetas y aliases "
            "basados en guia canonica de campos."
        )
    )
    parser.add_argument(
        "--input-dir",
        required=True,
        type=Path,
        help="Directorio raiz con archivos .xlsx/.xlsm a diligenciar",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        type=Path,
        help="Directorio donde se guardaran los formularios diligenciados",
    )
    parser.add_argument(
        "--data",
        type=Path,
        default=PROJECT_ROOT / "config" / "data" / "asteco_master_profile.json",
        help="JSON con datos canonicos de la empresa",
    )
    parser.add_argument(
        "--field-guide",
        type=Path,
        default=PROJECT_ROOT
        / "outputs"
        / "analysis_proveedores"
        / "canonical_field_guide.json",
        help="Guia de campos canonicos con variantes de labels",
    )
    parser.add_argument(
        "--report",
        type=Path,
        default=None,
        help="Ruta del reporte de ejecucion JSON",
    )
    parser.add_argument(
        "--min-docs",
        type=int,
        default=15,
        help="Minimo de documentos detectados para priorizar reglas del field-guide",
    )
    parser.add_argument(
        "--max-files",
        type=int,
        default=None,
        help="Limita cantidad de formularios a procesar (util para pruebas)",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Si se activa, un bloqueo en un campo evita escritura de ese formulario.",
    )
    parser.add_argument(
        "--min-confidence",
        type=float,
        default=0.78,
        help="Confianza minima global para escribir (0..1)",
    )
    parser.add_argument(
        "--allow-overwrite-existing",
        action="store_true",
        help="Permite sobreescribir celdas que ya tienen contenido",
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

    input_dir = args.input_dir.resolve()
    output_dir = args.output_dir.resolve()
    data_path = args.data.resolve()
    guide_path = args.field_guide.resolve()

    if not input_dir.exists() or not input_dir.is_dir():
        LOGGER.error("Directorio de entrada no valido: %s", input_dir)
        return 1
    if not data_path.exists():
        LOGGER.error("No existe archivo de datos: %s", data_path)
        return 1
    if not guide_path.exists():
        LOGGER.error("No existe guia de campos: %s", guide_path)
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)

    report_path = (
        args.report.resolve()
        if args.report is not None
        else output_dir / "bulk_autofill_report.json"
    )
    report_path.parent.mkdir(parents=True, exist_ok=True)

    data = _load_json_dict(data_path)
    guide = _load_json_dict(guide_path)
    rules = _build_mapping_rules(data, guide, min_docs=args.min_docs)
    if not rules:
        LOGGER.error("No se pudieron construir reglas de mapeo desde datos/guia.")
        return 1

    excel_files = _discover_excel_files(input_dir=input_dir, output_dir=output_dir)
    if args.max_files is not None:
        excel_files = excel_files[: args.max_files]

    if not excel_files:
        LOGGER.error("No se encontraron archivos Excel en %s", input_dir)
        return 1

    LOGGER.info(
        "Inicio lote: archivos=%d, reglas=%d, strict=%s, min_confidence=%.2f",
        len(excel_files),
        len(rules),
        args.strict,
        args.min_confidence,
    )

    items: list[dict[str, Any]] = []
    success_count = 0

    for index, template_path in enumerate(excel_files, start=1):
        output_path = _build_output_path(
            input_root=input_dir,
            template_path=template_path,
            output_root=output_dir,
        )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        LOGGER.info("[%d/%d] %s", index, len(excel_files), template_path)

        use_case = PrecisionFillUseCase(
            template_path=template_path,
            strict_mode=args.strict,
            min_confidence=args.min_confidence,
            allow_overwrite_existing=args.allow_overwrite_existing,
        )
        try:
            result = use_case.execute(
                data=data,
                mapping_rules=rules,
                output_path=output_path,
            )
            success_count += 1
            items.append(
                {
                    "template_path": str(template_path),
                    "output_path": str(output_path),
                    "status": "success",
                    "written_fields": result.written_fields,
                    "blocked_fields": result.blocked_fields,
                    "decisions": [_decision_to_dict(item) for item in result.decisions],
                }
            )
        except FormBotError as exc:
            LOGGER.warning("Formulario bloqueado: %s", exc)
            items.append(
                {
                    "template_path": str(template_path),
                    "output_path": str(output_path),
                    "status": "error",
                    "error": f"{type(exc).__name__}: {exc}",
                }
            )
        except Exception as exc:  # pragma: no cover - defensa operativa
            LOGGER.exception("Fallo no controlado procesando %s", template_path)
            items.append(
                {
                    "template_path": str(template_path),
                    "output_path": str(output_path),
                    "status": "error",
                    "error": f"{type(exc).__name__}: {exc}",
                }
            )
        finally:
            use_case.close()

    report = {
        "generated_at": datetime.now().isoformat(),
        "input_dir": str(input_dir),
        "output_dir": str(output_dir),
        "data_path": str(data_path),
        "field_guide_path": str(guide_path),
        "strict_mode": args.strict,
        "min_confidence": args.min_confidence,
        "allow_overwrite_existing": args.allow_overwrite_existing,
        "files_total": len(excel_files),
        "files_success": success_count,
        "files_error": len(excel_files) - success_count,
        "rules_count": len(rules),
        "rules_fields": [rule.field_name for rule in rules],
        "items": items,
    }
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    LOGGER.info(
        "Lote finalizado: ok=%d, error=%d, reporte=%s",
        success_count,
        len(excel_files) - success_count,
        report_path,
    )
    return 0 if success_count > 0 else 1


def _load_json_dict(path: Path) -> dict[str, Any]:
    content = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(content, dict):
        raise ValueError(f"JSON invalido (se esperaba objeto): {path}")
    return content


def _build_mapping_rules(
    data: dict[str, Any],
    field_guide: dict[str, Any],
    *,
    min_docs: int,
) -> list[MappingRule]:
    fields_by_name: dict[str, dict[str, Any]] = {}
    for section in ("core_fields", "extended_fields"):
        for item in field_guide.get(section, []):
            field_name = item.get("field")
            if not isinstance(field_name, str):
                continue
            fields_by_name[field_name] = item

    rules: list[MappingRule] = []
    for field_name, value in data.items():
        if value is None:
            continue
        value_text = str(value).strip()
        if not value_text:
            continue

        guide_item = fields_by_name.get(field_name, {})
        docs_detected = int(guide_item.get("documents_detected", 0) or 0)
        if guide_item and docs_detected < min_docs:
            continue

        sample_location = guide_item.get("sample_location", {}) if guide_item else {}
        sample_label = sample_location.get("label")
        variants = list(guide_item.get("label_variants", [])) if guide_item else []

        primary_label = _first_non_empty([sample_label, *variants, field_name.replace("_", " ")])
        aliases = _unique_aliases([*variants, *(EXTRA_ALIASES.get(field_name, ()))], primary_label)
        value_type = VALUE_TYPE_BY_FIELD.get(field_name)
        confidence_threshold = _threshold_for_field(
            field_name=field_name,
            docs_detected=docs_detected,
        )

        rules.append(
            MappingRule(
                field_name=field_name,
                label=primary_label,
                row_offset=0,
                column_offset=0,
                sheet_name=None,
                required=False,
                aliases=tuple(aliases),
                value_type=value_type,
                target_strategy="offset_or_infer",
                confidence_threshold=confidence_threshold,
            )
        )

    return rules


def _discover_excel_files(input_dir: Path, output_dir: Path) -> list[Path]:
    results: list[Path] = []
    output_dir_resolved = output_dir.resolve()
    for path in sorted(input_dir.rglob("*")):
        if not path.is_file():
            continue
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        if path.name.startswith("~$"):
            continue
        try:
            path.resolve().relative_to(output_dir_resolved)
            continue
        except ValueError:
            pass
        results.append(path.resolve())
    return results


def _build_output_path(input_root: Path, template_path: Path, output_root: Path) -> Path:
    relative = template_path.relative_to(input_root)
    output_name = f"{template_path.stem}_autofilled{template_path.suffix}"
    return output_root / relative.parent / output_name


def _threshold_for_field(field_name: str, docs_detected: int) -> float:
    if field_name in {"numero_identificacion_nit", "representante_legal_nombre"}:
        return 0.88
    if docs_detected >= 120:
        return 0.82
    if docs_detected >= 50:
        return 0.86
    return 0.90


def _first_non_empty(values: list[Any]) -> str:
    for value in values:
        if not isinstance(value, str):
            continue
        cleaned = value.strip()
        if cleaned:
            return cleaned
    raise ValueError("No se encontro label valido para regla")


def _unique_aliases(candidates: list[str], primary_label: str) -> list[str]:
    seen_normalized: set[str] = {_norm(primary_label)}
    aliases: list[str] = []
    for candidate in candidates:
        if not isinstance(candidate, str):
            continue
        cleaned = candidate.strip()
        if not cleaned:
            continue
        normalized = _norm(cleaned)
        if not normalized or normalized in seen_normalized:
            continue
        seen_normalized.add(normalized)
        aliases.append(cleaned)
    return aliases


def _norm(value: str) -> str:
    return " ".join(value.lower().split())


def _decision_to_dict(item: PrecisionDecision) -> dict[str, Any]:
    return {
        "field_name": item.field_name,
        "status": item.status,
        "confidence": item.confidence,
        "reason": item.reason,
        "label": item.label,
        "label_position": _position_to_dict(item.label_position),
        "target_position": _position_to_dict(item.target_position),
        "value_preview": item.value_preview,
    }


def _position_to_dict(position: Any) -> dict[str, Any] | None:
    if position is None:
        return None
    return {
        "sheet_name": position.sheet_name,
        "row": position.row,
        "column": position.column,
    }


if __name__ == "__main__":
    raise SystemExit(main())
