from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Mapping, Sequence

from formbot.domain.exceptions import (
    DataValidationError,
    LabelNotFoundError,
    MappingRuleError,
    ValidationException,
)
from formbot.domain.models import (
    CellPosition,
    FieldWriteResult,
    FieldWriteTrace,
    FillFormResult,
    MappingRule,
)
from formbot.domain.ports.document_adapter import DocumentAdapter
from formbot.domain.ports.field_mapper import FieldMapper
from formbot.shared.utils import find_duplicates

LOGGER = logging.getLogger(__name__)


class FillFormUseCase:
    def __init__(self, document_adapter: DocumentAdapter, field_mapper: FieldMapper) -> None:
        self._document_adapter = document_adapter
        self._field_mapper = field_mapper

    def execute(
        self,
        data: Mapping[str, Any],
        mapping_rules: Sequence[MappingRule],
        output_path: Path,
    ) -> FillFormResult:
        self._validate_rules(mapping_rules)
        self._validate_required_data(data, mapping_rules)

        written_fields: list[FieldWriteTrace] = []
        write_results: list[FieldWriteResult] = []
        skipped_optional_fields: list[str] = []
        mapped_fields = {rule.field_name for rule in mapping_rules}
        unmapped_input_fields = sorted(set(data.keys()) - mapped_fields)

        if unmapped_input_fields:
            LOGGER.warning(
                "Se recibieron campos sin regla de mapeo: %s",
                ", ".join(unmapped_input_fields),
            )

        for rule in mapping_rules:
            if rule.field_name not in data:
                skipped_optional_fields.append(rule.field_name)
                write_results.append(
                    FieldWriteResult(
                        field_name=rule.field_name,
                        status="skipped",
                        detail="Campo opcional ausente en payload",
                    )
                )
                LOGGER.info(
                    "Se omite campo opcional '%s' por no venir en el payload",
                    rule.field_name,
                )
                continue

            try:
                label_position = self._find_label_with_aliases(rule)
            except LabelNotFoundError as exc:
                if rule.required:
                    raise
                write_results.append(
                    FieldWriteResult(
                        field_name=rule.field_name,
                        status="skipped",
                        detail=str(exc),
                    )
                )
                LOGGER.warning(
                    "No se encontro label opcional para campo '%s': %s",
                    rule.field_name,
                    exc,
                )
                continue

            target_position = self._field_mapper.resolve_target(rule, label_position)
            # Si la estrategia incluye inferencia y el offset es cero (target == label),
            # buscar la primera celda adyacente vacía para no sobrescribir el label.
            if (
                rule.target_strategy in {"infer", "offset_or_infer"}
                and target_position == label_position
            ):
                inferred = self._document_adapter.find_adjacent_empty(label_position)
                if inferred is not None:
                    target_position = inferred
                elif rule.write_mode == "mark":
                    # Sin celda adyacente libre: omitir antes de corromper la etiqueta
                    LOGGER.warning(
                        "No se encontró celda vacía adyacente para campo mark '%s'. Se omite.",
                        rule.field_name,
                    )
                    write_results.append(
                        FieldWriteResult(
                            field_name=rule.field_name,
                            status="skipped",
                            detail="No hay celda adyacente libre para marcar",
                        )
                    )
                    continue
            raw_value = data[rule.field_name]
            if rule.write_mode == "mark":
                # Solo marcar si el valor es verdadero; si es False/falsy, omitir campo
                if not raw_value:
                    write_results.append(
                        FieldWriteResult(
                            field_name=rule.field_name,
                            status="skipped",
                            detail="Valor falso: no se marca",
                        )
                    )
                    continue
                value_to_write = rule.mark_symbol
            else:
                value_to_write = raw_value
            try:
                self._document_adapter.write_value(target_position, value_to_write)
            except ValidationException as exc:
                write_results.append(
                    FieldWriteResult(
                        field_name=rule.field_name,
                        status="error",
                        detail=str(exc),
                    )
                )
                LOGGER.warning(
                    "Validacion de dropdown fallida para campo '%s': %s",
                    rule.field_name,
                    exc,
                )
                continue

            written_fields.append(
                FieldWriteTrace(
                    field_name=rule.field_name,
                    label_position=label_position,
                    target_position=target_position,
                    value=value_to_write,
                )
            )
            write_results.append(
                FieldWriteResult(
                    field_name=rule.field_name,
                    status="written",
                    detail=None,
                )
            )

        self._document_adapter.save(output_path)
        return FillFormResult(
            output_path=output_path,
            written_fields=written_fields,
            write_results=write_results,
            skipped_optional_fields=skipped_optional_fields,
            unmapped_input_fields=unmapped_input_fields,
        )

    def close(self) -> None:
        self._document_adapter.close()

    @staticmethod
    def _validate_rules(mapping_rules: Sequence[MappingRule]) -> None:
        if not mapping_rules:
            raise MappingRuleError("No hay reglas de mapeo configuradas")

        duplicated_fields = find_duplicates([rule.field_name for rule in mapping_rules])
        if duplicated_fields:
            raise MappingRuleError(
                "Existen campos duplicados en mapping: "
                + ", ".join(sorted(duplicated_fields))
            )

    @staticmethod
    def _validate_required_data(
        data: Mapping[str, Any],
        mapping_rules: Sequence[MappingRule],
    ) -> None:
        missing_required_fields = sorted(
            rule.field_name
            for rule in mapping_rules
            if rule.required and rule.field_name not in data
        )
        if missing_required_fields:
            raise DataValidationError(
                "Faltan campos requeridos en el payload: "
                + ", ".join(missing_required_fields)
            )

    def _find_label_with_aliases(self, rule: MappingRule) -> CellPosition:
        search_terms = [rule.label, *rule.aliases]
        for term in search_terms:
            try:
                return self._document_adapter.find_label(
                    text=term,
                    sheet_name=rule.sheet_name,
                )
            except LabelNotFoundError:
                continue

        raise LabelNotFoundError(
            f"No se encontro etiqueta para campo '{rule.field_name}' "
            f"(label='{rule.label}', aliases={list(rule.aliases)})"
        )
