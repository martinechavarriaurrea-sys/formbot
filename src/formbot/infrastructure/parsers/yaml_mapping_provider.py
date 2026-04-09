from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

import yaml

from formbot.domain.exceptions import MappingRuleError
from formbot.domain.models import MappingRule
from formbot.domain.ports.mapping_provider import MappingProvider

LOGGER = logging.getLogger(__name__)

# Patrón de encabezado de sección (ej. "4. REFERENCIAS COMERCIALES")
_SECTION_HEADER_RE = re.compile(r"^\d+\s*\.")

# Campos de firma que nunca deben bloquear el fill
_FIRMA_FIELDS = frozenset({
    "firma_representante_nombre",
    "firma_representante_documento",
    "firma_diligencio_nombre",
    "firma_diligencio_documento",
})

# Campos web/opcionales que raramente existen en todos los forms
_OPTIONAL_FIELDS = frozenset({"pagina_web", "correo_gerencia", "correo_notificaciones"})

# Offset de columna a partir del cual se emite aviso
_COL_OFFSET_WARN = 6


class YamlMappingProvider(MappingProvider):
    def __init__(self, file_path: Path) -> None:
        self._file_path = file_path

    def load(self) -> list[MappingRule]:
        if not self._file_path.exists():
            raise MappingRuleError(f"No existe el archivo de mapping: {self._file_path}")

        try:
            with self._file_path.open("r", encoding="utf-8") as fp:
                raw_mapping = yaml.safe_load(fp) or {}
        except yaml.YAMLError as exc:
            raise MappingRuleError(
                f"YAML invalido en {self._file_path}: {exc}"
            ) from exc
        except OSError as exc:
            raise MappingRuleError(
                f"No fue posible leer el mapping {self._file_path}: {exc}"
            ) from exc

        if not isinstance(raw_mapping, dict):
            raise MappingRuleError(
                "El archivo YAML de mapping debe ser un diccionario "
                "con llave=field_name y valor=configuracion de regla"
            )

        mapping_rules: list[MappingRule] = []
        for raw_field_name, raw_rule in raw_mapping.items():
            if not isinstance(raw_field_name, str) or not raw_field_name.strip():
                raise MappingRuleError(
                    "Cada llave del YAML debe ser field_name textual no vacio"
                )
            field_name = raw_field_name.strip()
            mapping_rules.append(self._parse_rule(field_name, raw_rule))
        return mapping_rules

    @staticmethod
    def _parse_rule(field_name: str, raw_rule: Any) -> MappingRule:
        if not isinstance(raw_rule, dict):
            raise MappingRuleError(
                f"La regla '{field_name}' debe ser un objeto con label/offset"
            )

        label = raw_rule.get("label")
        if not isinstance(label, str) or not label.strip():
            raise MappingRuleError(f"La regla '{field_name}' requiere 'label' textual")

        offset = raw_rule.get("offset", {})
        if not isinstance(offset, dict):
            raise MappingRuleError(f"La regla '{field_name}' tiene 'offset' invalido")

        row_offset = _as_int(offset.get("row", 0), field_name, "row")
        col_offset = _as_int(offset.get("col", 0), field_name, "col")

        sheet_name = raw_rule.get("sheet")
        if sheet_name is not None and not isinstance(sheet_name, str):
            raise MappingRuleError(
                f"La regla '{field_name}' tiene 'sheet' invalido; debe ser texto"
            )

        required = raw_rule.get("required", True)
        if not isinstance(required, bool):
            raise MappingRuleError(
                f"La regla '{field_name}' tiene 'required' invalido; debe ser booleano"
            )

        aliases = raw_rule.get("aliases", [])
        if aliases is None:
            aliases = []
        if not isinstance(aliases, list):
            raise MappingRuleError(
                f"La regla '{field_name}' tiene 'aliases' invalido; debe ser lista de textos"
            )
        for alias in aliases:
            if not isinstance(alias, str):
                raise MappingRuleError(
                    f"La regla '{field_name}' tiene alias invalido: {alias}"
                )

        value_type = raw_rule.get("type")
        if value_type is not None and not isinstance(value_type, str):
            raise MappingRuleError(
                f"La regla '{field_name}' tiene 'type' invalido; debe ser texto"
            )

        target_strategy = raw_rule.get("target_strategy", "offset_or_infer")
        if not isinstance(target_strategy, str):
            raise MappingRuleError(
                f"La regla '{field_name}' tiene 'target_strategy' invalido; debe ser texto"
            )

        confidence_threshold = raw_rule.get("confidence_threshold")
        if confidence_threshold is not None and not isinstance(
            confidence_threshold, (int, float)
        ):
            raise MappingRuleError(
                f"La regla '{field_name}' tiene 'confidence_threshold' invalido; debe ser numerico"
            )
        if confidence_threshold is not None and not (0 <= confidence_threshold <= 1):
            raise MappingRuleError(
                f"La regla '{field_name}' tiene 'confidence_threshold' fuera de rango 0..1"
            )

        write_mode = raw_rule.get("write_mode", "value")
        if not isinstance(write_mode, str):
            raise MappingRuleError(
                f"La regla '{field_name}' tiene 'write_mode' invalido; debe ser texto"
            )

        mark_symbol = raw_rule.get("mark_symbol", "X")
        if not isinstance(mark_symbol, str) or not mark_symbol.strip():
            raise MappingRuleError(
                f"La regla '{field_name}' tiene 'mark_symbol' invalido; debe ser texto no vacio"
            )

        # ── Validaciones preventivas ──────────────────────────────────────────
        # 1. Campos de firma marcados como required=True bloquean el pipeline
        if required and field_name in _FIRMA_FIELDS:
            LOGGER.warning(
                "MAPPING [%s]: 'required: true' en campo de firma '%s' — "
                "si el label no existe en el form el fill falla. Usa 'required: false'.",
                field_name, field_name,
            )

        # 2. Campos opcionales por naturaleza marcados como required=True
        if required and field_name in _OPTIONAL_FIELDS:
            LOGGER.warning(
                "MAPPING [%s]: 'required: true' en campo opcional '%s' — "
                "no todos los formularios incluyen este campo. Usa 'required: false'.",
                field_name, field_name,
            )

        # 3. Campos mark (boolean) con required=True raramente tienen sentido
        if required and write_mode.strip().lower() == "mark":
            LOGGER.warning(
                "MAPPING [%s]: 'required: true' combinado con 'write_mode: mark' — "
                "los campos de marcación suelen ser opcionales.",
                field_name,
            )

        # 4. col_offset inusualmente alto (>= 6): posible error de diseño
        if col_offset >= _COL_OFFSET_WARN:
            LOGGER.warning(
                "MAPPING [%s]: col_offset=%d es inusualmente alto. "
                "Verifica que la celda destino esté realmente %d columnas a la derecha del label '%s'.",
                field_name, col_offset, col_offset, label,
            )

        # 5. col_offset=0 con row_offset>1 sobre un encabezado de sección:
        #    escribe en la misma columna que el header, no en la columna del valor.
        if col_offset == 0 and row_offset > 1 and _SECTION_HEADER_RE.match(label.strip()):
            LOGGER.warning(
                "MAPPING [%s]: col_offset=0 sobre encabezado de sección '%s' (row=%d). "
                "El valor se escribe en la MISMA columna que el header. "
                "Probablemente necesites col_offset=1.",
                field_name, label, row_offset,
            )

        # 6. Nombre del campo no coincide con la sección semántica del label
        #    Detecta el caso correo_electronico anclado en sección de representante legal.
        _RL_LABELS = {"datos representante legal", "representante legal", "datos del representante"}
        _label_norm = label.strip().lower()
        if field_name == "correo_electronico" and any(t in _label_norm for t in _RL_LABELS):
            LOGGER.warning(
                "MAPPING [%s]: el campo 'correo_electronico' (email general de la empresa) "
                "está anclado en la sección '%s'. "
                "Si apunta al correo del RL, usa 'representante_legal_correo' como field_name.",
                field_name, label,
            )

        return MappingRule(
            field_name=field_name,
            label=label,
            row_offset=row_offset,
            column_offset=col_offset,
            sheet_name=sheet_name,
            required=required,
            aliases=tuple(aliases),
            value_type=value_type.strip().lower() if value_type is not None else None,
            target_strategy=target_strategy.strip().lower(),
            confidence_threshold=float(confidence_threshold)
            if confidence_threshold is not None
            else None,
            write_mode=write_mode.strip().lower(),
            mark_symbol=mark_symbol,
        )


def _as_int(value: Any, field_name: str, key: str) -> int:
    if isinstance(value, bool):
        raise MappingRuleError(
            f"La regla '{field_name}' tiene offset '{key}' invalido: {value}"
        )
    if isinstance(value, int):
        return value
    raise MappingRuleError(
        f"La regla '{field_name}' tiene offset '{key}' invalido: {value}"
    )
