from __future__ import annotations

from formbot.domain.exceptions import PositionOutOfBoundsError
from formbot.domain.models import CellPosition, MappingRule
from formbot.domain.ports.field_mapper import FieldMapper


class LabelOffsetMapper(FieldMapper):
    def resolve_target(self, rule: MappingRule, label_position: CellPosition) -> CellPosition:
        target = label_position.shifted(rule.row_offset, rule.column_offset)
        if target.row < 1 or target.column < 1:
            raise PositionOutOfBoundsError(
                f"Offset invalido para '{rule.field_name}': "
                f"fila={target.row}, columna={target.column}"
            )
        return target

