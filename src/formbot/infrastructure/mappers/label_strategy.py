from __future__ import annotations

from collections.abc import Sequence

from formbot.shared.utils import normalize_text


class ExactLabelStrategy:
    """Busca una etiqueta por coincidencia exacta (case-insensitive)."""

    def find(
        self,
        grid: Sequence[Sequence[object]],
        label: str,
    ) -> tuple[int, int] | None:
        target = normalize_text(label)
        if not target:
            return None

        for row_index, row in enumerate(grid, start=1):
            for column_index, value in enumerate(row, start=1):
                if not isinstance(value, str):
                    continue
                if normalize_text(value) == target:
                    return (row_index, column_index)
        return None
