from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any

from formbot.domain.models import CellPosition


class DocumentAdapter(ABC):
    @abstractmethod
    def find_label(self, text: str, sheet_name: str | None = None) -> CellPosition:
        raise NotImplementedError

    @abstractmethod
    def write_value(self, position: CellPosition, value: Any) -> None:
        raise NotImplementedError

    @abstractmethod
    def save(self, output_path: Path) -> None:
        raise NotImplementedError

    @abstractmethod
    def close(self) -> None:
        raise NotImplementedError

    def find_adjacent_empty(self, position: CellPosition) -> CellPosition | None:
        """Primera celda vacía adyacente (derecha o abajo). None si no aplica."""
        return None

