from __future__ import annotations

from abc import ABC, abstractmethod

from formbot.domain.models import CellPosition, MappingRule


class FieldMapper(ABC):
    @abstractmethod
    def resolve_target(self, rule: MappingRule, label_position: CellPosition) -> CellPosition:
        raise NotImplementedError

