from __future__ import annotations

from abc import ABC, abstractmethod

from formbot.domain.models import MappingRule


class MappingProvider(ABC):
    @abstractmethod
    def load(self) -> list[MappingRule]:
        raise NotImplementedError

