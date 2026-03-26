from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any


class DataProvider(ABC):
    @abstractmethod
    def load(self) -> dict[str, Any]:
        raise NotImplementedError

