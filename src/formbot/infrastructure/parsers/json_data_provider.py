from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from formbot.domain.exceptions import DataValidationError
from formbot.domain.ports.data_provider import DataProvider


class JsonFileDataProvider(DataProvider):
    def __init__(self, file_path: Path) -> None:
        self._file_path = file_path

    def load(self) -> dict[str, Any]:
        if not self._file_path.exists():
            raise DataValidationError(f"No existe el archivo JSON: {self._file_path}")

        try:
            with self._file_path.open("r", encoding="utf-8") as fp:
                payload = json.load(fp)
        except json.JSONDecodeError as exc:
            raise DataValidationError(
                f"JSON invalido en {self._file_path}: {exc}"
            ) from exc
        except OSError as exc:
            raise DataValidationError(
                f"No fue posible leer el JSON {self._file_path}: {exc}"
            ) from exc

        if not isinstance(payload, dict):
            raise DataValidationError(
                f"El contenido de {self._file_path} debe ser un objeto JSON"
            )

        return payload

