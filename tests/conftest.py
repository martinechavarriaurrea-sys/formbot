from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import Any

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))


@pytest.fixture(scope="session")
def fixtures_dir() -> Path:
    return Path(__file__).resolve().parent / "fixtures"


@pytest.fixture(scope="session")
def excel_path(fixtures_dir: Path) -> Path:
    configured_name = os.getenv("FORMBOT_TEST_EXCEL", "").strip()
    if configured_name:
        explicit_path = fixtures_dir / configured_name
        if explicit_path.exists():
            return explicit_path
        pytest.skip(
            "FORMBOT_TEST_EXCEL fue definido pero el archivo no existe en tests/fixtures/"
        )

    candidates = sorted(fixtures_dir.glob("*.xlsx")) + sorted(fixtures_dir.glob("*.xlsm"))
    if not candidates:
        pytest.skip(
            "No hay Excel real en tests/fixtures/. Agrega el archivo para habilitar pruebas de adapter/integration."
        )
    return candidates[0]


@pytest.fixture
def sample_mapping() -> dict[str, dict[str, Any]]:
    return {
        "nombre_representante": {
            "label": "Nombre del representante legal",
            "offset": {"row": 0, "col": 3},
            "required": True,
            "type": "text",
            "sheet": "Formulario",
        },
        "fecha_firma": {
            "label": "Fecha de firma",
            "offset": {"row": 0, "col": 3},
            "required": False,
            "type": "date",
            "sheet": "Formulario",
        },
        "tipo_persona": {
            "label": "Tipo de persona",
            "offset": {"row": 0, "col": 3},
            "required": True,
            "type": "dropdown",
            "options": ["Natural", "Juridica"],
            "sheet": "Formulario",
        },
    }


@pytest.fixture
def sample_data() -> dict[str, Any]:
    return {
        "nombre_representante": "Juan Perez",
        "fecha_firma": "2026-03-20",
        "tipo_persona": "Natural",
    }


@pytest.fixture
def sample_data_missing_required() -> dict[str, Any]:
    return {
        "fecha_firma": "2026-03-20",
        "tipo_persona": "Natural",
    }

