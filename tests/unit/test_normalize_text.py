from __future__ import annotations

import pytest

utils_module = pytest.importorskip("formbot.shared.utils", reason="utils no disponible")

normalize_text = utils_module.normalize_text


def test_normalize_text_quita_tildes_y_espacios() -> None:
    assert normalize_text("  Correo   electr\u00f3nico  ") == "correo electronico"


def test_normalize_text_unifica_separadores() -> None:
    assert normalize_text("E-mail___Representante.Legal") == "e mail representante legal"
