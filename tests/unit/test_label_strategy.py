from __future__ import annotations

import pytest

label_strategy_module = pytest.importorskip(
    "formbot.infrastructure.mappers.label_strategy",
    reason="Modulo de estrategia exacta aun no implementado",
)

ExactLabelStrategy = getattr(label_strategy_module, "ExactLabelStrategy", None)
if ExactLabelStrategy is None:
    pytest.skip("ExactLabelStrategy no existe aun", allow_module_level=True)


def test_exact_match_retorna_posicion_correcta() -> None:
    strategy = ExactLabelStrategy()
    grid = [
        ["ID", "Nombre del representante legal", None],
        [None, None, None],
    ]

    result = strategy.find(grid, "Nombre del representante legal")

    assert result == (1, 2)


def test_case_insensitive_funciona() -> None:
    strategy = ExactLabelStrategy()
    grid = [["Nombre Del Representante Legal"]]

    result = strategy.find(grid, "nombre del representante legal")

    assert result == (1, 1)


def test_label_inexistente_retorna_none() -> None:
    strategy = ExactLabelStrategy()
    grid = [["NIT", "Razon social"]]

    result = strategy.find(grid, "Nombre del representante legal")

    assert result is None


def test_hoja_vacia_retorna_none() -> None:
    strategy = ExactLabelStrategy()
    grid = []

    result = strategy.find(grid, "Cualquier label")

    assert result is None


def test_texto_parcial_no_hace_match_en_exact_strategy() -> None:
    strategy = ExactLabelStrategy()
    grid = [["Nombre del representante legal"]]

    result = strategy.find(grid, "Nombre del representante")

    assert result is None

