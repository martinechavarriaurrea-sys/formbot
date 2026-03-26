from __future__ import annotations

from pathlib import Path

import pytest

fill_form_module = pytest.importorskip(
    "formbot.application.fill_form",
    reason="Caso de uso fill_form no disponible",
)
domain_models = pytest.importorskip("formbot.domain.models")
exceptions_module = pytest.importorskip("formbot.domain.exceptions")

FillFormUseCase = fill_form_module.FillFormUseCase
MappingRule = domain_models.MappingRule
CellPosition = domain_models.CellPosition
LabelNotFoundError = exceptions_module.LabelNotFoundError


def test_fill_form_usa_alias_cuando_label_principal_no_existe(
    mocker,
    tmp_path: Path,
) -> None:
    rule = MappingRule(
        field_name="correo",
        label="Correo electronico",
        row_offset=0,
        column_offset=3,
        sheet_name="Formulario",
        required=True,
        aliases=("Email", "E-mail"),
    )

    adapter = mocker.Mock()
    mapper = mocker.Mock()
    mapper.resolve_target.return_value = CellPosition(
        sheet_name="Formulario",
        row=6,
        column=5,
    )

    attempts: list[str] = []

    def _find_label(text: str, sheet_name: str | None = None) -> CellPosition:
        attempts.append(text)
        if text == "E-mail":
            return CellPosition(sheet_name=sheet_name or "Formulario", row=6, column=2)
        raise LabelNotFoundError(f"No existe label: {text}")

    adapter.find_label.side_effect = _find_label

    use_case = FillFormUseCase(document_adapter=adapter, field_mapper=mapper)
    output_path = tmp_path / "alias_fallback.xlsx"
    result = use_case.execute({"correo": "contacto@empresa.com"}, [rule], output_path)

    assert attempts == ["Correo electronico", "Email", "E-mail"]
    assert len(result.written_fields) == 1
    adapter.write_value.assert_called_once()
    adapter.save.assert_called_once_with(output_path)
