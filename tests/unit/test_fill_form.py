from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

fill_form_module = pytest.importorskip(
    "formbot.application.fill_form",
    reason="Caso de uso fill_form aun no implementado",
)
domain_models = pytest.importorskip("formbot.domain.models")
exceptions_module = pytest.importorskip("formbot.domain.exceptions")

FillFormUseCase = getattr(fill_form_module, "FillFormUseCase", None)
if FillFormUseCase is None:
    pytest.skip("No existe FillFormUseCase", allow_module_level=True)

MappingRule = domain_models.MappingRule
CellPosition = domain_models.CellPosition
LabelNotFoundError = exceptions_module.LabelNotFoundError
ValidationException = getattr(exceptions_module, "ValidationException", None)


def test_todos_los_campos_ok_retorna_resultado_sin_errores(
    mocker,
    sample_mapping: dict[str, dict[str, Any]],
    sample_data: dict[str, Any],
    tmp_path: Path,
) -> None:
    rules = _to_mapping_rules(sample_mapping)
    adapter = mocker.Mock()
    mapper = mocker.Mock()
    mapper.resolve_target.side_effect = lambda rule, label_position: CellPosition(
        sheet_name=label_position.sheet_name,
        row=label_position.row + rule.row_offset,
        column=label_position.column + rule.column_offset,
    )

    label_positions = {
        rule.label: CellPosition(sheet_name=rule.sheet_name or "Formulario", row=10, column=1)
        for rule in rules
    }
    adapter.find_label.side_effect = lambda text, sheet_name=None: label_positions[text]

    use_case = FillFormUseCase(document_adapter=adapter, field_mapper=mapper)
    output_path = tmp_path / "fill_form_ok.xlsx"
    result = use_case.execute(sample_data, rules, output_path)

    assert len(result.written_fields) == 3
    assert result.skipped_optional_fields == []
    assert adapter.write_value.call_count == 3
    adapter.save.assert_called_once_with(output_path)


def test_label_required_true_no_encontrado_detiene_y_no_genera_output(
    mocker,
    sample_mapping: dict[str, dict[str, Any]],
    sample_data: dict[str, Any],
    tmp_path: Path,
) -> None:
    rules = _to_mapping_rules(sample_mapping)
    required_rule = next(rule for rule in rules if rule.required)

    adapter = mocker.Mock()
    mapper = mocker.Mock()
    mapper.resolve_target.side_effect = lambda rule, label_position: label_position

    def _find_label(text: str, sheet_name: str | None = None):
        if text == required_rule.label:
            raise LabelNotFoundError("Label no encontrado")
        return CellPosition(sheet_name=sheet_name or "Formulario", row=1, column=1)

    adapter.find_label.side_effect = _find_label

    use_case = FillFormUseCase(document_adapter=adapter, field_mapper=mapper)
    output_path = tmp_path / "required_missing.xlsx"

    with pytest.raises(LabelNotFoundError):
        use_case.execute(sample_data, rules, output_path)

    adapter.save.assert_not_called()
    assert not output_path.exists()


def test_label_required_false_no_encontrado_skipped_y_continua(
    mocker,
    sample_mapping: dict[str, dict[str, Any]],
    sample_data: dict[str, Any],
    tmp_path: Path,
) -> None:
    rules = _to_mapping_rules(sample_mapping)
    optional_rule = next(rule for rule in rules if not rule.required)

    adapter = mocker.Mock()
    mapper = mocker.Mock()
    mapper.resolve_target.side_effect = lambda rule, label_position: label_position

    def _find_label(text: str, sheet_name: str | None = None):
        if text == optional_rule.label:
            raise LabelNotFoundError("Label opcional ausente")
        return CellPosition(sheet_name=sheet_name or "Formulario", row=1, column=1)

    adapter.find_label.side_effect = _find_label

    use_case = FillFormUseCase(document_adapter=adapter, field_mapper=mapper)
    output_path = tmp_path / "optional_pending.xlsx"
    result = use_case.execute(sample_data, rules, output_path)

    write_results = getattr(result, "write_results")
    assert any(
        item.field_name == optional_rule.field_name and item.status == "skipped"
        for item in write_results
    )
    adapter.save.assert_called_once_with(output_path)


def test_dropdown_invalido_se_registra_como_error_y_pipeline_continua(
    mocker,
    sample_mapping: dict[str, dict[str, Any]],
    sample_data: dict[str, Any],
    tmp_path: Path,
) -> None:
    if ValidationException is None:
        pytest.xfail("ValidationException aun no implementada")

    rules = _to_mapping_rules(sample_mapping)
    dropdown_rule = next(rule for rule in rules if rule.field_name == "tipo_persona")
    invalid_data = dict(sample_data)
    invalid_data[dropdown_rule.field_name] = "VALOR_INVALIDO"

    adapter = mocker.Mock()
    mapper = mocker.Mock()
    mapper.resolve_target.side_effect = lambda rule, label_position: label_position
    adapter.find_label.return_value = CellPosition(sheet_name="Formulario", row=5, column=5)

    def _write_value(position, value):
        if value == "VALOR_INVALIDO":
            raise ValidationException("Dropdown invalido")

    adapter.write_value.side_effect = _write_value

    use_case = FillFormUseCase(document_adapter=adapter, field_mapper=mapper)
    output_path = tmp_path / "dropdown_invalid.xlsx"

    result = use_case.execute(invalid_data, rules, output_path)

    write_results = getattr(result, "write_results")
    assert any(
        item.field_name == dropdown_rule.field_name and item.status == "error"
        for item in write_results
    )
    adapter.save.assert_called_once_with(output_path)


def test_write_mode_mark_escribe_simbolo_configurado(mocker, tmp_path: Path) -> None:
    adapter = mocker.Mock()
    mapper = mocker.Mock()
    mapper.resolve_target.return_value = CellPosition(
        sheet_name="Formulario",
        row=10,
        column=3,
    )
    adapter.find_label.return_value = CellPosition(
        sheet_name="Formulario",
        row=10,
        column=1,
    )

    rule = MappingRule(
        field_name="op_internacionales",
        label="¿Realiza operaciones internacionales?",
        row_offset=0,
        column_offset=0,
        sheet_name="Formulario",
        required=True,
        write_mode="mark",
        mark_symbol="✔",
    )
    data = {"op_internacionales": True}

    use_case = FillFormUseCase(document_adapter=adapter, field_mapper=mapper)
    output_path = tmp_path / "mark_mode.xlsx"
    result = use_case.execute(data, [rule], output_path)

    adapter.write_value.assert_called_once_with(
        CellPosition(sheet_name="Formulario", row=10, column=3),
        "✔",
    )
    assert result.written_fields[0].value == "✔"


def _to_mapping_rules(sample_mapping: dict[str, dict[str, Any]]) -> list:
    rules = []
    for field_name, definition in sample_mapping.items():
        offset = definition.get("offset", {})
        rules.append(
            MappingRule(
                field_name=field_name,
                label=definition["label"],
                row_offset=offset.get("row", 0),
                column_offset=offset.get("col", 0),
                sheet_name=definition.get("sheet"),
                required=definition.get("required", True),
            )
        )
    return rules
