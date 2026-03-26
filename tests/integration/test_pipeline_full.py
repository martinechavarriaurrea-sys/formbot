from __future__ import annotations

import json
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl", reason="openpyxl no disponible")
yaml = pytest.importorskip("yaml", reason="PyYAML no disponible")

bootstrap_module = pytest.importorskip(
    "formbot.app.bootstrap",
    reason="Bootstrap de pipeline aun no disponible",
)
exceptions_module = pytest.importorskip("formbot.domain.exceptions")

bootstrap_excel_pipeline = bootstrap_module.bootstrap_excel_pipeline
DataValidationError = exceptions_module.DataValidationError


@pytest.fixture
def pipeline_mapping_path(fixtures_dir: Path) -> Path:
    path = fixtures_dir / "pipeline_mapping.yaml"
    if not path.exists():
        pytest.skip("Falta tests/fixtures/pipeline_mapping.yaml")
    return path


@pytest.fixture
def pipeline_data_valid_path(fixtures_dir: Path) -> Path:
    path = fixtures_dir / "pipeline_data_valid.json"
    if not path.exists():
        pytest.skip("Falta tests/fixtures/pipeline_data_valid.json")
    return path


@pytest.fixture
def pipeline_data_missing_required_path(fixtures_dir: Path) -> Path:
    path = fixtures_dir / "pipeline_data_missing_required.json"
    if not path.exists():
        pytest.skip("Falta tests/fixtures/pipeline_data_missing_required.json")
    return path


def test_excel_real_yaml_datos_validos_genera_documento_correcto(
    excel_path: Path,
    pipeline_mapping_path: Path,
    pipeline_data_valid_path: Path,
    tmp_path: Path,
) -> None:
    output_path = tmp_path / "pipeline_valid_output.xlsx"
    context = bootstrap_excel_pipeline(
        template_path=excel_path,
        mapping_path=pipeline_mapping_path,
        data_path=pipeline_data_valid_path,
    )
    try:
        result = context.use_case.execute(
            data=context.data,
            mapping_rules=context.mapping_rules,
            output_path=output_path,
        )
    finally:
        context.use_case.close()

    assert output_path.exists()
    assert result.output_path == output_path
    _assert_values_written_by_mapping(output_path, pipeline_mapping_path, pipeline_data_valid_path)


def test_required_faltante_falla_limpiamente_sin_output_corrupto(
    excel_path: Path,
    pipeline_mapping_path: Path,
    pipeline_data_missing_required_path: Path,
    tmp_path: Path,
) -> None:
    output_path = tmp_path / "pipeline_missing_required_output.xlsx"
    context = bootstrap_excel_pipeline(
        template_path=excel_path,
        mapping_path=pipeline_mapping_path,
        data_path=pipeline_data_missing_required_path,
    )
    try:
        with pytest.raises(DataValidationError):
            context.use_case.execute(
                data=context.data,
                mapping_rules=context.mapping_rules,
                output_path=output_path,
            )
    finally:
        context.use_case.close()

    assert not output_path.exists()


def test_documento_generado_mantiene_misma_estructura_y_estilos(
    excel_path: Path,
    pipeline_mapping_path: Path,
    pipeline_data_valid_path: Path,
    tmp_path: Path,
) -> None:
    output_path = tmp_path / "pipeline_structure_output.xlsx"
    context = bootstrap_excel_pipeline(
        template_path=excel_path,
        mapping_path=pipeline_mapping_path,
        data_path=pipeline_data_valid_path,
    )
    try:
        context.use_case.execute(
            data=context.data,
            mapping_rules=context.mapping_rules,
            output_path=output_path,
        )
    finally:
        context.use_case.close()

    assert output_path.exists()
    _assert_same_structure_and_styles(excel_path, output_path)


def _assert_values_written_by_mapping(
    output_path: Path,
    mapping_path: Path,
    data_path: Path,
) -> None:
    mapping = yaml.safe_load(mapping_path.read_text(encoding="utf-8")) or {}
    data = json.loads(data_path.read_text(encoding="utf-8"))

    wb = openpyxl.load_workbook(output_path)
    try:
        for field_name, rule in mapping.items():
            if field_name not in data:
                continue
            label = rule["label"]
            sheet_name = rule.get("sheet")
            row_offset = rule.get("offset", {}).get("row", 0)
            col_offset = rule.get("offset", {}).get("col", 0)
            label_position = _find_label_position(wb, label, sheet_name)
            assert label_position is not None, f"No se encontro label '{label}'"

            target_sheet = wb[label_position[0]]
            target_cell = target_sheet.cell(
                row=label_position[1] + row_offset,
                column=label_position[2] + col_offset,
            )
            assert target_cell.value == data[field_name]
    finally:
        wb.close()


def _find_label_position(workbook, label: str, sheet_name: str | None):
    target = " ".join(label.strip().lower().split())
    sheets = [workbook[sheet_name]] if sheet_name else workbook.worksheets
    for sheet in sheets:
        for row in sheet.iter_rows():
            for cell in row:
                if not isinstance(cell.value, str):
                    continue
                current = " ".join(cell.value.strip().lower().split())
                if current == target:
                    return (sheet.title, cell.row, cell.column)
    return None


def _assert_same_structure_and_styles(original_path: Path, generated_path: Path) -> None:
    original = openpyxl.load_workbook(original_path)
    generated = openpyxl.load_workbook(generated_path)
    try:
        assert original.sheetnames == generated.sheetnames
        for sheet_name in original.sheetnames:
            original_sheet = original[sheet_name]
            generated_sheet = generated[sheet_name]

            assert original_sheet.max_row == generated_sheet.max_row
            assert original_sheet.max_column == generated_sheet.max_column
            assert sorted(map(str, original_sheet.merged_cells.ranges)) == sorted(
                map(str, generated_sheet.merged_cells.ranges)
            )

            for row in range(1, original_sheet.max_row + 1):
                for col in range(1, original_sheet.max_column + 1):
                    original_cell = original_sheet.cell(row=row, column=col)
                    generated_cell = generated_sheet.cell(row=row, column=col)
                    assert original_cell.style_id == generated_cell.style_id
                    assert original_cell.number_format == generated_cell.number_format
                    assert _alignment_signature(original_cell.alignment) == _alignment_signature(
                        generated_cell.alignment
                    )

            original_column_keys = set(original_sheet.column_dimensions.keys())
            generated_column_keys = set(generated_sheet.column_dimensions.keys())
            assert original_column_keys == generated_column_keys
            for key in original_column_keys:
                assert (
                    original_sheet.column_dimensions[key].width
                    == generated_sheet.column_dimensions[key].width
                )
    finally:
        original.close()
        generated.close()


def _alignment_signature(alignment) -> tuple:
    return (
        alignment.horizontal,
        alignment.vertical,
        alignment.textRotation,
        alignment.wrapText,
        alignment.shrinkToFit,
        alignment.indent,
        alignment.relativeIndent,
        alignment.justifyLastLine,
        alignment.readingOrder,
    )
