from __future__ import annotations

import hashlib
import shutil
from datetime import date
from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl", reason="openpyxl no disponible")
load_workbook = openpyxl.load_workbook
range_boundaries = openpyxl.utils.cell.range_boundaries

adapter_module = pytest.importorskip(
    "formbot.infrastructure.document_readers.excel_document_adapter",
    reason="Excel adapter aun no implementado",
)
domain_models = pytest.importorskip("formbot.domain.models")
exceptions_module = pytest.importorskip("formbot.domain.exceptions")

ExcelAdapter = getattr(adapter_module, "ExcelAdapter", None) or getattr(
    adapter_module, "ExcelDocumentAdapter", None
)
if ExcelAdapter is None:
    pytest.skip("No existe clase ExcelAdapter/ExcelDocumentAdapter", allow_module_level=True)

CellPosition = domain_models.CellPosition
ValidationException = getattr(exceptions_module, "ValidationException", None)


def test_carga_excel_real_sin_modificarlo(excel_path: Path) -> None:
    original_hash = _sha256(excel_path)

    adapter = ExcelAdapter(excel_path)
    adapter.close()

    assert _sha256(excel_path) == original_hash


def test_escribe_texto_numero_fecha_y_dropdown_valido(excel_path: Path, tmp_path: Path) -> None:
    working_excel = _copy_to_tmp(excel_path, tmp_path)
    adapter = ExcelAdapter(working_excel)
    sheet_name = _first_sheet_name(working_excel)

    text_pos = CellPosition(sheet_name=sheet_name, row=1, column=50)
    number_pos = CellPosition(sheet_name=sheet_name, row=2, column=50)
    date_pos = CellPosition(sheet_name=sheet_name, row=3, column=50)

    adapter.write_value(text_pos, "Texto QA")
    adapter.write_value(number_pos, 123456)
    adapter.write_value(date_pos, date(2026, 3, 20))

    dropdown_target = _find_dropdown_target(working_excel)
    if dropdown_target is not None:
        adapter.write_value(dropdown_target.position, dropdown_target.valid_option)

    output_path = tmp_path / "excel_adapter_write_ok.xlsx"
    adapter.save(output_path)
    adapter.close()

    workbook = load_workbook(output_path)
    sheet = workbook[sheet_name]
    assert sheet.cell(text_pos.row, text_pos.column).value == "Texto QA"
    assert sheet.cell(number_pos.row, number_pos.column).value == 123456
    assert str(sheet.cell(date_pos.row, date_pos.column).value).startswith("2026-03-20")
    if dropdown_target is not None:
        assert (
            sheet.cell(dropdown_target.position.row, dropdown_target.position.column).value
            == dropdown_target.valid_option
        )
    workbook.close()


def test_dropdown_invalido_lanza_validation_exception(excel_path: Path, tmp_path: Path) -> None:
    if ValidationException is None:
        pytest.xfail("ValidationException aun no existe en domain.exceptions")

    working_excel = _copy_to_tmp(excel_path, tmp_path)
    dropdown_target = _find_dropdown_target(working_excel)
    if dropdown_target is None:
        pytest.skip("El Excel fixture no contiene dropdown detectable para este test")

    adapter = ExcelAdapter(working_excel)
    with pytest.raises(ValidationException):
        adapter.write_value(dropdown_target.position, "__INVALID_DROPDOWN_VALUE__")
    adapter.close()


def test_celda_merged_escribe_en_celda_origen(excel_path: Path, tmp_path: Path) -> None:
    working_excel = _copy_to_tmp(excel_path, tmp_path)
    merged_origin = _find_first_merged_origin(working_excel)
    if merged_origin is None:
        pytest.skip("El Excel fixture no tiene celdas merged")

    adapter = ExcelAdapter(working_excel)
    adapter.write_value(merged_origin, "MERGED_ORIGIN_OK")
    output_path = tmp_path / "excel_adapter_merged.xlsx"
    adapter.save(output_path)
    adapter.close()

    wb = load_workbook(output_path)
    sheet = wb[merged_origin.sheet_name]
    assert sheet.cell(merged_origin.row, merged_origin.column).value == "MERGED_ORIGIN_OK"
    wb.close()


def test_save_nunca_sobreescribe_original(excel_path: Path, tmp_path: Path) -> None:
    original_hash_before = _sha256(excel_path)
    working_excel = _copy_to_tmp(excel_path, tmp_path)

    adapter = ExcelAdapter(working_excel)
    output_path = tmp_path / "excel_adapter_output.xlsx"
    adapter.save(output_path)
    adapter.close()

    assert output_path.exists()
    assert _sha256(excel_path) == original_hash_before
    assert output_path.resolve() != excel_path.resolve()


def _sha256(file_path: Path) -> str:
    return hashlib.sha256(file_path.read_bytes()).hexdigest()


def _copy_to_tmp(excel_path: Path, tmp_path: Path) -> Path:
    target = tmp_path / excel_path.name
    shutil.copy2(excel_path, target)
    return target


def _first_sheet_name(excel_path: Path) -> str:
    wb = load_workbook(excel_path)
    try:
        return wb.sheetnames[0]
    finally:
        wb.close()


def _find_first_merged_origin(excel_path: Path):
    wb = load_workbook(excel_path)
    try:
        for sheet in wb.worksheets:
            if not sheet.merged_cells.ranges:
                continue
            merged_range = next(iter(sheet.merged_cells.ranges))
            return CellPosition(
                sheet_name=sheet.title,
                row=merged_range.min_row,
                column=merged_range.min_col,
            )
        return None
    finally:
        wb.close()


class _DropdownTarget:
    def __init__(self, position, valid_option: str) -> None:
        self.position = position
        self.valid_option = valid_option


def _find_dropdown_target(excel_path: Path):
    wb = load_workbook(excel_path)
    try:
        for sheet in wb.worksheets:
            data_validations = getattr(sheet.data_validations, "dataValidation", [])
            for validation in data_validations:
                if getattr(validation, "type", None) != "list":
                    continue
                formula = (validation.formula1 or "").strip()
                if not (formula.startswith('"') and formula.endswith('"')):
                    continue
                options = [item.strip() for item in formula.strip('"').split(",") if item.strip()]
                if not options:
                    continue

                ranges = list(getattr(validation, "cells", []))
                if not ranges and getattr(validation, "sqref", None) is not None:
                    ranges = list(validation.sqref.ranges)
                if not ranges:
                    continue
                target_range = ranges[0]
                bounds = _extract_bounds(target_range)
                if bounds is None:
                    continue
                return _DropdownTarget(
                    position=CellPosition(
                        sheet_name=sheet.title,
                        row=bounds[0],
                        column=bounds[1],
                    ),
                    valid_option=options[0],
                )
        return None
    finally:
        wb.close()


def _extract_bounds(target_range) -> tuple[int, int] | None:
    if hasattr(target_range, "min_row") and hasattr(target_range, "min_col"):
        return (target_range.min_row, target_range.min_col)
    if isinstance(target_range, str):
        min_col, min_row, _, _ = range_boundaries(target_range)
        return (min_row, min_col)
    return None
