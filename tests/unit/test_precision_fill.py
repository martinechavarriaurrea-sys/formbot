from __future__ import annotations

from pathlib import Path

import pytest

openpyxl = pytest.importorskip("openpyxl", reason="openpyxl no disponible")

precision_module = pytest.importorskip(
    "formbot.application.precision_fill",
    reason="Modo precision no disponible",
)
domain_models = pytest.importorskip("formbot.domain.models")
exceptions_module = pytest.importorskip("formbot.domain.exceptions")

PrecisionFillUseCase = precision_module.PrecisionFillUseCase
MappingRule = domain_models.MappingRule
DataValidationError = exceptions_module.DataValidationError


def test_precision_offset_explicito_escribe_con_confianza_alta(tmp_path: Path) -> None:
    template = tmp_path / "template_offset.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A1", "Nombre / Razón social:"),
        ],
    )

    rules = [
        MappingRule(
            field_name="razon_social",
            label="Nombre / Razón social:",
            row_offset=0,
            column_offset=1,
            sheet_name="Formulario",
            required=True,
            value_type="text",
            target_strategy="offset",
        )
    ]
    data = {"razon_social": "ACME INDUSTRIAL SAS"}
    output = tmp_path / "precision_offset_output.xlsx"

    use_case = PrecisionFillUseCase(template_path=template, strict_mode=True, min_confidence=0.85)
    try:
        result = use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert output.exists()
    assert result.blocked_fields == []
    assert "razon_social" in result.written_fields

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Formulario"]["B1"].value == "ACME INDUSTRIAL SAS"
    finally:
        wb.close()


def test_precision_infiere_celda_derecha_si_no_hay_offset(tmp_path: Path) -> None:
    template = tmp_path / "template_infer.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A2", "E-mail:"),
        ],
    )

    rules = [
        MappingRule(
            field_name="correo",
            label="Correo electrónico",
            aliases=("E-mail", "Email"),
            row_offset=0,
            column_offset=0,
            sheet_name="Formulario",
            required=True,
            value_type="email",
            target_strategy="infer",
        )
    ]
    data = {"correo": "proveedor@acme.co"}
    output = tmp_path / "precision_infer_output.xlsx"

    use_case = PrecisionFillUseCase(template_path=template, strict_mode=True, min_confidence=0.8)
    try:
        result = use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert output.exists()
    assert result.blocked_fields == []
    assert "correo" in result.written_fields

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Formulario"]["B2"].value == "proveedor@acme.co"
    finally:
        wb.close()


def test_precision_strict_bloquea_si_hay_label_ambiguo(tmp_path: Path) -> None:
    template = tmp_path / "template_ambiguo.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A1", "NIT"),
            ("Formulario", "A3", "NIT"),
        ],
    )

    rules = [
        MappingRule(
            field_name="nit",
            label="NIT",
            row_offset=0,
            column_offset=1,
            sheet_name="Formulario",
            required=True,
            value_type="nit",
            target_strategy="offset_or_infer",
        )
    ]
    data = {"nit": "900123456"}
    output = tmp_path / "precision_ambiguous_output.xlsx"

    use_case = PrecisionFillUseCase(template_path=template, strict_mode=True, min_confidence=0.85)
    try:
        with pytest.raises(DataValidationError):
            use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert not output.exists()


def test_precision_nit_duplicado_prioriza_contexto_de_identificacion(tmp_path: Path) -> None:
    template = tmp_path / "template_nit_context.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A1", "NIT"),
            ("Formulario", "B1", "Gravado"),
            ("Formulario", "A3", "NIT"),
            ("Formulario", "C3", "Numero:"),
        ],
    )

    rules = [
        MappingRule(
            field_name="nit",
            label="NIT",
            row_offset=0,
            column_offset=0,
            sheet_name="Formulario",
            required=True,
            value_type="nit",
            target_strategy="infer",
        )
    ]
    data = {"nit": "890900240"}
    output = tmp_path / "precision_nit_context_output.xlsx"

    use_case = PrecisionFillUseCase(template_path=template, strict_mode=True, min_confidence=0.8)
    try:
        result = use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert output.exists()
    assert result.blocked_fields == []
    assert "nit" in result.written_fields

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Formulario"]["B3"].value == "890900240"
    finally:
        wb.close()


def test_precision_infer_no_sobrescribe_numerico_existente_por_defecto(tmp_path: Path) -> None:
    template = tmp_path / "template_no_overwrite_default.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A1", "NIT"),
            ("Formulario", "B1", "111111111"),
        ],
    )

    rules = [
        MappingRule(
            field_name="nit",
            label="NIT",
            row_offset=0,
            column_offset=0,
            sheet_name="Formulario",
            required=True,
            value_type="nit",
            target_strategy="infer",
        )
    ]
    data = {"nit": "890900240"}
    output = tmp_path / "precision_no_overwrite_default.xlsx"

    use_case = PrecisionFillUseCase(
        template_path=template,
        strict_mode=True,
        min_confidence=0.85,
    )
    try:
        result = use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert output.exists()
    assert result.blocked_fields == []
    assert "nit" in result.written_fields

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Formulario"]["B1"].value == "111111111"
        assert wb["Formulario"]["C1"].value == "890900240"
    finally:
        wb.close()


def test_precision_infer_bloquea_si_solo_hay_celdas_numericas_ocupadas(tmp_path: Path) -> None:
    template = tmp_path / "template_anchor_low_quality.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A1", "NIT"),
            ("Formulario", "A2", "Nombre:"),
            ("Formulario", "A3", "Direccion:"),
            ("Formulario", "A4", "Telefono:"),
            ("Formulario", "B1", "111111111"),
            ("Formulario", "C1", "222222222"),
            ("Formulario", "D1", "333333333"),
            ("Formulario", "E1", "444444444"),
            ("Formulario", "F1", "555555555"),
            ("Formulario", "G1", "666666666"),
            ("Formulario", "H1", "777777777"),
        ],
    )

    rules = [
        MappingRule(
            field_name="nit",
            label="NIT",
            row_offset=0,
            column_offset=0,
            sheet_name="Formulario",
            required=True,
            value_type="nit",
            target_strategy="infer",
            confidence_threshold=0.7,
        )
    ]
    data = {"nit": "890900240"}
    output = tmp_path / "precision_anchor_low_quality.xlsx"

    use_case = PrecisionFillUseCase(
        template_path=template,
        strict_mode=True,
        min_confidence=0.7,
    )
    try:
        with pytest.raises(DataValidationError):
            use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert not output.exists()


def test_precision_infer_permite_sobrescribir_si_flag_esta_activo(tmp_path: Path) -> None:
    template = tmp_path / "template_allow_overwrite.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A1", "NIT"),
            ("Formulario", "B1", "111111111"),
            ("Formulario", "C1", "222222222"),
            ("Formulario", "D1", "333333333"),
        ],
    )

    rules = [
        MappingRule(
            field_name="nit",
            label="NIT",
            row_offset=0,
            column_offset=0,
            sheet_name="Formulario",
            required=True,
            value_type="nit",
            target_strategy="infer",
            confidence_threshold=0.7,
        )
    ]
    data = {"nit": "890900240"}
    output = tmp_path / "precision_allow_overwrite.xlsx"

    use_case = PrecisionFillUseCase(
        template_path=template,
        strict_mode=True,
        min_confidence=0.7,
        allow_overwrite_existing=True,
    )
    try:
        result = use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert output.exists()
    assert result.blocked_fields == []
    assert "nit" in result.written_fields

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Formulario"]["B1"].value == "890900240"
    finally:
        wb.close()


def test_precision_mark_si_no_compuesto_marca_columna_correcta(tmp_path: Path) -> None:
    template = tmp_path / "template_mark_si_no.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A1", "¿Realiza operaciones internacionales?"),
            ("Formulario", "C1", "Si           No"),
        ],
    )

    rules = [
        MappingRule(
            field_name="op_internacionales",
            label="¿Realiza operaciones internacionales?",
            row_offset=0,
            column_offset=0,
            sheet_name="Formulario",
            required=True,
            target_strategy="offset_or_infer",
            write_mode="mark",
            mark_symbol="X",
        )
    ]
    data = {"op_internacionales": "si"}
    output = tmp_path / "precision_mark_si_no.xlsx"

    use_case = PrecisionFillUseCase(
        template_path=template,
        strict_mode=True,
        min_confidence=0.75,
    )
    try:
        result = use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert output.exists()
    assert result.blocked_fields == []
    assert "op_internacionales" in result.written_fields

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Formulario"]["B1"].value == "X"
        assert wb["Formulario"]["D1"].value is None
    finally:
        wb.close()


def test_precision_mark_permite_simbolo_personalizado_y_no_booleano(tmp_path: Path) -> None:
    template = tmp_path / "template_mark_custom.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A1", "¿Maneja recursos públicos?"),
            ("Formulario", "C1", "Si         No"),
        ],
    )

    rules = [
        MappingRule(
            field_name="pep_recursos",
            label="¿Maneja recursos públicos?",
            row_offset=0,
            column_offset=0,
            sheet_name="Formulario",
            required=True,
            target_strategy="offset_or_infer",
            write_mode="mark",
            mark_symbol="✔",
        )
    ]
    data = {"pep_recursos": False}
    output = tmp_path / "precision_mark_custom.xlsx"

    use_case = PrecisionFillUseCase(
        template_path=template,
        strict_mode=True,
        min_confidence=0.75,
    )
    try:
        result = use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert output.exists()
    assert result.blocked_fields == []
    assert "pep_recursos" in result.written_fields

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Formulario"]["D1"].value == "✔"
        assert wb["Formulario"]["B1"].value is None
    finally:
        wb.close()


def test_precision_mark_opcion_textual_no_binaria(tmp_path: Path) -> None:
    template = tmp_path / "template_mark_text_option.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A1", "Tipo de contraparte"),
            ("Formulario", "D1", "Proveedor   Cliente   Accionista"),
        ],
    )

    rules = [
        MappingRule(
            field_name="tipo_contraparte",
            label="Tipo de contraparte",
            row_offset=0,
            column_offset=0,
            sheet_name="Formulario",
            required=True,
            target_strategy="offset_or_infer",
            write_mode="mark",
            mark_symbol="X",
        )
    ]
    data = {"tipo_contraparte": "cliente"}
    output = tmp_path / "precision_mark_text_option.xlsx"

    use_case = PrecisionFillUseCase(
        template_path=template,
        strict_mode=True,
        min_confidence=0.75,
    )
    try:
        result = use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert output.exists()
    assert result.blocked_fields == []
    assert "tipo_contraparte" in result.written_fields

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Formulario"]["C1"].value == "X"
    finally:
        wb.close()


def test_precision_mark_detecta_opcion_en_fila_vecina(tmp_path: Path) -> None:
    template = tmp_path / "template_mark_near_row.xlsx"
    _create_workbook(
        template,
        [
            ("Formulario", "A1", "¿Maneja activos virtuales?"),
            ("Formulario", "C2", "Si           No"),
        ],
    )

    rules = [
        MappingRule(
            field_name="activos_virtuales",
            label="¿Maneja activos virtuales?",
            row_offset=0,
            column_offset=0,
            sheet_name="Formulario",
            required=True,
            target_strategy="offset_or_infer",
            write_mode="mark",
            mark_symbol="X",
        )
    ]
    data = {"activos_virtuales": True}
    output = tmp_path / "precision_mark_near_row.xlsx"

    use_case = PrecisionFillUseCase(
        template_path=template,
        strict_mode=True,
        min_confidence=0.75,
    )
    try:
        result = use_case.execute(data=data, mapping_rules=rules, output_path=output)
    finally:
        use_case.close()

    assert output.exists()
    assert result.blocked_fields == []
    assert "activos_virtuales" in result.written_fields

    wb = openpyxl.load_workbook(output)
    try:
        assert wb["Formulario"]["B2"].value == "X"
    finally:
        wb.close()


def _create_workbook(path: Path, seeded_cells: list[tuple[str, str, str]]) -> None:
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    default_sheet.title = "Formulario"

    for sheet_name, cell_ref, value in seeded_cells:
        if sheet_name not in wb.sheetnames:
            wb.create_sheet(sheet_name)
        wb[sheet_name][cell_ref] = value

    wb.save(path)
    wb.close()
