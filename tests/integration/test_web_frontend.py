from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest

fastapi_testclient = pytest.importorskip(
    "fastapi.testclient",
    reason="fastapi no disponible",
)
openpyxl = pytest.importorskip("openpyxl", reason="openpyxl no disponible")

web_module = pytest.importorskip(
    "formbot.web.app",
    reason="Frontend web no disponible",
)

TestClient = fastapi_testclient.TestClient
app = web_module.app


def test_frontend_index_responde_html() -> None:
    client = TestClient(app)
    response = client.get("/")
    assert response.status_code == 200
    assert "FormBot" in response.text
    assert "/api/fill" in response.text


def test_frontend_api_fill_retorna_excel(
    fixtures_dir: Path,
) -> None:
    client = TestClient(app)
    template_path = fixtures_dir / "fixture_form.xlsx"
    mapping_path = fixtures_dir / "pipeline_mapping.yaml"
    data_path = fixtures_dir / "pipeline_data_valid.json"

    if not template_path.exists() or not mapping_path.exists() or not data_path.exists():
        pytest.skip("Faltan fixtures requeridos para probar frontend web")

    with (
        template_path.open("rb") as template_fp,
        mapping_path.open("rb") as mapping_fp,
        data_path.open("rb") as data_fp,
    ):
        response = client.post(
            "/api/fill",
            files={
                "template": (
                    template_path.name,
                    template_fp,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
                "mapping": (mapping_path.name, mapping_fp, "application/x-yaml"),
                "data": (data_path.name, data_fp, "application/json"),
            },
        )

    assert response.status_code == 200
    assert (
        response.headers.get("content-type")
        == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    content_disposition = response.headers.get("content-disposition", "")
    assert "attachment;" in content_disposition
    assert "_filled_" in content_disposition

    workbook = openpyxl.load_workbook(filename=BytesIO(response.content))
    try:
        assert workbook.sheetnames
    finally:
        workbook.close()
