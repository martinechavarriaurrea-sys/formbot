from __future__ import annotations

from pathlib import Path

import pytest

yaml_provider_module = pytest.importorskip(
    "formbot.infrastructure.parsers.yaml_mapping_provider",
    reason="YamlMappingProvider no disponible",
)

YamlMappingProvider = yaml_provider_module.YamlMappingProvider


def test_yaml_mapping_provider_soporta_campos_precision(tmp_path: Path) -> None:
    mapping_path = tmp_path / "mapping_precision.yaml"
    mapping_path.write_text(
        """
correo:
  label: Correo electronico
  aliases:
    - E-mail
    - Email
  offset:
    row: 0
    col: 0
  required: true
  type: email
  target_strategy: infer
  confidence_threshold: 0.93
  write_mode: mark
  mark_symbol: "✔"
  sheet: Formulario
""".strip(),
        encoding="utf-8",
    )

    rules = YamlMappingProvider(mapping_path).load()
    assert len(rules) == 1
    rule = rules[0]

    assert rule.field_name == "correo"
    assert rule.label == "Correo electronico"
    assert rule.aliases == ("E-mail", "Email")
    assert rule.value_type == "email"
    assert rule.target_strategy == "infer"
    assert rule.confidence_threshold == 0.93
    assert rule.write_mode == "mark"
    assert rule.mark_symbol == "✔"
    assert rule.sheet_name == "Formulario"
