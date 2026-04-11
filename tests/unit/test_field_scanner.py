"""Tests unitarios para el módulo field_scanner.

Cubre:
  - B2: _is_all_caps_multi_word  → rechaza valores como nombres empresa/persona
  - B3: _is_decorative_text      → rechaza textos instructivos/decorativos
  - _is_likely_form_label        → acepta labels válidos, rechaza todo lo demás
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from formbot.infrastructure.document_scanners.field_scanner import (
    _is_all_caps_multi_word,
    _is_decorative_text,
    _is_likely_form_label,
)


# ---------------------------------------------------------------------------
# B2 — _is_all_caps_multi_word
# ---------------------------------------------------------------------------

class TestIsAllCapsMultiWord:
    """FIX: B2 — Los valores en mayúsculas no deben confundirse con labels."""

    def test_nombre_empresa_rechazado(self) -> None:
        assert _is_all_caps_multi_word("LABORATORIOS PHARMA SA") is True

    def test_nombre_persona_rechazado(self) -> None:
        assert _is_all_caps_multi_word("JORGE ENRIQUE SALCEDO MORA") is True

    def test_tres_palabras_sin_token_rechazado(self) -> None:
        assert _is_all_caps_multi_word("DISTRIBUIDORA REGIONAL SAS") is True

    # Excepciones legítimas
    def test_con_token_banco_no_rechazado(self) -> None:
        # Contiene "banco" → puede ser encabezado de columna válido
        assert _is_all_caps_multi_word("BANCO COMERCIAL SAS") is False

    def test_con_slash_no_rechazado(self) -> None:
        # "/" indica estructura de label (NIT/CC/CE)
        assert _is_all_caps_multi_word("NIT/CC/CE EMPRESA") is False

    def test_con_dos_puntos_no_rechazado(self) -> None:
        assert _is_all_caps_multi_word("NOMBRE:") is False

    def test_dos_palabras_no_rechazado(self) -> None:
        # Menos de 3 palabras: puede ser encabezado de columna corto
        assert _is_all_caps_multi_word("RAZON SOCIAL") is False

    def test_una_sola_palabra_no_rechazado(self) -> None:
        assert _is_all_caps_multi_word("BANCO") is False

    def test_minusculas_no_rechazado(self) -> None:
        assert _is_all_caps_multi_word("Razon social empresa") is False

    def test_mixto_no_rechazado(self) -> None:
        assert _is_all_caps_multi_word("Razon Social Empresa") is False


# ---------------------------------------------------------------------------
# B3 — _is_decorative_text
# ---------------------------------------------------------------------------

class TestIsDecorativeText:
    """FIX: B3 — Los textos instructivos/decorativos no deben ser campos."""

    def test_instruccion_por_favor(self) -> None:
        assert _is_decorative_text("Por favor complete todos los campos del formulario") is True

    def test_instruccion_nota(self) -> None:
        assert _is_decorative_text("Nota: todos los campos marcados son obligatorios") is True

    def test_instruccion_importante(self) -> None:
        assert _is_decorative_text("Importante: adjunte copia del RUT") is True

    def test_instruccion_diligenciar(self) -> None:
        assert _is_decorative_text("Diligenciar con letra legible") is True

    def test_instruccion_complete(self) -> None:
        assert _is_decorative_text("Complete la siguiente información con datos verídicos") is True

    def test_texto_muy_largo(self) -> None:
        # Más de 90 chars → párrafo, no label
        texto = "Este es un texto demasiado largo para ser una etiqueta de campo en un formulario empresarial"
        assert len(texto) > 90
        assert _is_decorative_text(texto) is True

    # Labels válidos que NO deben ser rechazados
    def test_razon_social_no_decorativo(self) -> None:
        assert _is_decorative_text("Razón social") is False

    def test_nit_no_decorativo(self) -> None:
        assert _is_decorative_text("NIT") is False

    def test_banco_no_decorativo(self) -> None:
        assert _is_decorative_text("Banco") is False

    def test_fecha_no_decorativa(self) -> None:
        assert _is_decorative_text("Fecha de diligenciamiento") is False

    def test_correo_electronico_no_decorativo(self) -> None:
        assert _is_decorative_text("Correo electrónico") is False

    def test_representante_legal_no_decorativo(self) -> None:
        assert _is_decorative_text("Nombre del representante legal") is False


# ---------------------------------------------------------------------------
# _is_likely_form_label — integración de B2 + B3
# ---------------------------------------------------------------------------

class TestIsLikelyFormLabel:
    """Verifica que B2 y B3 funcionan correctamente dentro de la heurística principal."""

    # B2: valores en mayúsculas rechazados
    def test_b2_empresa_mayusculas_rechazado(self) -> None:
        assert _is_likely_form_label("LABORATORIOS PHARMA SA") is False

    def test_b2_persona_mayusculas_rechazado(self) -> None:
        assert _is_likely_form_label("JORGE ENRIQUE SALCEDO MORA") is False

    # B3: textos decorativos rechazados
    def test_b3_instruccion_rechazada(self) -> None:
        assert _is_likely_form_label("Por favor complete todos los campos") is False

    def test_b3_texto_largo_rechazado(self) -> None:
        texto = (
            "Este campo es de uso exclusivo del area de contabilidad "
            "y finanzas segun la circular interna numero 001 del presente ano vigente"
        )
        assert len(texto) > 90, f"El texto de prueba debe tener >90 chars, tiene {len(texto)}"
        assert _is_likely_form_label(texto) is False

    # Labels válidos que DEBEN ser aceptados
    def test_razon_social_aceptado(self) -> None:
        assert _is_likely_form_label("Razón social") is True

    def test_razon_social_minusculas_aceptado(self) -> None:
        assert _is_likely_form_label("Razon social") is True

    def test_nit_aceptado(self) -> None:
        assert _is_likely_form_label("NIT") is True

    def test_nit_cc_slash_aceptado(self) -> None:
        assert _is_likely_form_label("NIT/CC/CE") is True

    def test_correo_electronico_aceptado(self) -> None:
        assert _is_likely_form_label("Correo electrónico") is True

    def test_representante_legal_aceptado(self) -> None:
        assert _is_likely_form_label("Nombre del representante legal") is True

    def test_banco_aceptado(self) -> None:
        assert _is_likely_form_label("Banco") is True

    def test_numero_cuenta_aceptado(self) -> None:
        assert _is_likely_form_label("Número de cuenta") is True

    def test_tipo_de_persona_aceptado(self) -> None:
        assert _is_likely_form_label("Tipo de persona") is True

    def test_fecha_diligenciamiento_aceptado(self) -> None:
        assert _is_likely_form_label("Fecha de diligenciamiento") is True

    def test_con_dos_puntos_aceptado(self) -> None:
        assert _is_likely_form_label("Nombre del proveedor:") is True

    def test_nit_con_dv_aceptado(self) -> None:
        assert _is_likely_form_label("C.C. / NIT con DV") is True
