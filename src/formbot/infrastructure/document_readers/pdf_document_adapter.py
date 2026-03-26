from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from formbot.domain.exceptions import (
    DocumentProcessingError,
    DocumentSaveError,
    LabelNotFoundError,
    MappingRuleError,
    PositionOutOfBoundsError,
)
from formbot.domain.models import CellPosition
from formbot.domain.ports.document_adapter import DocumentAdapter
from formbot.shared.utils import normalize_text

LOGGER = logging.getLogger(__name__)

# Nombre sintético de "hoja" usado en CellPosition para identificar campos AcroForm.
# No representa una hoja real — es un marcador que write_value usa para enrutar la escritura.
_PDF_SHEET = "pdf_acroform"

try:
    from pypdf import PdfReader, PdfWriter

    _PYPDF_AVAILABLE = True
except ImportError:  # pragma: no cover
    _PYPDF_AVAILABLE = False


class PdfDocumentAdapter(DocumentAdapter):
    """Adaptador para formularios PDF con campos AcroForm interactivos (fillable PDFs).

    Estrategia de búsqueda de etiquetas:
    - Compara el texto buscado contra el nombre del campo AcroForm (normalizado).
    - Soporta coincidencia exacta y parcial, igual que el adaptador Excel.
    - Soporta alias: si el primer término no coincide, intenta los siguientes.

    Estrategia de escritura:
    - find_label() devuelve una CellPosition sintética (sheet="pdf_acroform", row=N, column=1).
    - El mapper aplica el offset configurado (para PDF normalmente row_offset=0, column_offset=0
      porque la etiqueta Y el campo de escritura son el mismo objeto AcroForm).
    - write_value() almacena el valor; save() aplica todos los cambios al PDF de salida.

    Limitaciones conocidas:
    - Solo funciona con PDFs que tienen campos AcroForm interactivos.
    - PDFs estáticos (escaneados o sin campos) no son soportados.
    - El modo precisión (PrecisionFillUseCase) es Excel-exclusivo — usar FillFormUseCase.
    - El parámetro sheet_name de find_label() se ignora (los campos no están ligados a páginas
      en la API de AcroForm; un mismo campo puede aparecer en múltiples páginas).
    """

    def __init__(self, template_path: Path) -> None:
        if not _PYPDF_AVAILABLE:
            raise DocumentProcessingError(
                "La librería 'pypdf' no está instalada. "
                "Instálela con:  pip install pypdf>=4.3.1"
            )
        if not template_path.exists():
            raise DocumentProcessingError(
                f"No existe el template PDF: {template_path}"
            )
        if template_path.suffix.lower() not in {".pdf"}:
            raise DocumentProcessingError(
                f"Se esperaba un archivo .pdf, se recibió '{template_path.suffix}': {template_path}"
            )
        try:
            self._reader: Any = PdfReader(str(template_path))
        except Exception as exc:
            raise DocumentProcessingError(
                f"No fue posible leer el PDF: {template_path}"
            ) from exc

        self._template_path = template_path
        # Registro sintético: CellPosition → nombre local del campo AcroForm (/T)
        self._position_to_field: dict[CellPosition, str] = {}
        # Actualizaciones pendientes: nombre local → valor a escribir
        self._pending_updates: dict[str, str] = {}
        self._counter: int = 1

        # Verificar que el PDF tiene campos AcroForm
        fields = self._reader.get_fields() or {}
        if not fields:
            LOGGER.warning(
                "El PDF '%s' no tiene campos AcroForm detectados. "
                "Solo los PDFs interactivos (fillable) son soportados.",
                template_path.name,
            )
        else:
            LOGGER.debug(
                "PDF '%s' cargado con %d campos AcroForm.",
                template_path.name,
                len(fields),
            )

    # ------------------------------------------------------------------
    # Interfaz pública (DocumentAdapter)
    # ------------------------------------------------------------------

    def find_label(self, text: str, sheet_name: str | None = None) -> CellPosition:
        if sheet_name is not None:
            LOGGER.warning(
                "PdfDocumentAdapter: sheet_name='%s' ignorado — los campos AcroForm "
                "no están asociados a páginas específicas en esta implementación.",
                sheet_name,
            )

        normalized_target = normalize_text(text)
        if not normalized_target:
            raise LabelNotFoundError("No se puede buscar una etiqueta vacía")

        fields = self._reader.get_fields() or {}
        if not fields:
            raise DocumentProcessingError(
                f"El PDF no tiene campos AcroForm rellenables: {self._template_path}"
            )

        exact: list[tuple[str, str]] = []   # (qualified_name, local_name)
        partial: list[tuple[str, str]] = []
        seen_qualified: set[str] = set()

        for qualified, field_obj in fields.items():
            if qualified in seen_qualified:
                continue
            seen_qualified.add(qualified)

            local = str(field_obj.get("/T", "")).strip()

            # Buscar coincidencia en: nombre local, nombre completo, última parte del nombre completo
            candidates = {local, qualified, qualified.split(".")[-1]}
            matched_as: str | None = None
            for candidate_name in candidates:
                norm = normalize_text(candidate_name)
                if not norm:
                    continue
                if norm == normalized_target:
                    matched_as = "exact"
                    break
                if matched_as is None and (
                    normalized_target in norm or norm in normalized_target
                ):
                    matched_as = "partial"

            if matched_as == "exact":
                exact.append((qualified, local))
            elif matched_as == "partial":
                partial.append((qualified, local))

        if len(exact) == 1:
            qualified, local = exact[0]
        elif len(exact) > 1:
            names = ", ".join(q for q, _ in exact)
            raise MappingRuleError(
                f"Etiqueta ambigua '{text}': {len(exact)} campos AcroForm coinciden exactamente "
                f"({names})"
            )
        elif len(partial) == 1:
            qualified, local = partial[0]
        elif len(partial) > 1:
            names = ", ".join(q for q, _ in partial)
            raise MappingRuleError(
                f"Etiqueta ambigua '{text}': {len(partial)} campos AcroForm coinciden parcialmente "
                f"({names})"
            )
        else:
            available = ", ".join(list(fields.keys())[:10])
            suffix = "..." if len(fields) > 10 else ""
            raise LabelNotFoundError(
                f"No se encontró campo AcroForm para '{text}'. "
                f"Campos disponibles: {available}{suffix}"
            )

        LOGGER.debug(
            "Campo AcroForm encontrado para '%s': qualified='%s', local='%s'",
            text,
            qualified,
            local,
        )

        position = CellPosition(sheet_name=_PDF_SHEET, row=self._counter, column=1)
        self._position_to_field[position] = local
        self._counter += 1
        return position

    def write_value(self, position: CellPosition, value: Any) -> None:
        if position.sheet_name != _PDF_SHEET:
            raise PositionOutOfBoundsError(
                f"Posición no pertenece a un campo AcroForm: {position.sheet_name}"
            )
        local_name = self._position_to_field.get(position)
        if local_name is None:
            raise PositionOutOfBoundsError(
                f"Posición PDF no registrada (row={position.row}). "
                "Asegúrese de llamar find_label() antes de write_value()."
            )
        str_value = "" if value is None else str(value)
        self._pending_updates[local_name] = str_value
        LOGGER.debug("Campo AcroForm '%s' marcado para escritura: '%s'", local_name, str_value)

    def save(self, output_path: Path) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            writer: Any = PdfWriter(clone_from=self._reader)
        except Exception as exc:
            raise DocumentSaveError(
                f"No fue posible crear el escritor PDF: {exc}"
            ) from exc

        if self._pending_updates:
            for page in writer.pages:
                try:
                    writer.update_page_form_field_values(
                        page,
                        self._pending_updates,
                        auto_regenerate=False,
                    )
                except Exception as exc:  # pragma: no cover — depende del PDF
                    LOGGER.warning(
                        "Error al actualizar campos en página: %s", exc
                    )

        try:
            with open(output_path, "wb") as fh:
                writer.write(fh)
        except OSError as exc:
            raise DocumentSaveError(
                f"No fue posible guardar el PDF de salida: {output_path}"
            ) from exc

        LOGGER.info(
            "PDF guardado en '%s' con %d campo(s) actualizado(s).",
            output_path,
            len(self._pending_updates),
        )

    def close(self) -> None:
        # PdfReader no requiere cierre explícito
        return None
