"""Auto-detección de campos rellenables en documentos (Excel, PDF, Word).

El escáner analiza la estructura del documento para identificar etiquetas
(labels) que tienen celdas/campos vacíos adyacentes donde se deben escribir datos.

Retorna una lista de DetectedField con:
- field_key: clave snake_case única para usar en el payload de datos
- label: texto original de la etiqueta tal como aparece en el documento
- sheet: nombre de la hoja (Excel) o None (PDF/Word)
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from formbot.shared.utils import normalize_text

# Número máximo de columnas a escanear hacia la derecha buscando celda vacía
_INFER_RIGHT_MAX = 8
# Número máximo de filas a escanear hacia abajo buscando celda vacía
_INFER_DOWN_MAX = 3
# Longitud mínima del texto de la etiqueta
_MIN_LABEL_LEN = 2

# Palabras clave conocidas que indican que un texto corto es una etiqueta de formulario
_KNOWN_LABEL_TOKENS = {
    "nit", "nombre", "ciudad", "pais", "fecha", "correo", "email",
    "celular", "telefono", "telefax", "banco", "cargo", "firma",
    "tipo", "numero", "cuenta", "rut", "contacto", "codigo",
    "actividad", "digito", "cupo", "departamento", "dv",
    "plazo", "representante", "identificacion", "direccion",
}


@dataclass
class DetectedField:
    """Campo rellenable detectado automáticamente en un documento."""

    field_key: str       # Clave snake_case (única dentro del documento)
    label: str           # Texto original de la etiqueta en el documento
    sheet: str | None    # Nombre de la hoja (Excel) o None (PDF/Word)


# ---------------------------------------------------------------------------
# Punto de entrada público
# ---------------------------------------------------------------------------

def scan_document(template_path: Path) -> list[DetectedField]:
    """Detecta automáticamente los campos rellenables en el documento."""
    suffix = template_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return _scan_excel(template_path)
    if suffix == ".pdf":
        return _scan_pdf(template_path)
    if suffix == ".docx":
        return _scan_word(template_path)
    return []


def label_to_key(text: str) -> str:
    """Convierte el texto de una etiqueta a una clave snake_case (máx. 60 chars)."""
    normalized = normalize_text(text)
    return "_".join(normalized.split())[:60]


# ---------------------------------------------------------------------------
# Escáner Excel
# ---------------------------------------------------------------------------

def _scan_excel(path: Path) -> list[DetectedField]:
    try:
        from openpyxl import load_workbook
        from openpyxl.cell.cell import MergedCell
    except ImportError:
        return []

    try:
        wb = load_workbook(path, data_only=True)
    except Exception:
        return []

    fields: list[DetectedField] = []
    seen_norm: set[str] = set()

    try:
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell, MergedCell):
                        continue
                    raw = cell.value
                    if raw is None:
                        continue
                    text = str(raw).strip()
                    if len(text) < _MIN_LABEL_LEN:
                        continue
                    if _is_numeric(text):
                        continue
                    if not _is_likely_form_label(text):
                        continue

                    norm = normalize_text(text)
                    if norm in seen_norm:
                        continue

                    if not _has_adjacent_empty_excel(sheet, cell.row, cell.column):
                        continue

                    seen_norm.add(norm)
                    fields.append(DetectedField(
                        field_key=label_to_key(text),
                        label=text,
                        sheet=sheet.title,
                    ))
    finally:
        wb.close()

    return fields


def _is_numeric(text: str) -> bool:
    """True si el texto es puramente numérico (ignora separadores)."""
    clean = re.sub(r"[\s.,\-/]", "", text)
    return bool(clean) and clean.isdigit()


def _is_likely_form_label(text: str) -> bool:
    """Heurística: ¿es este texto probablemente una etiqueta de formulario?"""
    # Termina en dos puntos → definitivamente una etiqueta
    if text.rstrip().endswith(":"):
        return True
    # Texto multi-palabra de longitud razonable → probablemente etiqueta
    if " " in text and 4 <= len(text) <= 120:
        return True
    # Palabras clave cortas conocidas como etiquetas
    normalized = normalize_text(text)
    if normalized in _KNOWN_LABEL_TOKENS:
        return True
    # Contiene algún token clave
    for token in _KNOWN_LABEL_TOKENS:
        if token in normalized:
            return True
    return False


def _has_adjacent_empty_excel(sheet: object, row: int, col: int) -> bool:
    """True si hay al menos una celda vacía adyacente (derecha o abajo)."""
    try:
        from openpyxl.cell.cell import MergedCell
    except ImportError:
        return False

    # Escanear hacia la derecha
    for dc in range(1, _INFER_RIGHT_MAX + 1):
        try:
            cell = sheet.cell(row=row, column=col + dc)  # type: ignore[union-attr]
        except Exception:
            break
        if isinstance(cell, MergedCell):
            continue
        val = cell.value
        if val is None or str(val).strip() == "":
            return True
        # Celda no vacía: detenemos el escaneo hacia la derecha
        break

    # Escanear hacia abajo
    for dr in range(1, _INFER_DOWN_MAX + 1):
        try:
            cell = sheet.cell(row=row + dr, column=col)  # type: ignore[union-attr]
        except Exception:
            break
        if isinstance(cell, MergedCell):
            continue
        val = cell.value
        if val is None or str(val).strip() == "":
            return True
        break

    return False


# ---------------------------------------------------------------------------
# Escáner PDF (campos AcroForm)
# ---------------------------------------------------------------------------

def _scan_pdf(path: Path) -> list[DetectedField]:
    try:
        from pypdf import PdfReader
    except ImportError:
        return []

    try:
        reader = PdfReader(str(path))
        raw_fields = reader.get_fields() or {}
    except Exception:
        return []

    result: list[DetectedField] = []
    seen: set[str] = set()

    for qualified, field_obj in raw_fields.items():
        # Usar nombre local (/T) preferentemente; caer en último segmento del nombre calificado
        local = str(field_obj.get("/T", "")).strip()
        if not local:
            local = qualified.split(".")[-1].strip()
        if not local or local in seen:
            continue
        seen.add(local)
        result.append(DetectedField(
            field_key=label_to_key(local),
            label=local,
            sheet=None,
        ))

    return result


# ---------------------------------------------------------------------------
# Escáner Word (tablas)
# ---------------------------------------------------------------------------

def _scan_word(path: Path) -> list[DetectedField]:
    try:
        from docx import Document as DocxDocument
    except ImportError:
        return []

    try:
        doc = DocxDocument(str(path))
    except Exception:
        return []

    fields: list[DetectedField] = []
    seen_norm: set[str] = set()

    for table in doc.tables:
        for row in table.rows:
            cells = row.cells
            for i, cell in enumerate(cells[:-1]):
                text = cell.text.strip()
                if len(text) < _MIN_LABEL_LEN:
                    continue
                if _is_numeric(text):
                    continue
                if not _is_likely_form_label(text):
                    continue
                # La celda siguiente debe estar vacía (campo a rellenar)
                next_text = cells[i + 1].text.strip()
                if next_text:
                    continue
                norm = normalize_text(text)
                if norm in seen_norm:
                    continue
                seen_norm.add(norm)
                fields.append(DetectedField(
                    field_key=label_to_key(text),
                    label=text,
                    sheet=None,
                ))

    return fields
