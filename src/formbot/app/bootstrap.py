from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from formbot.application.fill_form import FillFormUseCase
from formbot.domain.exceptions import DocumentProcessingError
from formbot.domain.models import MappingRule
from formbot.domain.ports.document_adapter import DocumentAdapter
from formbot.infrastructure.document_readers.excel_document_adapter import ExcelDocumentAdapter
from formbot.infrastructure.mappers.label_offset_mapper import LabelOffsetMapper
from formbot.infrastructure.parsers.json_data_provider import JsonFileDataProvider
from formbot.infrastructure.parsers.yaml_mapping_provider import YamlMappingProvider

# Extensiones soportadas y su MIME type de salida correspondiente
SUPPORTED_EXTENSIONS: dict[str, str] = {
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".xlsm": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pdf":  "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


@dataclass(frozen=True)
class PipelineContext:
    use_case: FillFormUseCase
    mapping_rules: list[MappingRule]
    data: dict[str, Any]
    output_extension: str   # ".xlsx", ".pdf" o ".docx"
    mime_type: str


def create_document_adapter(template_path: Path) -> DocumentAdapter:
    """Instancia el adaptador correcto según la extensión del template."""
    suffix = template_path.suffix.lower()

    if suffix in {".xlsx", ".xlsm"}:
        return ExcelDocumentAdapter(template_path=template_path)

    if suffix == ".pdf":
        # Import diferido: pypdf puede no estar instalado
        from formbot.infrastructure.document_readers.pdf_document_adapter import (
            PdfDocumentAdapter,
        )
        return PdfDocumentAdapter(template_path=template_path)

    if suffix == ".docx":
        # Import diferido: python-docx puede no estar instalado
        from formbot.infrastructure.document_readers.word_document_adapter import (
            WordDocumentAdapter,
        )
        return WordDocumentAdapter(template_path=template_path)

    if suffix == ".xls":
        raise DocumentProcessingError(
            f"El formato .xls (Excel 97-2003) no está soportado. "
            f"Convierta el archivo a .xlsx o .xlsm: {template_path}"
        )

    supported = ", ".join(SUPPORTED_EXTENSIONS)
    raise DocumentProcessingError(
        f"Formato '{suffix}' no soportado. Use uno de: {supported}"
    )


def bootstrap_pipeline(
    template_path: Path,
    mapping_path: Path,
    data_path: Path,
) -> PipelineContext:
    """Bootstrap genérico: crea el pipeline correcto para Excel, PDF o Word."""
    suffix = template_path.suffix.lower()
    mime_type = SUPPORTED_EXTENSIONS.get(suffix, "application/octet-stream")

    document_adapter = create_document_adapter(template_path)
    try:
        mapping_provider = YamlMappingProvider(file_path=mapping_path)
        data_provider = JsonFileDataProvider(file_path=data_path)
        field_mapper = LabelOffsetMapper()

        use_case = FillFormUseCase(
            document_adapter=document_adapter,
            field_mapper=field_mapper,
        )
        mapping_rules = mapping_provider.load()
        data = data_provider.load()

        return PipelineContext(
            use_case=use_case,
            mapping_rules=mapping_rules,
            data=data,
            output_extension=suffix,
            mime_type=mime_type,
        )
    except Exception:
        document_adapter.close()
        raise


# ------------------------------------------------------------------
# Alias de compatibilidad hacia atrás — el código existente que
# importa bootstrap_excel_pipeline sigue funcionando sin cambios.
# ------------------------------------------------------------------

@dataclass(frozen=True)
class ExcelPipelineContext:
    use_case: FillFormUseCase
    mapping_rules: list[MappingRule]
    data: dict[str, Any]


def bootstrap_excel_pipeline(
    template_path: Path,
    mapping_path: Path,
    data_path: Path,
) -> ExcelPipelineContext:
    """Bootstrap exclusivo para Excel. Mantenido por compatibilidad con código existente."""
    document_adapter = ExcelDocumentAdapter(template_path=template_path)
    try:
        mapping_provider = YamlMappingProvider(file_path=mapping_path)
        data_provider = JsonFileDataProvider(file_path=data_path)
        field_mapper = LabelOffsetMapper()

        use_case = FillFormUseCase(
            document_adapter=document_adapter,
            field_mapper=field_mapper,
        )
        mapping_rules = mapping_provider.load()
        data = data_provider.load()

        return ExcelPipelineContext(
            use_case=use_case,
            mapping_rules=mapping_rules,
            data=data,
        )
    except Exception:
        document_adapter.close()
        raise
