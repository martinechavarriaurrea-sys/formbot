class FormBotError(Exception):
    """Base exception for FormBot."""


class DocumentProcessingError(FormBotError):
    """Raised when a template cannot be loaded or parsed."""


class LabelNotFoundError(FormBotError):
    """Raised when a configured label is not found in the document."""


class MappingRuleError(FormBotError):
    """Raised when mapping configuration is invalid or ambiguous."""


class PositionOutOfBoundsError(FormBotError):
    """Raised when the computed target position is invalid for the document."""


class DataValidationError(FormBotError):
    """Raised when input data is invalid."""


class ValidationException(DataValidationError):
    """Raised when a value violates a field-level validation rule."""


class DocumentSaveError(FormBotError):
    """Raised when the output file cannot be persisted."""
