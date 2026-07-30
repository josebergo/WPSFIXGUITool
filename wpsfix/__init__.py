"""WPS-compatible DOCX repair engine."""

from .engine import ConversionReport, convert_docx, inspect_docx, unique_output_path

__all__ = ["ConversionReport", "convert_docx", "inspect_docx", "unique_output_path"]
