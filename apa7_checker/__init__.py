"""
apa7_checker
============
APA 7 reference checker for Word (.docx) documents.
"""

from .core import (
    YEAR_PATTERN,
    check_alphabetical_order,
    check_docx_references,
    classify_reference_type,
    cross_check_citations,
    detect_jammed_entries,
    find_references_section,
    generate_html_report,
    safe_write_file,
    split_references,
    validate_by_type,
    validate_doi_live,
    validate_italic_formatting,
    validate_reference_entry,
)

__version__ = "1.0.0"

__all__ = [
    "YEAR_PATTERN",
    "check_alphabetical_order",
    "check_docx_references",
    "classify_reference_type",
    "cross_check_citations",
    "detect_jammed_entries",
    "find_references_section",
    "generate_html_report",
    "safe_write_file",
    "split_references",
    "validate_by_type",
    "validate_doi_live",
    "validate_italic_formatting",
    "validate_reference_entry",
    "__version__",
]
