"""
Legacy PDF utilities - maintained for backward compatibility
This module is deprecated. Use document_parser.py instead.
"""
import warnings
from typing import List

from invoice import parse_invoice_pdf, InvoiceData


def extract_invoice_data(pdf_path: str) -> List[InvoiceData]:
    """
    Legacy function for backward compatibility
    
    Args:
        pdf_path: Path to the PDF file
        
    Returns:
        List of InvoiceData objects
    """
    warnings.warn(
        "pdf_utils.extract_invoice_data is deprecated. Use document_parser.parse_pdf instead.",
        DeprecationWarning,
        stacklevel=2
    )
    return parse_invoice_pdf(pdf_path, debug=True)