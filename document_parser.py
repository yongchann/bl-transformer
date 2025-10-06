"""
Unified document parser for both invoices and packing lists
"""
from typing import Union, List, Dict, Any
from enum import Enum
import os

from invoice import parse_invoice_pdf, InvoiceData
from packing_list import parse_packing_list_pdf, PackingListItem


class DocumentType(Enum):
    """Supported document types"""
    INVOICE = "invoice"
    PACKING_LIST = "packing_list"
    AUTO = "auto"


class DocumentParser:
    """Unified parser for different document types"""
    
    def __init__(self, debug: bool = False):
        self.debug = debug
    
    def parse_document(
        self, 
        pdf_path: str, 
        doc_type: DocumentType = DocumentType.AUTO
    ) -> Dict[str, Any]:
        """
        Parse a PDF document and return structured data
        
        Args:
            pdf_path: Path to the PDF file
            doc_type: Type of document to parse (auto-detect if AUTO)
            
        Returns:
            Dictionary containing parsed data and metadata
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        
        # Auto-detect document type if needed
        if doc_type == DocumentType.AUTO:
            doc_type = self._detect_document_type(pdf_path)
        
        # Parse based on document type
        if doc_type == DocumentType.INVOICE:
            data = parse_invoice_pdf(pdf_path, debug=self.debug)
            return {
                'document_type': 'invoice',
                'file_path': pdf_path,
                'data': data,
                'count': len(data)
            }
        
        elif doc_type == DocumentType.PACKING_LIST:
            data = parse_packing_list_pdf(pdf_path, debug=self.debug)
            return {
                'document_type': 'packing_list',
                'file_path': pdf_path,
                'data': data,
                'count': len(data)
            }
        
        else:
            raise ValueError(f"Unsupported document type: {doc_type}")
    
    def _detect_document_type(self, pdf_path: str) -> DocumentType:
        """
        Auto-detect document type based on filename and content
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            Detected document type
        """
        filename = os.path.basename(pdf_path).lower()
        
        # Check filename patterns
        if any(keyword in filename for keyword in ['invoice', 'ci', 'commercial']):
            return DocumentType.INVOICE
        
        if any(keyword in filename for keyword in ['packing', 'pl', 'pack']):
            return DocumentType.PACKING_LIST
        
        # If filename doesn't give clear indication, try content-based detection
        try:
            import pdfplumber
            with pdfplumber.open(pdf_path) as pdf:
                if pdf.pages:
                    first_page_text = pdf.pages[0].extract_text() or ""
                    first_page_upper = first_page_text.upper()
                    
                    # Look for invoice indicators
                    invoice_keywords = [
                        'COMMERCIAL INVOICE', 'INVOICE', 'BILL TO', 'SHIP TO',
                        'INVOICE NUMBER', 'INVOICE DATE', 'EAN', 'UNIT PRICE'
                    ]
                    
                    # Look for packing list indicators
                    packing_keywords = [
                        'PACKING LIST', 'PACKING', 'SHIPPER', 'CONSIGNEE',
                        'VESSEL', 'VOYAGE', 'PORT OF LOADING', 'GROSS WEIGHT'
                    ]
                    
                    invoice_score = sum(1 for keyword in invoice_keywords if keyword in first_page_upper)
                    packing_score = sum(1 for keyword in packing_keywords if keyword in first_page_upper)
                    
                    if invoice_score > packing_score:
                        return DocumentType.INVOICE
                    elif packing_score > invoice_score:
                        return DocumentType.PACKING_LIST
        
        except Exception as e:
            print(f"Warning: Could not analyze PDF content for type detection: {e}")
        
        # Default to invoice if uncertain
        return DocumentType.INVOICE
    
    def parse_multiple_documents(
        self, 
        pdf_paths: List[str], 
        doc_types: Union[List[DocumentType], DocumentType] = DocumentType.AUTO
    ) -> List[Dict[str, Any]]:
        """
        Parse multiple PDF documents
        
        Args:
            pdf_paths: List of PDF file paths
            doc_types: Document types (single type for all, or list matching pdf_paths)
            
        Returns:
            List of parsed document data
        """
        results = []
        
        # Handle single doc_type for all files
        if isinstance(doc_types, DocumentType):
            doc_types = [doc_types] * len(pdf_paths)
        
        # Ensure doc_types list matches pdf_paths length
        if len(doc_types) != len(pdf_paths):
            raise ValueError("doc_types list must match pdf_paths length")
        
        for pdf_path, doc_type in zip(pdf_paths, doc_types):
            try:
                result = self.parse_document(pdf_path, doc_type)
                results.append(result)
            except Exception as e:
                results.append({
                    'document_type': 'error',
                    'file_path': pdf_path,
                    'error': str(e),
                    'data': None,
                    'count': 0
                })
        
        return results


def parse_pdf(
    pdf_path: str, 
    doc_type: DocumentType = DocumentType.AUTO, 
    debug: bool = False
) -> Dict[str, Any]:
    """
    Convenience function to parse a single PDF document
    
    Args:
        pdf_path: Path to the PDF file
        doc_type: Type of document to parse
        debug: Enable debug output
        
    Returns:
        Parsed document data
    """
    parser = DocumentParser(debug=debug)
    return parser.parse_document(pdf_path, doc_type)


def parse_multiple_pdfs(
    pdf_paths: List[str], 
    doc_types: Union[List[DocumentType], DocumentType] = DocumentType.AUTO,
    debug: bool = False
) -> List[Dict[str, Any]]:
    """
    Convenience function to parse multiple PDF documents
    
    Args:
        pdf_paths: List of PDF file paths
        doc_types: Document types
        debug: Enable debug output
        
    Returns:
        List of parsed document data
    """
    parser = DocumentParser(debug=debug)
    return parser.parse_multiple_documents(pdf_paths, doc_types)
