"""
Invoice PDF parser with improved structure and error handling
"""
import pdfplumber
from typing import List, Dict, Optional, Tuple
import logging

from .models import InvoiceData, InvoiceItem
from .patterns import PATTERNS


class InvoiceParser:
    """Main parser class for invoice PDFs"""
    
    def __init__(self, debug: bool = False):
        self.debug = debug
        self.logger = self._setup_logger()
    
    def _setup_logger(self) -> logging.Logger:
        """Setup logger for debugging"""
        logger = logging.getLogger(__name__)
        if self.debug and not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            logger.setLevel(logging.DEBUG)
        return logger
    
    def parse_pdf(self, pdf_path: str) -> List[InvoiceData]:
        """
        Parse invoice PDF and extract all invoice data
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            List of InvoiceData objects
            
        Raises:
            Exception: If PDF reading fails
        """
        invoice_data = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                invoice = InvoiceData()
                
                for page_num, page in enumerate(pdf.pages, 1):
                    self.logger.info(f"Processing page {page_num}")
                    page_text = page.extract_text()
                    
                    if not page_text:
                        self.logger.warning(f"No text found on page {page_num}")
                        continue
                    
                    # Check if this is a terms of sale page (end of invoice)
                    if self._is_terms_of_sale_page(page_text):
                        if invoice.invoice_number:  # Only add if we have data
                            invoice_data.append(invoice)
                            self.logger.info(f"Finished invoice {invoice.invoice_number}")
                        invoice = InvoiceData()
                        continue
                    
                    # Extract metadata and items
                    page_tables = page.extract_tables()
                    metadata = self._extract_metadata(page_tables, page_text)
                    invoice.set_metadata(metadata)
                    
                    items = self._extract_items(page_tables, page_text)
                    invoice.append_items(items)
                
                # Add the last invoice if it has data
                if invoice.invoice_number:
                    invoice_data.append(invoice)
                    
        except Exception as e:
            raise Exception(f"PDF reading error: {str(e)}")
        
        return invoice_data
    
    def _is_terms_of_sale_page(self, page_text: str) -> bool:
        """Check if page contains terms of sale"""
        return PATTERNS.terms_of_sale in page_text
    
    def _extract_metadata(self, page_tables: List, page_text: str) -> Dict[str, Optional[str]]:
        """
        Extract invoice metadata from tables and text
        
        Args:
            page_tables: List of tables extracted from the page
            page_text: Raw text from the page
            
        Returns:
            Dictionary containing metadata
        """
        metadata = {}
        
        # Extract from tables if available
        if len(page_tables) >= 3:
            try:
                metadata["edi_number"] = page_tables[0][2][1]
                metadata["delivery_number"] = page_tables[1][3][1]
                metadata["invoice_number"] = page_tables[2][0][1]
                metadata["invoice_date"] = page_tables[2][0][3]
            except (IndexError, TypeError) as e:
                self.logger.warning(f"Error extracting table metadata: {e}")
                # Set defaults
                metadata.update({
                    "edi_number": None,
                    "delivery_number": None,
                    "invoice_number": None,
                    "invoice_date": None
                })
        else:
            self.logger.warning("Insufficient tables for metadata extraction")
            metadata.update({
                "edi_number": None,
                "delivery_number": None,
                "invoice_number": None,
                "invoice_date": None
            })
        
        # Extract shipment number from text
        shipment_match = PATTERNS.shipment_number.search(page_text)
        metadata["shipment_number"] = shipment_match.group(1) if shipment_match else None
        
        # Extract total quantity from text
        quantity_match = PATTERNS.total_quantity.search(page_text)
        metadata["total_quantity"] = quantity_match.group(1) if quantity_match else None
        
        return metadata
    
    def _extract_items(self, page_tables: List, page_text: str) -> Dict[str, InvoiceItem]:
        """
        Extract invoice items using two-step regex matching
        
        Args:
            page_tables: List of tables from the page
            page_text: Raw text from the page
            
        Returns:
            Dictionary of items keyed by EAN number
        """
        items = {}
        
        # Skip if insufficient tables
        if len(page_tables) < 4:
            return items
        
        # Find potential item lines (starting with 13 digits)
        lines = page_text.split('\n')
        potential_lines = []
        
        for i, line in enumerate(lines):
            if PATTERNS.ean_line.match(line.strip()):
                potential_lines.append((i, line.strip()))
        
        # Process each potential line with two-step matching
        matches_found = 0
        for line_num, line in potential_lines:
            item = self._parse_item_line(line)
            if item:
                matches_found += 1
                items[item.ean_number] = item
                if self.debug:
                    print(f"detected item {matches_found}: {item.ean_number} - {item.description}")
        
        return items
    
    def _parse_item_line(self, line: str) -> Optional[InvoiceItem]:
        """
        Parse a single item line using two-step regex matching
        
        Args:
            line: Text line to parse
            
        Returns:
            InvoiceItem if parsing successful, None otherwise
        """
        # Step 1: EAN + description + weight + G
        match1 = PATTERNS.item_step1.search(line)
        if not match1:
            if self.debug:
                print(f"first step failed: {line}")
            return None
        
        # Step 2: G + quantity + unit_price + total_price + code + country + product_code
        match2 = PATTERNS.item_step2.search(line)
        if not match2:
            if self.debug:
                print(f"first step success, second step failed: {line}")
                print(f"  EAN={match1.group(1)}, description={match1.group(2)}, weight={match1.group(3)}")
            return None
        
        # Create item object
        item = InvoiceItem()
        item.ean_number = match1.group(1)
        item.description = match1.group(2).strip()
        item.quantity = match2.group(1)
        item.unit_price = match2.group(2)
        item.total_price_usd = match2.group(3)
        item.country = match2.group(5)
        item.product_code = match2.group(6)
        
        return item


def parse_invoice_pdf(pdf_path: str, debug: bool = False) -> List[InvoiceData]:
    """
    Convenience function to parse invoice PDF
    
    Args:
        pdf_path: Path to the PDF file
        debug: Enable debug output
        
    Returns:
        List of InvoiceData objects
    """
    parser = InvoiceParser(debug=debug)
    return parser.parse_pdf(pdf_path)
