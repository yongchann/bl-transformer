"""
Regular expression patterns for invoice PDF parsing
"""
import re
from typing import Pattern


class InvoicePatterns:
    """Container for all invoice parsing patterns"""
    
    def __init__(self):
        # Pattern for EAN + description + weight + G
        self.item_step1: Pattern = re.compile(
            r'^(\d{13})\s+(.+?)\s+([\d,\.]+)\s+G',
            re.MULTILINE
        )
        
        # Pattern for G + quantity + unit_price + total_price + code + country + product_code
        self.item_step2: Pattern = re.compile(
            r'G\s+(\d+[\d,]*)\s+([\d,\.]+)\s+([\d,\.]+)\s+(\d+)\s+([A-Z]{2})\s+(\S+)$'
        )
        
        # Pattern for lines starting with 13 digits (potential items)
        self.ean_line: Pattern = re.compile(r'^\d{13}')
        
        # Metadata patterns
        self.shipment_number: Pattern = re.compile(r"Shipment Number: (\d+)")
        self.total_quantity: Pattern = re.compile(r"TOTAL QUANTITY (\d+)")
        
        # Terms of sale page detection
        self.terms_of_sale: str = "GENERAL TERMS OF SALE"


# Global instance for easy access
PATTERNS = InvoicePatterns()
