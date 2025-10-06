"""
Regular expression patterns for packing list PDF parsing
"""
import re
from typing import Pattern


class PackingListPatterns:
    """Container for all packing list parsing patterns"""
    
    def __init__(self):
        # Common metadata patterns
        self.edi_number: Pattern = re.compile(
            r'Your\s+Reference\s+([A-Z0-9]+)',
            re.IGNORECASE
        )
        
        self.order_number: Pattern = re.compile(
            r'Order\s+Number\s*:\s*(\d+)',
            re.IGNORECASE
        )
        
        self.shipment_number: Pattern = re.compile(
            r'Ship\s+Group\s+ID\s*:\s*(\d+)',
            re.IGNORECASE
        )
        
        # Main item pattern based on the actual data format
        # hs_code, brand, sku, description, items_qty, ean, batch, mfg_date, exp_date, coo, dg
        # items_qty can have commas (e.g., 1,008)
        self.item_line: Pattern = re.compile(
            r'^(\d+)\s+(\w+)\s+(\S+)\s+(.+?)\s+([\d,]+)\s+(\d{13})\s+(\S+)\s+(\d{2}-\d{2}-\d{4})\s+(\d{2}-\d{2}-\d{4})\s+([A-Z]{1,2})\s+([YN])',
            re.MULTILINE
        )
        
        # More flexible pattern to handle line breaks and spacing
        # items_qty can have commas (e.g., 1,008)
        self.item_line_flexible: Pattern = re.compile(
            r'(\d+)\s+(\w+)\s+(\S+)\s+(.+?)\s+([\d,]+)\s+(\d{13})\s+(\S+)\s+(\d{2}-\d{2}-\d{4})\s+(\d{2}-\d{2}-\d{4})\s+([A-Z]{1,2})\s*\n?([YN])',
            re.MULTILINE | re.DOTALL
        )


# Global instance for easy access
PATTERNS = PackingListPatterns()
