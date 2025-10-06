"""
Invoice data models for PDF parsing
"""
from typing import Dict, List, Optional


class InvoiceItem:
    """Individual item in an invoice"""
    
    def __init__(self):
        self.ean_number: Optional[str] = None
        self.description: Optional[str] = None
        self.quantity: Optional[str] = None
        self.unit_price: Optional[str] = None
        self.total_price_usd: Optional[str] = None
        self.country: Optional[str] = None
        self.product_code: Optional[str] = None

    def __str__(self) -> str:
        return (f"ean_number={self.ean_number}, description={self.description}, "
                f"quantity={self.quantity}, unit_price={self.unit_price}, "
                f"total_price_usd={self.total_price_usd}, country={self.country}, "
                f"product_code={self.product_code}")

    def to_dict(self) -> Dict:
        """Convert to dictionary for serialization"""
        return {
            'ean_number': self.ean_number,
            'description': self.description,
            'quantity': self.quantity,
            'unit_price': self.unit_price,
            'total_price_usd': self.total_price_usd,
            'country': self.country,
            'product_code': self.product_code
        }


class InvoiceData:
    """Complete invoice data including metadata and items"""
    
    def __init__(self):
        self.edi_number: Optional[str] = None
        self.delivery_number: Optional[str] = None
        self.invoice_number: Optional[str] = None
        self.invoice_date: Optional[str] = None
        self.shipment_number: Optional[str] = None
        self.total_quantity: Optional[str] = None
        self.items: Dict[str, InvoiceItem] = {}

    def append_items(self, items: Dict[str, InvoiceItem]) -> None:
        """Add items to the invoice"""
        for ean, item in items.items():
            self.items[ean] = item

    def set_metadata(self, metadata: Dict[str, str]) -> None:
        """Set invoice metadata"""
        self.edi_number = metadata.get("edi_number")
        self.delivery_number = metadata.get("delivery_number")
        self.invoice_number = metadata.get("invoice_number")
        self.invoice_date = metadata.get("invoice_date")
        
        # Only set if not already set (for multi-page invoices)
        if self.shipment_number is None:
            self.shipment_number = metadata.get("shipment_number")
        if self.total_quantity is None:
            self.total_quantity = metadata.get("total_quantity")

    def get_item_count(self) -> int:
        """Get total number of items"""
        return len(self.items)

    def to_dict(self) -> Dict:
        """Convert to dictionary for serialization"""
        return {
            'edi_number': self.edi_number,
            'delivery_number': self.delivery_number,
            'invoice_number': self.invoice_number,
            'invoice_date': self.invoice_date,
            'shipment_number': self.shipment_number,
            'total_quantity': self.total_quantity,
            'item_count': self.get_item_count(),
            'items': [item.to_dict() for item in self.items.values()]
        }

    def __str__(self) -> str:
        items_str = "\n".join(str(item) for item in self.items.values())
        return f"""
InvoiceData(
    edi_number={self.edi_number}
    delivery_number={self.delivery_number}
    invoice_number={self.invoice_number}
    invoice_date={self.invoice_date}
    shipment_number={self.shipment_number}
    total_quantity={self.total_quantity}
    item_count={self.get_item_count()}
    items=[
{items_str}
    ]
)"""
