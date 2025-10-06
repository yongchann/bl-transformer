"""
Packing List data models for PDF parsing
"""
from typing import Dict, List, Optional


class PackingListItem:
    """Individual item in a packing list"""
    
    def __init__(self):
        # Common metadata (shared across all items on the page)
        self.edi_number: Optional[str] = None
        self.order_number: Optional[str] = None
        self.shipment_number: Optional[str] = None
        
        # Item-specific data
        self.hs_code: Optional[str] = None
        self.brand: Optional[str] = None
        self.sku: Optional[str] = None
        self.description: Optional[str] = None
        self.items_qty: Optional[str] = None
        self.ean: Optional[str] = None
        self.batch: Optional[str] = None
        self.mfg_date: Optional[str] = None
        self.exp_date: Optional[str] = None
        self.coo: Optional[str] = None  # Country of Origin
        self.dg: Optional[str] = None   # Dangerous Goods

    def __str__(self) -> str:
        return (f"edi_number={self.edi_number}, order_number={self.order_number}, "
                f"shipment_number={self.shipment_number}, hs_code={self.hs_code}, "
                f"brand={self.brand}, sku={self.sku}, description={self.description}, "
                f"items_qty={self.items_qty}, ean={self.ean}, batch={self.batch}, "
                f"mfg_date={self.mfg_date}, exp_date={self.exp_date}, coo={self.coo}, dg={self.dg}")

    def to_dict(self) -> Dict:
        """Convert to dictionary for serialization"""
        return {
            'edi_number': self.edi_number,
            'order_number': self.order_number,
            'shipment_number': self.shipment_number,
            'hs_code': self.hs_code,
            'brand': self.brand,
            'sku': self.sku,
            'description': self.description,
            'items_qty': self.items_qty,
            'ean': self.ean,
            'batch': self.batch,
            'mfg_date': self.mfg_date,
            'exp_date': self.exp_date,
            'coo': self.coo,
            'dg': self.dg
        }



