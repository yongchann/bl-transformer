"""
Packing List PDF parser with improved structure and error handling
"""
import pdfplumber
from typing import List, Dict, Optional, Tuple
import logging

from .models import PackingListItem
from .patterns import PATTERNS


class PackingListParser:
    """Main parser class for packing list PDFs"""
    
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
    
    def parse_pdf(self, pdf_path: str) -> List[PackingListItem]:
        """
        Parse packing list PDF and extract all data
        
        Args:
            pdf_path: Path to the PDF file
            
        Returns:
            List of PackingListItem objects
            
        Raises:
            Exception: If PDF reading fails
        """
        all_items = []
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    self.logger.info(f"Processing page {page_num}")

                    page_text = page.extract_text()
                    if not page_text:
                        self.logger.warning(f"No text found on page {page_num}")
                        continue
                    
                    # Extract common metadata for this page
                    common_data = self._extract_common_data(page_text)
                    
                    # Extract items from this page
                    items = self._extract_items(page_text)
                    
                    # Apply common data to all items
                    for item in items:
                        item.edi_number = common_data.get('edi_number')
                        item.order_number = common_data.get('order_number')
                        item.shipment_number = common_data.get('shipment_number')
                    
                    all_items.extend(items)
                    
        except Exception as e:
            raise Exception(f"PDF reading error: {str(e)}")
        
        # EAN + Batch 기준으로 아이템 그룹핑 및 수량 합산
        grouped_items = self._group_items_by_ean_batch(all_items)
        
        return grouped_items
    
    def _group_items_by_ean_batch(self, items: List[PackingListItem]) -> List[PackingListItem]:
        """
        EAN과 Batch가 동일한 아이템들을 그룹핑하여 수량 합산
        
        Args:
            items: 원본 아이템 리스트
            
        Returns:
            그룹핑된 아이템 리스트 (수량 합산됨)
        """
        grouped = {}
        
        for item in items:
            # EAN과 Batch를 키로 사용
            key = f"{item.ean}_{item.batch}"
            
            if key not in grouped:
                # 첫 번째 아이템은 그대로 저장
                grouped[key] = item
                if self.debug:
                    print(f"새 그룹 생성: {key} (수량: {item.items_qty})")
            else:
                # 기존 아이템에 수량 합산
                try:
                    existing_qty = int(grouped[key].items_qty) if grouped[key].items_qty else 0
                    additional_qty = int(item.items_qty) if item.items_qty else 0
                    total_qty = existing_qty + additional_qty
                    grouped[key].items_qty = str(total_qty)
                    
                    if self.debug:
                        print(f"수량 합산: {key} ({existing_qty} + {additional_qty} = {total_qty})")
                except ValueError:
                    if self.debug:
                        print(f"수량 변환 오류: {key}, 기존값 유지")
        
        result = list(grouped.values())
        
        if self.debug:
            print(f"그룹핑 결과: {len(items)}개 → {len(result)}개 아이템")
        
        return result
    
    def _extract_common_data(self, page_text: str) -> Dict[str, Optional[str]]:
        """Extract common metadata from page text"""
        common_data = {}
        
        # Extract EDI number
        edi_match = PATTERNS.edi_number.search(page_text)
        common_data['edi_number'] = edi_match.group(1) if edi_match else None
        
        # Extract order number
        order_match = PATTERNS.order_number.search(page_text)
        common_data['order_number'] = order_match.group(1) if order_match else None
        
        # Extract shipment number
        shipment_match = PATTERNS.shipment_number.search(page_text)
        common_data['shipment_number'] = shipment_match.group(1) if shipment_match else None
        
        if self.debug:
            print(f"Common data: {common_data}")
        
        return common_data
    
    def _extract_items(self, page_text: str) -> List[PackingListItem]:
        """Extract items from page text using regex patterns"""
        items = []
        matches_found = 0
        
        if self.debug:
            # Show lines that start with digits for debugging
            lines = page_text.split('\n')
            print(f"Lines starting with digits:")
            for i, line in enumerate(lines):
                if line.strip() and line.strip()[0].isdigit():
                    print(f"  Line {i}: {line.strip()}")
        
        # Try main pattern first
        for match in PATTERNS.item_line.finditer(page_text):
            matches_found += 1
            item = self._create_item_from_match(match)
            if item:
                items.append(item)
                if self.debug:
                    print(f"detected item {matches_found}: {item.ean} - {item.description}")
        
        # Try flexible pattern if no matches found
        if not items:
            for match in PATTERNS.item_line_flexible.finditer(page_text):
                matches_found += 1
                item = self._create_item_from_match(match)
                if item:
                    items.append(item)
                    if self.debug:
                        print(f"detected item (flexible) {matches_found}: {item.ean} - {item.description}")
        
        if self.debug:
            print(f"Total items found: {len(items)}")
        
        return items
    
    def _create_item_from_match(self, match) -> Optional[PackingListItem]:
        """Create PackingListItem from regex match"""
        try:
            item = PackingListItem()
            # hs_code, brand, sku, description, items_qty, ean, batch, mfg_date, exp_date, coo, dg
            item.hs_code = match.group(1)
            item.brand = match.group(2)
            item.sku = match.group(3)
            item.description = match.group(4).strip()
            # Remove commas from items_qty (e.g., "1,008" -> "1008")
            item.items_qty = match.group(5).replace(',', '')
            item.ean = match.group(6)
            item.batch = match.group(7)
            item.mfg_date = match.group(8)
            item.exp_date = match.group(9)
            item.coo = match.group(10)
            item.dg = match.group(11)
            
            return item
        except (IndexError, AttributeError) as e:
            if self.debug:
                print(f"Error creating item from match: {e}")
            return None
    

def parse_packing_list_pdf(pdf_path: str, debug: bool = False) -> List[PackingListItem]:
    """
    Convenience function to parse packing list PDF
    
    Args:
        pdf_path: Path to the PDF file
        debug: Enable debug output
        
    Returns:
        List of PackingListData objects
    """
    parser = PackingListParser(debug=debug)
    return parser.parse_pdf(pdf_path)
