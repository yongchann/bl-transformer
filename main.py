"""
Main script to parse PDF documents and create Excel output
"""
from document_parser import parse_pdf, DocumentType
from excel_utils import create_structured_excel
import os
import glob
import sys


def find_pdf_files():
    """현재 디렉토리에서 인보이스와 패킹리스트 PDF 파일을 찾습니다."""
    # 인보이스 파일 찾기 (CI로 끝나는 파일)
    invoice_files = glob.glob("*CI.pdf") + glob.glob("*CI.PDF")
    
    # 패킹리스트 파일 찾기 (PL로 끝나는 파일)
    packing_files = glob.glob("*PL.pdf") + glob.glob("*PL.PDF")
    
    return invoice_files, packing_files


def main():
    """Main function to parse PDF documents and create Excel output"""
    print("=== PDF Parser - Invoice & Packing List ===")
    print("현재 디렉토리에서 PDF 파일을 찾는 중...")
    
    # PDF 파일 자동 검색
    invoice_files, packing_files = find_pdf_files()
    
    if not invoice_files and not packing_files:
        print("❌ PDF 파일을 찾을 수 없습니다.")
        print("   다음 형식의 파일이 필요합니다:")
        print("   - 인보이스: *CI.pdf 또는 *CI.PDF")
        print("   - 패킹리스트: *PL.pdf 또는 *PL.PDF")
        if getattr(sys, 'frozen', False):
            input("\n아무 키나 누르면 종료됩니다...")
        return
    
    # 첫 번째 파일들 사용
    invoice_pdf = invoice_files[0] if invoice_files else None
    packing_pdf = packing_files[0] if packing_files else None
    
    # 출력 파일명 생성
    base_name = ""
    if invoice_pdf:
        base_name = invoice_pdf.replace(" CI.pdf", "").replace(" CI.PDF", "")
    elif packing_pdf:
        base_name = packing_pdf.replace(" PL.pdf", "").replace(" PL.PDF", "")
    
    output_excel = f"{base_name}_parsed_data.xlsx" if base_name else "parsed_data.xlsx"
    
    try:
        print(f"\n발견된 파일:")
        if invoice_pdf:
            print(f"   - 인보이스: {invoice_pdf}")
        if packing_pdf:
            print(f"   - 패킹리스트: {packing_pdf}")
        
        invoice_result = {'data': None, 'count': 0}
        packing_result = {'data': None, 'count': 0}
        
        # Parse Invoice PDF
        if invoice_pdf:
            print(f"\n1. 인보이스 파싱 중: {invoice_pdf}")
            invoice_result = parse_pdf(invoice_pdf, DocumentType.INVOICE, debug=False)
            print(f"   - {invoice_result['count']}개 인보이스 발견")
            
            if invoice_result['data']:
                total_items = sum(invoice.get_item_count() for invoice in invoice_result['data'])
                print(f"   - 총 {total_items}개 아이템")
        
        # Parse Packing List PDF
        if packing_pdf:
            print(f"\n2. 패킹리스트 파싱 중: {packing_pdf}")
            packing_result = parse_pdf(packing_pdf, DocumentType.PACKING_LIST, debug=False)
            print(f"   - {packing_result['count']}개 패킹리스트 아이템 발견")
        
        # Create Excel file
        print(f"\n3. Excel 파일 생성 중: {output_excel}")
        create_structured_excel(
            output_path=output_excel,
            invoices=invoice_result['data'] if invoice_result['data'] else None,
            packing_items=packing_result['data'] if packing_result['data'] else None
        )
        
        # Summary
        print(f"\n=== 처리 완료 ===")
        if invoice_result['data']:
            for i, invoice in enumerate(invoice_result['data'], 1):
                print(f"인보이스 {i}: {invoice.invoice_number} ({invoice.get_item_count()}개 아이템)")
        
        if packing_result['data']:
            # 동적으로 order_number 그룹 계산
            order_groups = {}
            for item in packing_result['data']:
                order_no = item.order_number
                if order_no not in order_groups:
                    order_groups[order_no] = 0
                order_groups[order_no] += 1
            
            group_info = " + ".join([f"{count}" for count in order_groups.values()])
            print(f"패킹리스트: {group_info} = {len(packing_result['data'])}개 아이템")
        
        print(f"\n✅ Excel 파일이 성공적으로 생성되었습니다: {output_excel}")
        print(f"   - Invoice 시트: {len(invoice_result['data']) if invoice_result['data'] else 0}개 인보이스")
        print(f"   - Packing_List 시트: {len(packing_result['data']) if packing_result['data'] else 0}개 아이템")
        
        # 실행파일에서만 입력 대기
        if getattr(sys, 'frozen', False):
            input("\n아무 키나 누르면 종료됩니다...")
            
    except Exception as e:
        print(f"❌ 파일 처리 중 오류 발생: {e}")
        if getattr(sys, 'frozen', False):
            input("\n아무 키나 누르면 종료됩니다...")


if __name__ == "__main__":
    main()