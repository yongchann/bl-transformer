"""
PDF Parser GUI - Invoice & Packing List
tkinter를 사용한 GUI 버전
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
from pathlib import Path

from document_parser import parse_pdf, DocumentType
from excel_utils import create_structured_excel


class PDFParserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Parser - Invoice & Packing List")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # 파일 경로 저장
        self.invoice_file = None
        self.packing_file = None
        
        self.setup_ui()
        
    def setup_ui(self):
        """UI 구성 요소 설정"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(main_frame, text="PDF Parser", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        subtitle_label = ttk.Label(main_frame, text="Invoice & Packing List → Excel 변환", font=("Arial", 10))
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 30))
        
        # 인보이스 파일 선택
        ttk.Label(main_frame, text="인보이스 파일 (*CI.pdf):").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.invoice_var = tk.StringVar()
        self.invoice_entry = ttk.Entry(main_frame, textvariable=self.invoice_var, width=50, state="readonly")
        self.invoice_entry.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(main_frame, text="파일 선택", command=self.select_invoice_file).grid(row=3, column=2, padx=(10, 0), pady=5)
        
        # 패킹리스트 파일 선택
        ttk.Label(main_frame, text="패킹리스트 파일 (*PL.pdf):").grid(row=4, column=0, sticky=tk.W, pady=(20, 5))
        
        self.packing_var = tk.StringVar()
        self.packing_entry = ttk.Entry(main_frame, textvariable=self.packing_var, width=50, state="readonly")
        self.packing_entry.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(main_frame, text="파일 선택", command=self.select_packing_file).grid(row=5, column=2, padx=(10, 0), pady=5)
        
        # 출력 파일 설정
        ttk.Label(main_frame, text="출력 Excel 파일명:").grid(row=6, column=0, sticky=tk.W, pady=(20, 5))
        
        self.output_var = tk.StringVar(value="parsed_data.xlsx")
        self.output_entry = ttk.Entry(main_frame, textvariable=self.output_var, width=50)
        self.output_entry.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(main_frame, text="저장 위치", command=self.select_output_file).grid(row=7, column=2, padx=(10, 0), pady=5)
        
        # 변환 버튼
        self.convert_button = ttk.Button(main_frame, text="📄 Excel로 변환", command=self.start_conversion, style="Accent.TButton")
        self.convert_button.grid(row=8, column=0, columnspan=3, pady=(30, 10), sticky=(tk.W, tk.E))
        
        # 진행 상태 표시
        self.progress_var = tk.StringVar(value="파일을 선택하고 변환 버튼을 클릭하세요.")
        self.progress_label = ttk.Label(main_frame, textvariable=self.progress_var, font=("Arial", 9))
        self.progress_label.grid(row=9, column=0, columnspan=3, pady=5)
        
        # 진행률 바
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=10, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 결과 텍스트 영역
        self.result_text = tk.Text(main_frame, height=10, width=70, wrap=tk.WORD)
        self.result_text.grid(row=11, column=0, columnspan=3, pady=(20, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.result_text.yview)
        scrollbar.grid(row=11, column=3, sticky=(tk.N, tk.S))
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(11, weight=1)
        
    def select_invoice_file(self):
        """인보이스 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="인보이스 파일 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=os.getcwd()
        )
        if file_path:
            self.invoice_file = file_path
            self.invoice_var.set(os.path.basename(file_path))
            self.update_output_filename()
            
    def select_packing_file(self):
        """패킹리스트 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="패킹리스트 파일 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=os.getcwd()
        )
        if file_path:
            self.packing_file = file_path
            self.packing_var.set(os.path.basename(file_path))
            self.update_output_filename()
            
    def select_output_file(self):
        """출력 파일 저장 위치 선택"""
        file_path = filedialog.asksaveasfilename(
            title="Excel 파일 저장 위치",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=os.getcwd()
        )
        if file_path:
            self.output_var.set(file_path)
            
    def update_output_filename(self):
        """선택된 파일을 기반으로 출력 파일명 자동 생성"""
        if self.invoice_file or self.packing_file:
            # 기본 이름 추출
            base_name = ""
            if self.invoice_file:
                base_name = Path(self.invoice_file).stem.replace(" CI", "")
            elif self.packing_file:
                base_name = Path(self.packing_file).stem.replace(" PL", "")
            
            if base_name:
                output_name = f"{base_name}_parsed_data.xlsx"
                self.output_var.set(output_name)
                
    def start_conversion(self):
        """변환 작업을 별도 스레드에서 시작"""
        if not self.invoice_file and not self.packing_file:
            messagebox.showwarning("파일 선택", "최소 하나의 PDF 파일을 선택해주세요.")
            return
            
        if not self.output_var.get().strip():
            messagebox.showwarning("출력 파일", "출력 파일명을 입력해주세요.")
            return
            
        # UI 비활성화
        self.convert_button.config(state="disabled")
        self.progress_bar.start()
        self.result_text.delete(1.0, tk.END)
        
        # 별도 스레드에서 변환 작업 실행
        thread = threading.Thread(target=self.convert_files)
        thread.daemon = True
        thread.start()
        
    def convert_files(self):
        """실제 파일 변환 작업"""
        try:
            self.update_progress("변환 작업을 시작합니다...")
            
            invoice_result = {'data': None, 'count': 0}
            packing_result = {'data': None, 'count': 0}
            
            # 인보이스 파일 처리
            if self.invoice_file:
                self.update_progress(f"인보이스 파일 파싱 중: {os.path.basename(self.invoice_file)}")
                invoice_result = parse_pdf(self.invoice_file, DocumentType.INVOICE, debug=False)
                self.log_result(f"✅ 인보이스: {invoice_result['count']}개 발견")
                
                if invoice_result['data']:
                    total_items = sum(invoice.get_item_count() for invoice in invoice_result['data'])
                    self.log_result(f"   총 {total_items}개 아이템")
            
            # 패킹리스트 파일 처리
            if self.packing_file:
                self.update_progress(f"패킹리스트 파일 파싱 중: {os.path.basename(self.packing_file)}")
                packing_result = parse_pdf(self.packing_file, DocumentType.PACKING_LIST, debug=False)
                self.log_result(f"✅ 패킹리스트: {packing_result['count']}개 아이템 발견")
            
            # Excel 파일 생성
            output_path = self.output_var.get()
            self.update_progress(f"Excel 파일 생성 중: {output_path}")
            
            create_structured_excel(
                output_path=output_path,
                invoices=invoice_result['data'] if invoice_result['data'] else None,
                packing_items=packing_result['data'] if packing_result['data'] else None
            )
            
            # 완료 메시지
            self.update_progress("✅ 변환 완료!")
            self.log_result(f"\n🎉 Excel 파일이 성공적으로 생성되었습니다!")
            self.log_result(f"📁 파일 위치: {os.path.abspath(output_path)}")
            self.log_result(f"📊 Invoice 시트: {len(invoice_result['data']) if invoice_result['data'] else 0}개 인보이스")
            self.log_result(f"📦 Packing_List 시트: {len(packing_result['data']) if packing_result['data'] else 0}개 아이템")
            
            # 완료 후 파일 열기 옵션
            self.root.after(0, lambda: self.show_completion_dialog(output_path))
            
        except Exception as e:
            self.update_progress("❌ 변환 중 오류 발생")
            self.log_result(f"오류: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("오류", f"변환 중 오류가 발생했습니다:\n{str(e)}"))
        
        finally:
            # UI 다시 활성화
            self.root.after(0, self.reset_ui)
            
    def update_progress(self, message):
        """진행 상태 업데이트 (스레드 안전)"""
        self.root.after(0, lambda: self.progress_var.set(message))
        
    def log_result(self, message):
        """결과 텍스트 영역에 메시지 추가 (스레드 안전)"""
        def add_text():
            self.result_text.insert(tk.END, message + "\n")
            self.result_text.see(tk.END)
        self.root.after(0, add_text)
        
    def reset_ui(self):
        """UI 상태 초기화"""
        self.convert_button.config(state="normal")
        self.progress_bar.stop()
        
    def show_completion_dialog(self, output_path):
        """완료 후 파일 열기 옵션 제공"""
        result = messagebox.askyesno(
            "변환 완료", 
            f"Excel 파일이 성공적으로 생성되었습니다!\n\n{os.path.basename(output_path)}\n\n파일을 열어보시겠습니까?"
        )
        if result:
            try:
                os.startfile(output_path)  # Windows
            except AttributeError:
                try:
                    os.system(f"open '{output_path}'")  # macOS
                except:
                    os.system(f"xdg-open '{output_path}'")  # Linux


def main():
    """GUI 애플리케이션 시작"""
    root = tk.Tk()
    
    # 아이콘 설정 (있는 경우)
    try:
        root.iconbitmap("icon.ico")
    except:
        pass
    
    app = PDFParserGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
