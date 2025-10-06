"""
로레알 PDF → Excel 변환 도구
tkinter GUI를 사용한 Windows용 Python 애플리케이션
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from pdf_utils import read_pdf_text, get_bl_number_from_filename
from excel_utils import write_to_excel, get_output_directory


class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 로레알 PDF → Excel 변환 도구")
        self.root.geometry("500x300")
        self.root.resizable(False, False)
        
        # 변수 초기화
        self.pl_file_path = tk.StringVar()
        self.ci_file_path = tk.StringVar()
        self.output_filename = tk.StringVar()
        
        self.setup_ui()
    
    def setup_ui(self):
        """GUI 구성 요소를 설정합니다."""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 제목
        title_label = ttk.Label(
            main_frame, 
            text="📄 로레알 PDF → Excel 변환 도구",
            font=("맑은 고딕", 14, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # PL 파일 업로드
        ttk.Label(main_frame, text="PL 파일:").grid(row=1, column=0, sticky=tk.W, pady=5)
        pl_entry = ttk.Entry(main_frame, textvariable=self.pl_file_path, width=40, state="readonly")
        pl_entry.grid(row=1, column=1, padx=(10, 5), pady=5)
        ttk.Button(
            main_frame, 
            text="선택", 
            command=lambda: self.select_file("PL")
        ).grid(row=1, column=2, pady=5)
        
        # CI 파일 업로드
        ttk.Label(main_frame, text="CI 파일:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ci_entry = ttk.Entry(main_frame, textvariable=self.ci_file_path, width=40, state="readonly")
        ci_entry.grid(row=2, column=1, padx=(10, 5), pady=5)
        ttk.Button(
            main_frame, 
            text="선택", 
            command=lambda: self.select_file("CI")
        ).grid(row=2, column=2, pady=5)
        
        # 출력 파일명
        ttk.Label(main_frame, text="출력 파일명:").grid(row=3, column=0, sticky=tk.W, pady=5)
        output_entry = ttk.Entry(main_frame, textvariable=self.output_filename, width=40)
        output_entry.grid(row=3, column=1, padx=(10, 5), pady=5)
        ttk.Label(main_frame, text=".xlsx").grid(row=3, column=2, sticky=tk.W, pady=5)
        
        # 변환 실행 버튼
        convert_btn = ttk.Button(
            main_frame,
            text="🔄 변환 실행",
            command=self.convert_files,
            style="Accent.TButton"
        )
        convert_btn.grid(row=4, column=0, columnspan=3, pady=30)
        
        # 상태 표시
        self.status_label = ttk.Label(main_frame, text="파일을 선택해주세요.", foreground="gray")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=10)
        
        # 그리드 가중치 설정
        main_frame.columnconfigure(1, weight=1)
    
    def select_file(self, file_type):
        """파일 선택 다이얼로그를 엽니다."""
        file_path = filedialog.askopenfilename(
            title=f"{file_type} 파일 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if file_path:
            if file_type == "PL":
                self.pl_file_path.set(file_path)
            else:  # CI
                self.ci_file_path.set(file_path)
            
            # 첫 번째 파일이 선택되면 출력 파일명 자동 설정
            if not self.output_filename.get():
                bl_number = get_bl_number_from_filename(file_path)
                self.output_filename.set(bl_number)
            
            self.update_status()
    
    def update_status(self):
        """상태 메시지를 업데이트합니다."""
        pl_selected = bool(self.pl_file_path.get())
        ci_selected = bool(self.ci_file_path.get())
        
        if pl_selected and ci_selected:
            self.status_label.config(text="✅ 파일 준비 완료 - 변환을 실행하세요.", foreground="green")
        elif pl_selected or ci_selected:
            self.status_label.config(text="⚠️ 하나의 파일이 더 필요합니다.", foreground="orange")
        else:
            self.status_label.config(text="파일을 선택해주세요.", foreground="gray")
    
    def convert_files(self):
        """PDF 파일들을 Excel로 변환합니다."""
        pl_path = self.pl_file_path.get()
        ci_path = self.ci_file_path.get()
        output_name = self.output_filename.get().strip()
        
        # 입력 검증
        if not pl_path and not ci_path:
            messagebox.showerror("오류", "최소 하나의 PDF 파일을 선택해주세요.")
            return
        
        if not output_name:
            messagebox.showerror("오류", "출력 파일명을 입력해주세요.")
            return
        
        try:
            self.status_label.config(text="🔄 변환 중...", foreground="blue")
            self.root.update()
            
            # PDF 텍스트 추출
            pl_text = None
            ci_text = None
            
            if pl_path:
                pl_text = read_pdf_text(pl_path)
                
            if ci_path:
                ci_text = read_pdf_text(ci_path)
            
            # 출력 경로 설정 (첫 번째 파일과 같은 디렉토리)
            reference_path = pl_path if pl_path else ci_path
            output_dir = get_output_directory(reference_path)
            output_path = os.path.join(output_dir, f"{output_name}.xlsx")
            
            # Excel 파일 생성
            write_to_excel(output_path, pl_text, ci_text)
            
            # 성공 메시지
            self.status_label.config(text="✅ 변환 완료!", foreground="green")
            messagebox.showinfo(
                "변환 완료", 
                f"Excel 파일이 생성되었습니다:\\n{output_path}"
            )
            
        except Exception as e:
            self.status_label.config(text="❌ 변환 실패", foreground="red")
            messagebox.showerror("변환 오류", f"변환 중 오류가 발생했습니다:\\n{str(e)}")


def main():
    """메인 함수 - 애플리케이션을 시작합니다."""
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
