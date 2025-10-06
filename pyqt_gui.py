"""
PDF Parser GUI - PyQt5 버전
더 현대적이고 안정적인 GUI
"""
import sys
import os
from pathlib import Path
import threading
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QPushButton, QLineEdit, 
                            QTextEdit, QProgressBar, QFileDialog, QMessageBox,
                            QGroupBox, QGridLayout, QFrame)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QUrl
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor, QDragEnterEvent, QDropEvent

from document_parser import parse_pdf, DocumentType
from excel_utils import create_structured_excel


class DragDropLineEdit(QLineEdit):
    """드래그 앤 드롭을 지원하는 QLineEdit"""
    
    file_dropped = pyqtSignal(str)  # 파일이 드롭되었을 때 시그널
    
    def __init__(self, placeholder_text=""):
        super().__init__()
        self.setAcceptDrops(True)
        self.setReadOnly(True)
        self.setPlaceholderText(placeholder_text)
        self.setStyleSheet("""
            QLineEdit {
                border: 2px dashed #bdc3c7;
                border-radius: 8px;
                padding: 10px;
                background-color: #f8f9fa;
                color: #2c3e50;
            }
            QLineEdit:hover {
                border-color: #3498db;
                background-color: #e3f2fd;
            }
            QLineEdit[readOnly="true"] {
                background-color: #f8f9fa;
            }
        """)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """드래그 진입 이벤트"""
        if event.mimeData().hasUrls():
            # PDF 파일인지 확인
            urls = event.mimeData().urls()
            if urls and urls[0].toLocalFile().lower().endswith('.pdf'):
                event.acceptProposedAction()
                self.setStyleSheet("""
                    QLineEdit {
                        border: 2px solid #27ae60;
                        border-radius: 8px;
                        padding: 10px;
                        background-color: #d5f4e6;
                        color: #2c3e50;
                    }
                """)
            else:
                event.ignore()
        else:
            event.ignore()
    
    def dragLeaveEvent(self, event):
        """드래그 떠남 이벤트"""
        self.setStyleSheet("""
            QLineEdit {
                border: 2px dashed #bdc3c7;
                border-radius: 8px;
                padding: 10px;
                background-color: #f8f9fa;
                color: #2c3e50;
            }
            QLineEdit:hover {
                border-color: #3498db;
                background-color: #e3f2fd;
            }
        """)
    
    def dropEvent(self, event: QDropEvent):
        """드롭 이벤트"""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                file_path = urls[0].toLocalFile()
                if file_path.lower().endswith('.pdf'):
                    self.setText(os.path.basename(file_path))
                    self.file_dropped.emit(file_path)
                    event.acceptProposedAction()
                    
                    # 스타일 복원
                    self.dragLeaveEvent(event)
                else:
                    event.ignore()
        else:
            event.ignore()


class ConversionWorker(QThread):
    """변환 작업을 위한 워커 스레드"""
    progress_update = pyqtSignal(str)
    log_update = pyqtSignal(str)
    finished = pyqtSignal(bool, str)  # success, message
    
    def __init__(self, invoice_file, packing_file, output_file):
        super().__init__()
        self.invoice_file = invoice_file
        self.packing_file = packing_file
        self.output_file = output_file
        
    def run(self):
        """변환 작업 실행"""
        try:
            self.progress_update.emit("변환 작업을 시작합니다...")
            
            invoice_result = {'data': None, 'count': 0}
            packing_result = {'data': None, 'count': 0}
            
            # 인보이스 파일 처리
            if self.invoice_file:
                self.progress_update.emit(f"인보이스 파일 파싱 중: {os.path.basename(self.invoice_file)}")
                invoice_result = parse_pdf(self.invoice_file, DocumentType.INVOICE, debug=False)
                self.log_update.emit(f"✅ 인보이스: {invoice_result['count']}개 발견")
                
                if invoice_result['data']:
                    total_items = sum(invoice.get_item_count() for invoice in invoice_result['data'])
                    self.log_update.emit(f"   총 {total_items}개 아이템")
            
            # 패킹리스트 파일 처리
            if self.packing_file:
                self.progress_update.emit(f"패킹리스트 파일 파싱 중: {os.path.basename(self.packing_file)}")
                packing_result = parse_pdf(self.packing_file, DocumentType.PACKING_LIST, debug=False)
                self.log_update.emit(f"✅ 패킹리스트: {packing_result['count']}개 아이템 발견")
            
            # Excel 파일 생성
            self.progress_update.emit(f"Excel 파일 생성 중: {os.path.basename(self.output_file)}")
            
            create_structured_excel(
                output_path=self.output_file,
                invoices=invoice_result['data'] if invoice_result['data'] else None,
                packing_items=packing_result['data'] if packing_result['data'] else None
            )
            
            # 완료 메시지
            self.progress_update.emit("✅ 변환 완료!")
            self.log_update.emit(f"\n🎉 Excel 파일이 성공적으로 생성되었습니다!")
            self.log_update.emit(f"📁 파일 위치: {os.path.abspath(self.output_file)}")
            self.log_update.emit(f"📊 Invoice 시트: {len(invoice_result['data']) if invoice_result['data'] else 0}개 인보이스")
            self.log_update.emit(f"📦 Packing_List 시트: {len(packing_result['data']) if packing_result['data'] else 0}개 아이템")
            
            self.finished.emit(True, self.output_file)
            
        except Exception as e:
            self.progress_update.emit("❌ 변환 중 오류 발생")
            self.log_update.emit(f"오류: {str(e)}")
            self.finished.emit(False, str(e))


class PDFParserGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.invoice_file = None
        self.packing_file = None
        self.worker = None
        
        self.init_ui()
        
    def init_ui(self):
        """UI 초기화"""
        self.setWindowTitle("PDF Parser - Invoice & Packing List")
        self.setGeometry(100, 100, 800, 700)
        self.center_window()
        
        # 중앙 위젯
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 메인 레이아웃
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # 제목
        title_label = QLabel("PDF Parser")
        title_font = QFont("Arial", 20, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #2c3e50; margin-bottom: 10px;")
        layout.addWidget(title_label)
        
        subtitle_label = QLabel("Invoice & Packing List → Excel 변환")
        subtitle_font = QFont("Arial", 12)
        subtitle_label.setFont(subtitle_font)
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_label.setStyleSheet("color: #7f8c8d; margin-bottom: 20px;")
        layout.addWidget(subtitle_label)
        
        # 파일 선택 그룹
        file_group = QGroupBox("파일 선택")
        file_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)
        file_layout = QGridLayout(file_group)
        
        # 인보이스 파일
        file_layout.addWidget(QLabel("인보이스 파일 (*CI.pdf):"), 0, 0)
        self.invoice_edit = DragDropLineEdit("인보이스 파일을 드래그하거나 버튼으로 선택하세요...")
        self.invoice_edit.file_dropped.connect(self.on_invoice_file_dropped)
        file_layout.addWidget(self.invoice_edit, 0, 1)
        
        invoice_btn = QPushButton("파일 선택")
        invoice_btn.clicked.connect(self.select_invoice_file)
        invoice_btn.setStyleSheet(self.get_button_style())
        file_layout.addWidget(invoice_btn, 0, 2)
        
        # 패킹리스트 파일
        file_layout.addWidget(QLabel("패킹리스트 파일 (*PL.pdf):"), 1, 0)
        self.packing_edit = DragDropLineEdit("패킹리스트 파일을 드래그하거나 버튼으로 선택하세요...")
        self.packing_edit.file_dropped.connect(self.on_packing_file_dropped)
        file_layout.addWidget(self.packing_edit, 1, 1)
        
        packing_btn = QPushButton("파일 선택")
        packing_btn.clicked.connect(self.select_packing_file)
        packing_btn.setStyleSheet(self.get_button_style())
        file_layout.addWidget(packing_btn, 1, 2)
        
        # 출력 파일
        file_layout.addWidget(QLabel("출력 Excel 파일:"), 2, 0)
        self.output_edit = QLineEdit()
        self.output_edit.setText(".xlsx")
        self.output_edit.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                background-color: white;
                font-size: 11px;
            }
            QLineEdit:focus {
                border-color: #3498db;
            }
        """)
        file_layout.addWidget(self.output_edit, 2, 1)
        
        output_btn = QPushButton("저장 위치")
        output_btn.clicked.connect(self.select_output_file)
        output_btn.setStyleSheet(self.get_button_style())
        file_layout.addWidget(output_btn, 2, 2)
        
        layout.addWidget(file_group)
        
        # 변환 버튼
        self.convert_btn = QPushButton("📄 Excel로 변환")
        self.convert_btn.clicked.connect(self.start_conversion)
        self.convert_btn.setStyleSheet(self.get_convert_button_style())
        self.convert_btn.setMinimumHeight(50)
        layout.addWidget(self.convert_btn)
        
        # 진행 상태
        self.progress_label = QLabel("파일을 선택하고 변환 버튼을 클릭하세요.")
        self.progress_label.setStyleSheet("color: #34495e; font-size: 12px;")
        layout.addWidget(self.progress_label)
        
        # 진행률 바
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # 결과 로그
        log_group = QGroupBox("변환 결과")
        log_group.setStyleSheet(file_group.styleSheet())
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(200)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #2c3e50;
                color: #ecf0f1;
                border: 1px solid #34495e;
                border-radius: 4px;
                font-family: 'Courier New', monospace;
                font-size: 12px;
                padding: 8px;
            }
        """)
        log_layout.addWidget(self.log_text)
        
        layout.addWidget(log_group)
        
        # 상태바
        self.statusBar().showMessage("준비됨")
        
    def center_window(self):
        """윈도우를 화면 중앙에 배치"""
        screen = QApplication.desktop().screenGeometry()
        size = self.geometry()
        x = (screen.width() - size.width()) // 2
        y = (screen.height() - size.height()) // 2
        self.move(x, y)
        
    def get_button_style(self):
        """일반 버튼 스타일"""
        return """
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #21618c;
            }
        """
        
    def get_convert_button_style(self):
        """변환 버튼 스타일"""
        return """
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                padding: 15px;
                border-radius: 8px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
            QPushButton:pressed {
                background-color: #1e8449;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """
        
    def select_invoice_file(self):
        """인보이스 파일 선택"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "인보이스 파일 선택", "", "PDF files (*.pdf);;All files (*.*)"
        )
        if file_path:
            self.invoice_file = file_path
            self.invoice_edit.setText(os.path.basename(file_path))
            self.update_output_filename()
            
    def on_invoice_file_dropped(self, file_path):
        """인보이스 파일 드래그 앤 드롭 핸들러"""
        self.invoice_file = file_path
        self.update_output_filename()
        self.add_log(f"📁 인보이스 파일이 추가되었습니다: {os.path.basename(file_path)}")
            
    def select_packing_file(self):
        """패킹리스트 파일 선택"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "패킹리스트 파일 선택", "", "PDF files (*.pdf);;All files (*.*)"
        )
        if file_path:
            self.packing_file = file_path
            self.packing_edit.setText(os.path.basename(file_path))
            self.update_output_filename()
            
    def on_packing_file_dropped(self, file_path):
        """패킹리스트 파일 드래그 앤 드롭 핸들러"""
        self.packing_file = file_path
        self.update_output_filename()
        self.add_log(f"📦 패킹리스트 파일이 추가되었습니다: {os.path.basename(file_path)}")
            
    def select_output_file(self):
        """출력 파일 저장 위치 선택"""
        # 기본 디렉토리를 인보이스 파일 위치로 설정
        default_dir = ""
        if self.invoice_file:
            default_dir = os.path.dirname(self.invoice_file)
        elif self.packing_file:
            default_dir = os.path.dirname(self.packing_file)
        
        # 기본 파일명 설정
        default_filename = self.output_edit.text()
        if default_dir and default_filename:
            default_path = os.path.join(default_dir, default_filename)
        else:
            default_path = default_filename
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Excel 파일 저장 위치", default_path, "Excel files (*.xlsx);;All files (*.*)"
        )
        if file_path:
            self.output_edit.setText(file_path)
            
    def update_output_filename(self):
        """선택된 파일을 기반으로 출력 파일명 자동 생성"""
        if self.invoice_file or self.packing_file:
            base_name = ""
            source_dir = ""
            
            if self.invoice_file:
                base_name = Path(self.invoice_file).stem.replace(" CI", "")
                source_dir = os.path.dirname(self.invoice_file)
            elif self.packing_file:
                base_name = Path(self.packing_file).stem.replace(" PL", "")
                source_dir = os.path.dirname(self.packing_file)
            
            if base_name and source_dir:
                # 윈도우 호환 파일명 생성 (특수문자 제거)
                safe_base_name = self.make_safe_filename(base_name)
                output_path = os.path.join(source_dir, f"{safe_base_name}.xlsx")
                # 경로 정규화 (윈도우 백슬래시 처리)
                output_path = os.path.normpath(output_path)
                self.output_edit.setText(output_path)
                self.add_log(f"💾 출력 파일 경로가 설정되었습니다: {output_path}")
                
    def make_safe_filename(self, filename):
        """윈도우 호환 안전한 파일명 생성"""
        import re
        # 윈도우에서 사용할 수 없는 문자들 제거
        unsafe_chars = r'[<>:"/\\|?*]'
        safe_name = re.sub(unsafe_chars, '_', filename)
        # 연속된 언더스코어 제거
        safe_name = re.sub(r'_+', '_', safe_name)
        # 앞뒤 공백과 점 제거
        safe_name = safe_name.strip(' .')
        # 빈 문자열이면 기본값 사용
        if not safe_name:
            safe_name = "parsed_data"
        return safe_name
                
    def start_conversion(self):
        """변환 작업 시작"""
        if not self.invoice_file and not self.packing_file:
            QMessageBox.warning(self, "파일 선택", "최소 하나의 PDF 파일을 선택해주세요.")
            return
            
        if not self.output_edit.text().strip():
            QMessageBox.warning(self, "출력 파일", "출력 파일명을 입력해주세요.")
            return
            
        # UI 비활성화
        self.convert_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # 무한 진행률
        self.log_text.clear()
        
        # 출력 파일 경로 정규화 (윈도우 호환성)
        output_path = os.path.normpath(self.output_edit.text())
        
        # 워커 스레드 시작
        self.worker = ConversionWorker(
            self.invoice_file, 
            self.packing_file, 
            output_path
        )
        self.worker.progress_update.connect(self.update_progress)
        self.worker.log_update.connect(self.add_log)
        self.worker.finished.connect(self.conversion_finished)
        self.worker.start()
        
    def update_progress(self, message):
        """진행 상태 업데이트"""
        self.progress_label.setText(message)
        self.statusBar().showMessage(message)
        
    def add_log(self, message):
        """로그 메시지 추가 (색상 포함)"""
        # HTML 형식으로 색상 적용
        if message.startswith("✅"):
            colored_message = f'<span style="color: #27ae60; font-weight: bold;">{message}</span>'
        elif message.startswith("❌"):
            colored_message = f'<span style="color: #e74c3c; font-weight: bold;">{message}</span>'
        elif message.startswith("🎉"):
            colored_message = f'<span style="color: #f39c12; font-weight: bold;">{message}</span>'
        elif message.startswith("📁") or message.startswith("📊") or message.startswith("📦"):
            colored_message = f'<span style="color: #3498db;">{message}</span>'
        elif "오류:" in message:
            colored_message = f'<span style="color: #e74c3c;">{message}</span>'
        else:
            colored_message = f'<span style="color: #ecf0f1;">{message}</span>'
        
        self.log_text.append(colored_message)
        
    def conversion_finished(self, success, message):
        """변환 완료 처리"""
        # UI 다시 활성화
        self.convert_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if success:
            self.statusBar().showMessage("변환 완료!")
            
            # 완료 대화상자
            reply = QMessageBox.question(
                self, "변환 완료", 
                f"Excel 파일이 성공적으로 생성되었습니다!\n\n{os.path.basename(message)}\n\n파일을 열어보시겠습니까?",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                try:
                    if sys.platform == "win32":
                        # 윈도우에서 안전한 파일 열기
                        import subprocess
                        # 경로를 정규화
                        normalized_path = os.path.normpath(message)
                        
                        # 방법 1: os.startfile 사용 (가장 안전)
                        try:
                            os.startfile(normalized_path)
                        except OSError:
                            # 방법 2: subprocess로 cmd 사용
                            subprocess.run(['cmd', '/c', 'start', '""', f'"{normalized_path}"'], check=True)
                    elif sys.platform == "darwin":
                        os.system(f"open '{message}'")
                    else:
                        os.system(f"xdg-open '{message}'")
                except Exception as e:
                    # 대안 방법 시도
                    try:
                        if sys.platform == "win32":
                            # 대안 1: explorer로 파일 선택
                            subprocess.run(['explorer', '/select,', os.path.normpath(message)], check=True)
                        else:
                            # 파일 탐색기에서 폴더 열기
                            folder_path = os.path.dirname(message)
                            if sys.platform == "darwin":
                                os.system(f"open '{folder_path}'")
                            else:
                                os.system(f"xdg-open '{folder_path}'")
                    except Exception as e2:
                        QMessageBox.warning(
                            self, "파일 열기 오류", 
                            f"파일을 열 수 없습니다.\n\n"
                            f"파일 위치: {message}\n\n"
                            f"수동으로 파일을 열어주세요.\n"
                            f"오류: {str(e)}"
                        )
        else:
            self.statusBar().showMessage("변환 실패")
            QMessageBox.critical(self, "변환 오류", f"변환 중 오류가 발생했습니다:\n{message}")


def main():
    """PyQt5 애플리케이션 시작"""
    app = QApplication(sys.argv)
    
    # 애플리케이션 정보 설정
    app.setApplicationName("PDF Parser")
    app.setApplicationVersion("1.0")
    app.setOrganizationName("PDF Parser")
    
    # 다크 테마 적용 (선택사항)
    # app.setStyle('Fusion')
    
    # 메인 윈도우 생성
    window = PDFParserGUI()
    window.show()
    
    # 이벤트 루프 시작
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
