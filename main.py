import sys
import re
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QTextEdit, QLineEdit, QFileDialog, QMessageBox, QListWidget, QInputDialog, QProgressBar, QListWidgetItem, QDialog, QDialogButtonBox
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import numpy as np
from paddleocr import PaddleOCR
import json
import os
import fitz
from PIL import Image

class OCRThread(QThread):
    progress_update = pyqtSignal(int)
    ocr_complete = pyqtSignal(list)

    def __init__(self, ocr, image):
        super().__init__()
        self.ocr = ocr
        self.image = image

    def run(self):
        result = self.ocr.ocr(self.image, cls=True)
        self.ocr_complete.emit(result[0] if result else [])

class EditableListWidgetItem(QListWidgetItem):
    def __init__(self, item_name, value, year):
        super().__init__(f"{item_name}: {value}")
        self.year = year
        self.item_name = item_name
        self.value = value

class FinancialStatementOCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Financial Statement OCR")
        self.setGeometry(100, 100, 1500, 800)

        self.ocr = PaddleOCR(use_angle_cls=True, lang="ch")
        
        self.file_path = None
        self.current_page = 0
        self.total_pages = 0
        self.pages = []
        self.ocr_results = []
        self.mapped_data = {}
        self.statement_type = None

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout()

        # Left side (Image preview and OCR control)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        self.file_btn = QPushButton("Select PDF/Image")
        self.file_btn.clicked.connect(self.select_file)
        left_layout.addWidget(self.file_btn)

        nav_layout = QHBoxLayout()
        self.prev_btn = QPushButton("Previous")
        self.prev_btn.clicked.connect(self.prev_page)
        self.next_btn = QPushButton("Next")
        self.next_btn.clicked.connect(self.next_page)
        self.page_label = QLabel("Page: 0/0")
        nav_layout.addWidget(self.prev_btn)
        nav_layout.addWidget(self.page_label)
        nav_layout.addWidget(self.next_btn)
        left_layout.addLayout(nav_layout)

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(self.image_label)

        self.ocr_btn = QPushButton("Perform OCR")
        self.ocr_btn.clicked.connect(self.perform_ocr)
        left_layout.addWidget(self.ocr_btn)

        self.progress_bar = QProgressBar()
        left_layout.addWidget(self.progress_bar)

        # Middle (OCR results)
        middle_widget = QWidget()
        middle_layout = QVBoxLayout(middle_widget)
        middle_layout.addWidget(QLabel("OCR Results:"))
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        middle_layout.addWidget(self.results_text)

        # Right side (Mapped data)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.addWidget(QLabel("Mapped Data:"))
        self.mapped_data_list = QListWidget()
        self.mapped_data_list.itemDoubleClicked.connect(self.edit_item)
        right_layout.addWidget(self.mapped_data_list)

        # Buttons for managing mapped data
        button_layout = QHBoxLayout()
        self.add_btn = QPushButton("Add Item")
        self.add_btn.clicked.connect(self.add_item)
        self.remove_btn = QPushButton("Remove Item")
        self.remove_btn.clicked.connect(self.remove_item)
        self.insert_space_btn = QPushButton("Insert Space")
        self.insert_space_btn.clicked.connect(self.insert_space)
        button_layout.addWidget(self.add_btn)
        button_layout.addWidget(self.remove_btn)
        button_layout.addWidget(self.insert_space_btn)
        right_layout.addLayout(button_layout)

        self.save_btn = QPushButton("Save Results")
        self.save_btn.clicked.connect(self.save_results)
        right_layout.addWidget(self.save_btn)

        # Add widgets to main layout
        main_layout.addWidget(left_widget, 2)
        main_layout.addWidget(middle_widget, 1)
        main_layout.addWidget(right_widget, 1)

        central_widget.setLayout(main_layout)

    def select_file(self):
        file_dialog = QFileDialog()
        self.file_path, _ = file_dialog.getOpenFileName(self, "Select PDF/Image", "", "PDF Files (*.pdf);;Image Files (*.png *.jpg *.jpeg)")
        if self.file_path:
            if self.file_path.lower().endswith('.pdf'):
                self.load_pdf()
            else:
                self.load_image()

    def load_pdf(self):
        try:
            pdf_document = fitz.open(self.file_path)
            self.total_pages = len(pdf_document)
            self.pages = [pdf_document[i].get_pixmap() for i in range(self.total_pages)]
            self.current_page = 0
            self.update_page_display()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load PDF: {str(e)}")

    def load_image(self):
        try:
            image = Image.open(self.file_path)
            self.pages = [image]
            self.total_pages = 1
            self.current_page = 0
            self.update_page_display()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load image: {str(e)}")

    def update_page_display(self):
        if self.pages:
            if isinstance(self.pages[self.current_page], fitz.Pixmap):
                img = Image.frombytes("RGB", [self.pages[self.current_page].width, self.pages[self.current_page].height], self.pages[self.current_page].samples)
            else:
                img = self.pages[self.current_page]
            self.display_image(img)
            self.page_label.setText(f"Page: {self.current_page + 1}/{self.total_pages}")

    def display_image(self, img):
        img = img.convert("RGB")
        width, height = img.size
        aspect_ratio = width / height
        new_width = 600
        new_height = int(new_width / aspect_ratio)
        img = img.resize((new_width, new_height), Image.LANCZOS)
        qimage = QImage(img.tobytes(), img.width, img.height, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(qimage)
        self.image_label.setPixmap(pixmap)

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.update_page_display()

    def next_page(self):
        if self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.update_page_display()

    def perform_ocr(self):
        if not self.pages:
            QMessageBox.warning(self, "No File", "Please select a PDF or image file first.")
            return

        if isinstance(self.pages[self.current_page], fitz.Pixmap):
            img = Image.frombytes("RGB", [self.pages[self.current_page].width, self.pages[self.current_page].height], self.pages[self.current_page].samples)
        else:
            img = self.pages[self.current_page]
        
        img_np = np.array(img)

        self.ocr_thread = OCRThread(self.ocr, img_np)
        self.ocr_thread.progress_update.connect(self.update_progress)
        self.ocr_thread.ocr_complete.connect(self.process_ocr_results)
        self.ocr_thread.start()

        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def process_ocr_results(self, results):
        self.ocr_results = results
        self.display_ocr_results()
        self.auto_match_items_values()
        self.progress_bar.setVisible(False)

    def display_ocr_results(self):
        self.results_text.clear()
        for line in self.ocr_results:
            text = line[1][0]
            confidence = line[1][1]
            self.results_text.append(f"{text} (Confidence: {confidence:.2f})")

    def auto_match_items_values(self):
        self.mapped_data = {}
        current_item = None
        years = []
        
        # Determine statement type and extract years
        for line in self.ocr_results:
            text = line[1][0]
            if "资产负债表" in text:
                self.statement_type = "资产负债表"
                break
            elif "利润表" in text:
                self.statement_type = "利润表"
                break

        # 연도 추출
        for line in self.ocr_results:
            text = line[1][0]
            year_match = re.search(r'\b(\d{4})年度?\b', text)
            if year_match:
                year = year_match.group(1)
                if year not in years:
                    years.append(year)
                    self.mapped_data[year] = {}

        if not years:
            QMessageBox.warning(self, "Error", "No years detected in the document.")
            return

        years.sort(reverse=True)  # 최신 연도가 먼저 오도록 정렬

        for line in self.ocr_results:
            text = line[1][0]
            
            # 항목명 식별 (중국어 문자와 일부 특수문자로 구성된 경우)
            if re.match(r'^[\u4e00-\u9fff:：()（）\s]+$', text) and len(text) > 1:
                current_item = text.strip()
            
            # 값 식별 (숫자, 쉼표, 소수점, 괄호 포함)
            elif re.match(r'^[-]?\(?[\d,]+\.?\d*\)?$', text):
                if current_item:
                    value = text.strip('()')  # 괄호 제거
                    value = value.strip(',')
                    
                    
                    # 연도별 값 할당
                    for year in years:
                        if current_item not in self.mapped_data[year]:
                            self.mapped_data[year][current_item] = value
                            break
        
        self.update_mapped_data_view()

    def update_mapped_data_view(self):
        self.mapped_data_list.clear()
        for year in self.mapped_data.keys():
            self.mapped_data_list.addItem(f"--- {year} ---")
            for item_name, value in self.mapped_data[year].items():
                list_item = EditableListWidgetItem(item_name, value, year)
                self.mapped_data_list.addItem(list_item)
            self.mapped_data_list.addItem("")  # 빈 줄 추가

    def edit_item(self, item):
        if isinstance(item, EditableListWidgetItem):
            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Item")
            layout = QVBoxLayout()

            name_label = QLabel("Item Name:")
            name_edit = QLineEdit(item.item_name)
            layout.addWidget(name_label)
            layout.addWidget(name_edit)

            value_label = QLabel("Value:")
            value_edit = QLineEdit(item.value)
            layout.addWidget(value_label)
            layout.addWidget(value_edit)

            button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)
            layout.addWidget(button_box)

            dialog.setLayout(layout)

            if dialog.exec_() == QDialog.Accepted:
                new_name = name_edit.text()
                new_value = value_edit.text()
                
                # 기존 항목 삭제
                del self.mapped_data[item.year][item.item_name]
                
                # 새 항목 추가
                self.mapped_data[item.year][new_name] = new_value
                
                # ListWidgetItem 업데이트
                item.item_name = new_name
                item.value = new_value
                item.setText(f"{new_name}: {new_value}")
                
                self.update_mapped_data_view()

    def add_item(self):
        years = list(self.mapped_data.keys())
        year, ok = QInputDialog.getItem(self, "Select Year", "Choose year:", years, 0, False)
        if ok:
            item, ok = QInputDialog.getText(self, "Add Item", "Enter item name:")
            if ok and item:
                value, ok = QInputDialog.getText(self, "Add Value", "Enter value:")
                if ok and value:
                    self.mapped_data[year][item] = value
                    self.update_mapped_data_view()

    def remove_item(self):
        current_item = self.mapped_data_list.currentItem()
        if isinstance(current_item, EditableListWidgetItem):
            year = current_item.year
            item = current_item.item_name
            del self.mapped_data[year][item]
            self.update_mapped_data_view()

    def insert_space(self):
        current_row = self.mapped_data_list.currentRow()
        if current_row >= 0:
            self.mapped_data_list.insertItem(current_row, "")

    def save_results(self):
        if self.mapped_data:
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Results", "", "JSON Files (*.json)")
            if file_path:
                data_to_save = {
                    "statement_type": self.statement_type,
                    "data": self.mapped_data
                }
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data_to_save, f, ensure_ascii=False, indent=4)
                QMessageBox.information(self, "Save Successful", f"Results saved to {file_path}")
        else:
            QMessageBox.warning(self, "No Data", "No data to save. Please perform OCR first.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FinancialStatementOCRApp()
    window.show()
    sys.exit(app.exec_())