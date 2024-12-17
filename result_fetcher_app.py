import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTextEdit, QMessageBox, QFileDialog)
from PyQt5.QtCore import Qt
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber

class ResultFetcherApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Result Fetcher")
        self.setGeometry(100, 100, 800, 600)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)

        self.button_layout = QHBoxLayout()
        self.button_layout.setAlignment(Qt.AlignCenter)
        self.button_layout.setSpacing(20)  # Space between buttons

        # Upload Button
        self.upload_button = QPushButton("Upload PDFs")
        self.upload_button.setFixedSize(200, 50)  # Set button size (width, height)
        self.upload_button.clicked.connect(self.upload_pdfs)
        self.button_layout.addWidget(self.upload_button)

        # Arrear Button
        self.arrear_button = QPushButton("Collect Arrear Pages")
        self.arrear_button.setFixedSize(200, 50)
        self.arrear_button.clicked.connect(self.collect_arrear_pages)
        self.arrear_button.setEnabled(False)
        self.button_layout.addWidget(self.arrear_button)

        # All Clear Button
        self.all_clear_button = QPushButton("Collect All Clear Pages")
        self.all_clear_button.setFixedSize(200, 50)
        self.all_clear_button.clicked.connect(self.collect_all_clear_pages)
        self.all_clear_button.setEnabled(False)
        self.button_layout.addWidget(self.all_clear_button)

        # Merge Button
        self.merge_button = QPushButton("Merge Selected Pages")
        self.merge_button.setFixedSize(200, 50)
        self.merge_button.clicked.connect(self.merge_selected_pages)
        self.merge_button.setEnabled(False)
        self.button_layout.addWidget(self.merge_button)

        # Clear Button
        self.clear_button = QPushButton("Clear")
        self.clear_button.setFixedSize(200, 50)
        self.clear_button.clicked.connect(self.clear_all)
        self.button_layout.addWidget(self.clear_button)

        # Stop Button
        self.stop_button = QPushButton("Stop")
        self.stop_button.setFixedSize(200, 50)
        self.stop_button.clicked.connect(self.stop_program)
        self.button_layout.addWidget(self.stop_button)

        self.layout.addLayout(self.button_layout)

        # Result Area
        self.result_area = QTextEdit()
        self.result_area.setReadOnly(True)
        self.layout.addWidget(self.result_area)

        self.uploaded_files = []
        self.pages_to_merge = []

    def upload_pdfs(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select PDF Files", "", "PDF Files (*.pdf)")
        if files:
            self.uploaded_files = files
            self.result_area.append(f"Uploaded {len(files)} PDFs.")
            self.arrear_button.setEnabled(True)
            self.all_clear_button.setEnabled(True)
        else:
            QMessageBox.information(self, "Info", "No files selected.")

    def collect_pages(self, check_fail):
        if not self.uploaded_files:
            QMessageBox.warning(self, "Warning", "No PDFs uploaded.")
            return

        self.pages_to_merge.clear()  # Clear any previous selections

        for file in self.uploaded_files:
            try:
                with pdfplumber.open(file) as pdf:
                    for page_number, page in enumerate(pdf.pages):
                        text = page.extract_text()

                        # Ensure text is not None and contains data
                        if text:
                            contains_fail = "Fail" in text

                            if (check_fail and contains_fail) or (not check_fail and not contains_fail):
                                self.pages_to_merge.append((file, page_number))
                self.result_area.append(f"Collected pages from: {os.path.basename(file)}")

            except Exception as e:
                self.result_area.append(f"Error processing {os.path.basename(file)}: {e}")

        if self.pages_to_merge:
            self.result_area.append(f"{len(self.pages_to_merge)} pages collected.")
            self.merge_button.setEnabled(True)
        else:
            self.result_area.append("No relevant pages collected.")

    def collect_arrear_pages(self):
        self.collect_pages(check_fail=True)

    def collect_all_clear_pages(self):
        self.collect_pages(check_fail=False)

    def merge_selected_pages(self):
        if not self.pages_to_merge:
            QMessageBox.warning(self, "Warning", "No pages to merge.")
            return

        output_folder = QFileDialog.getExistingDirectory(self, "Select Folder to Save Merged PDF")
        if not output_folder:
            QMessageBox.information(self, "Info", "Operation canceled.")
            return

        output_writer = PdfWriter()

        try:
            for file, page_number in self.pages_to_merge:
                reader = PdfReader(file)
                output_writer.add_page(reader.pages[page_number])

            output_path = os.path.join(output_folder, "merged_result.pdf")
            with open(output_path, "wb") as output_file:
                output_writer.write(output_file)

            self.result_area.append(f"Merged {len(self.pages_to_merge)} pages into: {output_path}")
            self.pages_to_merge.clear()
            self.merge_button.setEnabled(False)

        except Exception as e:
            self.result_area.append(f"Error merging pages: {e}")

    def clear_all(self):
        self.uploaded_files = []
        self.pages_to_merge.clear()
        self.result_area.clear()
        self.arrear_button.setEnabled(False)
        self.all_clear_button.setEnabled(False)
        self.merge_button.setEnabled(False)

    def stop_program(self):
        QApplication.quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ResultFetcherApp()
    window.show()
    sys.exit(app.exec_())
