import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, 
                             QPushButton, QTextEdit, QMessageBox, QScrollArea, QFileDialog, QDialog)
from PyQt5.QtCore import QThread, pyqtSignal
from auto_print import AutoPrint
from merge_pdfs import PDFMerger
from check_word import WordChecker
from PyPDF2 import PdfMerger

def sem(semester):
    sem_dict = {1: "a", 2: "b", 3: "c", 4: "d", 5: "e", 6: "f"}
    return sem_dict.get(semester, None)

class FetchThread(QThread):
    update_result = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, urls, save_folder):
        super().__init__()
        self.urls = urls
        self.save_folder = save_folder
        self._running = True

    def run(self):
        auto_print = AutoPrint(self.save_folder)
        word_checker = WordChecker()
        n = 1
        results = ""

        try:
            for url in self.urls:
                if not self._running:
                    break
                invalid = word_checker.check_word_in_webpage(url, 'Invalid Register number / No Results found')
                if invalid:
                    results += f"Skipped: {url} - 'Invalid Register number / No Results found'\n"
                else:
                    auto_print.auto_print(url, n)
                    results += f"Printed: {url}\n"
                    n += 1
                self.update_result.emit(results)
        except Exception as e:
            self.error_occurred.emit(f"An unexpected error occurred: {e}")
        finally:
            self.finished.emit()

    def stop(self):
        self._running = False

class ResultApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Result Fetcher")
        self.setGeometry(100, 100, 800, 600)
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        self.layout = QVBoxLayout(self.central_widget)
        
        self.input_layout = QHBoxLayout()
        
        self.register_label = QLabel("Register No:")
        self.register_input = QLineEdit()
        
        self.semester_label = QLabel("Semester:")
        self.semester_input = QLineEdit()
        
        self.students_label = QLabel("Number of Students:")
        self.students_input = QLineEdit()
        
        self.input_layout.addWidget(self.register_label)
        self.input_layout.addWidget(self.register_input)
        self.input_layout.addWidget(self.semester_label)
        self.input_layout.addWidget(self.semester_input)
        self.input_layout.addWidget(self.students_label)
        self.input_layout.addWidget(self.students_input)
        self.layout.addLayout(self.input_layout)
        
        self.button_layout = QHBoxLayout()
        
        self.fetch_button = QPushButton("Fetch Results")
        self.fetch_button.clicked.connect(self.fetch_results)
        self.button_layout.addWidget(self.fetch_button)
        
        self.clear_button = QPushButton("Clear")
        self.clear_button.clicked.connect(self.clear_all)
        self.button_layout.addWidget(self.clear_button)
        
        self.stop_button = QPushButton("Stop")
        self.stop_button.clicked.connect(self.stop_program)
        self.button_layout.addWidget(self.stop_button)

        self.merge_button = QPushButton("Merge PDFs")
        self.merge_button.clicked.connect(self.open_merge_window)
        self.button_layout.addWidget(self.merge_button)
        
        self.layout.addLayout(self.button_layout)
        
        self.result_area = QTextEdit()
        self.result_area.setReadOnly(True)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.result_area)
        self.layout.addWidget(self.scroll_area)
        
        self.fetch_thread = None
        self.save_folder = ""

    def fetch_results(self):
        link = "https://exam.pondiuni.edu.in/results/result.php?r="
        R_no = self.register_input.text()
        semester = int(self.semester_input.text())
        t_student = int(self.students_input.text())
        
        c = sem(semester)
        if c is None:
            QMessageBox.critical(self, "Error", "Invalid semester")
            return
        
        last = "&e=" + c.capitalize()
        pre = R_no[:6]
        suf = int(R_no[6:])
        urls = [link + pre + str(suf + i).zfill(2) + last for i in range(t_student)]
        
        save_folder = QFileDialog.getExistingDirectory(self, "Select Save Directory")
        if not save_folder:
            QMessageBox.critical(self, "Error", "No save folder selected.")
            return

        self.save_folder = save_folder
        
        self.fetch_thread = FetchThread(urls, save_folder)
        self.fetch_thread.update_result.connect(self.update_result_area)
        self.fetch_thread.error_occurred.connect(self.show_error)
        self.fetch_thread.finished.connect(self.fetch_finished)
        self.fetch_thread.start()

    def clear_all(self):
        self.register_input.clear()
        self.semester_input.clear()
        self.students_input.clear()
        self.result_area.clear()

    def stop_program(self):
        if self.fetch_thread is not None:
            self.fetch_thread.stop()
        QApplication.quit()

    def fetch_finished(self):
        self.fetch_thread = None

    def update_result_area(self, text):
        self.result_area.setPlainText(text)

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)

    def open_merge_window(self):
        if not self.save_folder:
            QMessageBox.information(self, "Info", "No folder selected, operation canceled.")
            return

        output_folder = QFileDialog.getExistingDirectory(self, "Select Folder to Save Merged PDF")
        if output_folder:
            output_file_path = os.path.join(output_folder, "merged_result.pdf")
            try:
                pdf_merger = PdfMerger()
                for root, _, files in os.walk(self.save_folder):
                    for file in sorted(files):
                        if file.endswith('.pdf'):
                            pdf_merger.append(os.path.join(root, file))

                with open(output_file_path, 'wb') as output_file:
                    pdf_merger.write(output_file)
                
                pdf_merger.close()
                QMessageBox.information(self, "Merged", f"PDFs merged successfully to {output_file_path}.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred while merging PDFs: {e}")
        else:
            QMessageBox.information(self, "Info", "PDF merging canceled.")
            

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ResultApp()
    window.show()
    sys.exit(app.exec_())
