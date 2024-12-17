import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox
from PyPDF2 import PdfMerger

class FolderSelectorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Folder Selector")
        self.setGeometry(100, 100, 500, 200)
        
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        self.path_label = QLabel("Enter the path:", self.central_widget)
        self.path_label.setGeometry(20, 20, 100, 30)
        self.path_input = QLineEdit(self.central_widget)
        self.path_input.setGeometry(130, 20, 250, 30)
        
        self.select_button = QPushButton("Select Folder", self.central_widget)
        self.select_button.setGeometry(400, 20, 80, 30)
        self.select_button.clicked.connect(self.select_folder)
        
        self.output_label = QLabel("Enter output file name:", self.central_widget)
        self.output_label.setGeometry(20, 70, 150, 30)
        self.output_input = QLineEdit(self.central_widget)
        self.output_input.setGeometry(180, 70, 250, 30)
        
        self.merge_button = QPushButton("Merge PDFs", self.central_widget)
        self.merge_button.setGeometry(180, 120, 150, 30)
        self.merge_button.clicked.connect(self.merge_pdfs)
        
    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.path_input.setText(folder_path)

    def merge_pdfs(self):
        folder_path = self.path_input.text()
        output_file_name = self.output_input.text()

        if not folder_path:
            QMessageBox.warning(self, "No Folder Selected", "Please select a folder to merge PDFs.")
            return
        
        if not output_file_name:
            QMessageBox.warning(self, "No Output File Name", "Please enter an output file name.")
            return
        
        if not output_file_name.endswith('.pdf'):
            output_file_name += '.pdf'

        try:
            pdf_merger = PdfMerger()

            # Collect and sort PDF files
            pdf_files = []
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if file.endswith('.pdf'):
                        pdf_files.append(os.path.join(root, file))
            pdf_files.sort()  # Sort the files by their names

            # Append sorted PDF files to the merger
            for pdf_file in pdf_files:
                pdf_merger.append(pdf_file)

            output_path = os.path.join(folder_path, output_file_name)
            pdf_merger.write(output_path)
            pdf_merger.close()
            QMessageBox.information(self, "Success", f"PDFs merged successfully to {output_path}.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while merging PDFs: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FolderSelectorApp()
    window.show()
    sys.exit(app.exec_())
