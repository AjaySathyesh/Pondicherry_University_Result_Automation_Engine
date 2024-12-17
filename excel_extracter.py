import pdfplumber
import pandas as pd
import re
import os
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
  # Import OverallWindow from overall.py

class Extract(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("PDF Table Extractor Dashboard")
        self.setGeometry(100, 100, 600, 400)

        # Set up the main layout
        layout = QVBoxLayout()

        # Input and output file paths section (dashboard-like)
        input_layout = QHBoxLayout()

        self.input_label = QLabel("No input file selected", self)
        input_layout.addWidget(self.input_label)

        self.select_input_button = QPushButton("Select PDF File", self)
        self.select_input_button.clicked.connect(self.select_input_pdf)
        input_layout.addWidget(self.select_input_button)

        layout.addLayout(input_layout)

        output_layout = QHBoxLayout()

        self.output_label = QLabel("No output directory selected", self)
        output_layout.addWidget(self.output_label)

        self.select_output_button = QPushButton("Select Output Directory", self)
        self.select_output_button.clicked.connect(self.select_output_directory)
        output_layout.addWidget(self.select_output_button)

        layout.addLayout(output_layout)

        # Add a text box for displaying logs or statuses
        self.log_box = QTextEdit(self)
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Logs will be displayed here...")
        layout.addWidget(self.log_box)

        # Buttons to trigger actions (like on a dashboard)
        button_layout = QHBoxLayout()

        self.extract_button = QPushButton("Extract Tables", self)
        self.extract_button.clicked.connect(self.run_extraction)
        button_layout.addWidget(self.extract_button)

        

        self.exit_button = QPushButton("Exit", self)
        self.exit_button.clicked.connect(self.close)
        button_layout.addWidget(self.exit_button)

        layout.addLayout(button_layout)

        # Set the layout to a container widget
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.pdf_path = None
        self.output_directory = None

    def select_input_pdf(self):
        # Open a file dialog to select the input PDF
        self.pdf_path, _ = QFileDialog.getOpenFileName(self, "Select PDF File", "", "PDF Files (*.pdf)")
        if self.pdf_path:
            self.input_label.setText(f"Input File: {self.pdf_path}")
            self.log_box.append(f"Selected PDF File: {self.pdf_path}")

    def select_output_directory(self):
        # Open a directory selection dialog to select the output directory
        self.output_directory = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if self.output_directory:
            self.output_label.setText(f"Output Directory: {self.output_directory}")
            self.log_box.append(f"Selected Output Directory: {self.output_directory}")

    def run_extraction(self):
        if not self.pdf_path or not self.output_directory:
            self.show_error_message("Please select both an input PDF and an output directory!")
            return

        self.extract_second_table_with_student_info(self.pdf_path, self.output_directory)

    def extract_second_table_with_student_info(self, pdf_path, output_directory):
        # Open the PDF file
        with pdfplumber.open(pdf_path) as pdf:
            # Loop through all the pages in the PDF
            for page_number, page in enumerate(pdf.pages):
                text = page.extract_text()  # Extract the full text on the page
                lines = text.split('\n')    # Split the text into lines
                tables = page.extract_tables()  # Extract the tables on the page

                # Initialize variables for storing extracted student details
                reg_no = None
                name = None

                # Loop through the lines to find Reg No and Name of the Student
                for line in lines:
                    # Extract Reg. No.
                    reg_no_match = re.search(r'Reg\. No\. :\s*(\S+)', line)
                    if reg_no_match:
                        reg_no = reg_no_match.group(1)

                    # Extract Name of the Student
                    name_match = re.search(r'Name of the Student :\s*(.*)', line)
                    if name_match:
                        name = name_match.group(1)

                # Check if at least two tables are available on the page
                if len(tables) >= 2:
                    # Extract the second table (index 1 refers to the second table)
                    second_table = tables[1]
                    df = pd.DataFrame(second_table[1:], columns=second_table[0])

                    # Add student info as new columns to the table in the desired order
                    df.insert(0, 'Student Name', name)  # Insert Name as the first column
                    df.insert(1, 'Reg No.', reg_no)  # Insert Reg No as the second column

                    # Create a unique filename for the output
                    output_file = os.path.join(output_directory, f"Page_{page_number+1}_Second_Table.xlsx")

                    # Save the DataFrame to an Excel file
                    df.to_excel(output_file, index=False)
                    self.log_box.append(f"Saved {output_file}")
                else:
                    self.log_box.append(f"Less than 2 tables found on page {page_number+1}. Skipping.")



    def show_error_message(self, message):
        error_dialog = QMessageBox()
        error_dialog.setIcon(QMessageBox.Warning)
        error_dialog.setText(message)
        error_dialog.setWindowTitle("Error")
        error_dialog.exec_()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    extractor = Extract()
    extractor.show()
    sys.exit(app.exec_())
