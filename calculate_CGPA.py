import os
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import warnings
from main_form import MainForm


# Suppress warnings for sheet names longer than 31 characters
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

class CGPAApp(QWidget):
    def __init__(self):
        super().__init__()
        self.input_directory = None
        self.output_directory = None
        self.output_combined_file = None

        # Set up the main UI
        self.initUI()

    def initUI(self):
        self.setWindowTitle('CGPA Calculator')
        self.setGeometry(100, 100, 600, 400)
        
        # Create the main layout
        layout = QVBoxLayout()

        # Create a form layout for the input and output directories
        form_layout = QFormLayout()

        # Input directory label and text field
        self.input_dir_textbox = QLineEdit(self)
        input_dir_button = QPushButton("Browse", self)
        input_dir_button.clicked.connect(self.select_input_folder)
        form_layout.addRow(QLabel("Input Directory:"), self.input_dir_textbox)
        form_layout.addRow(QLabel(""), input_dir_button)

        # Output directory label and text field
        self.output_dir_textbox = QLineEdit(self)
        output_dir_button = QPushButton("Browse", self)
        output_dir_button.clicked.connect(self.select_output_folder)
        form_layout.addRow(QLabel("Output Directory:"), self.output_dir_textbox)
        form_layout.addRow(QLabel(""), output_dir_button)

        layout.addLayout(form_layout)

        # Horizontal layout for buttons
        button_layout = QHBoxLayout()
        run_button = QPushButton("Run CGPA Calculation", self)
        run_button.clicked.connect(self.run)
        exit_button = QPushButton("Exit", self)
        exit_button.clicked.connect(self.close_application)

        button_layout.addWidget(run_button)
        button_layout.addWidget(exit_button)

        layout.addLayout(button_layout)

        # Log text area to display messages
        self.log_textbox = QTextEdit(self)
        self.log_textbox.setReadOnly(True)
        layout.addWidget(self.log_textbox)

        # Set the main layout for the widget
        self.setLayout(layout)

    def select_input_folder(self):
        self.input_directory = QFileDialog.getExistingDirectory(self, "Select Folder Containing Excel Files")
        self.input_dir_textbox.setText(self.input_directory)

    def select_output_folder(self):
        self.output_directory = QFileDialog.getExistingDirectory(self, "Select Folder for Output Files")
        self.output_dir_textbox.setText(self.output_directory)

    def run(self):
        if self.input_directory and self.output_directory:
            self.output_combined_file = os.path.join(self.output_directory, "combined_excel_output_with_CGPA.xlsx")
            self.combine_excel_sheets(self.input_directory, self.output_combined_file)
        else:
            QMessageBox.warning(self, "Missing Folders", "Please select both input and output folders.")
            self.log_message("No folder selected for input or output.")

    def log_message(self, message):
        """ Helper function to log messages to the log textbox. """
        self.log_textbox.append(message)

    def combine_excel_sheets(self, input_directory, output_combined_file):
        combined_df = pd.DataFrame()  # Create an empty DataFrame to hold combined data

        # Grade categories for summary
        grade_categories = ['Fail', 'O', 'A', 'A+', 'B', 'B+', 'C', 'P']

        for file_name in os.listdir(input_directory):
            if file_name.endswith('.xlsx') and not file_name.startswith('~$'):
                file_path = os.path.join(input_directory, file_name)

                try:
                    df = pd.read_excel(file_path)
                    df.columns = df.columns.str.replace('\n', ' ', regex=True)

                    combined_df = pd.concat([combined_df, df], ignore_index=True)
                    self.log_message(f"Added data from {file_name}")

                except PermissionError:
                    self.log_message(f"Permission denied for file: {file_name}. Please close it if it's open.")
                except Exception as e:
                    self.log_message(f"Failed to read {file_name}. Error: {e}")

        # Check for required columns before merging
        required_columns = ['Reg No.', 'Student Name', 'Credit Hours', 'Credit Point']
        if all(column in combined_df.columns for column in required_columns):
            # Merge to create a detailed DataFrame with all details
            merged_df = combined_df.groupby('Reg No.').agg(
                Student_Name=('Student Name', 'first'),  # Get first student name (assuming it's the same)
                Total_Credit_Hours=('Credit Hours', 'sum'),
                Total_Credit_Point=('Credit Point', 'sum')
            ).reset_index()

            # Calculate CGPA
            merged_df['CGPA'] = merged_df['Total_Credit_Point'] / merged_df['Total_Credit_Hours']

            # Save the merged DataFrame
            merged_df.to_excel(output_combined_file, index=False)
            self.log_message(f"Combined data saved into {output_combined_file}")

            # Display unique student details
            unique_students = merged_df[['Reg No.', 'Student_Name']]
            self.log_message("\nStudent Details (Reg No. and Name):")
            self.log_message(unique_students.to_string(index=False))

        else:
            self.log_message("One or more required columns are missing from the combined data.")
            missing_columns = set(required_columns) - set(combined_df.columns)
            self.log_message(f"Missing columns: {missing_columns}")
            
   

    def close_application(self):
        self.close()


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = CGPAApp()
    window.show()
    sys.exit(app.exec_())
