import os
import pandas as pd
import matplotlib.pyplot as plt
import warnings
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

# Suppress warnings for sheet names longer than 31 characters
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

class ExcelCombiner(QMainWindow):
    def __init__(self):
        super().__init__()
        self.input_directory = None
        self.output_combined_file = None
        self.output_individual_folder = None

        self.initUI()

    def initUI(self):
        self.setWindowTitle("Excel Combiner")
        self.setGeometry(100, 100, 800, 400)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        # Main layout
        main_layout = QVBoxLayout(central_widget)

        # Input Directory Selection
        input_layout = QHBoxLayout()
        self.select_input_button = QPushButton("Select Input Directory", self)
        self.select_input_button.clicked.connect(self.select_input_directory)
        self.input_text_box = QLineEdit(self)
        self.input_text_box.setPlaceholderText("Selected Input Directory...")
        self.input_text_box.setReadOnly(True)
        input_layout.addWidget(self.select_input_button)
        input_layout.addWidget(self.input_text_box)
        main_layout.addLayout(input_layout)

        # Output File Selection
        output_layout = QHBoxLayout()
        self.select_output_button = QPushButton("Select Output File", self)
        self.select_output_button.clicked.connect(self.select_output_file)
        self.output_text_box = QLineEdit(self)
        self.output_text_box.setPlaceholderText("Output Combined File...")
        self.output_text_box.setReadOnly(True)
        output_layout.addWidget(self.select_output_button)
        output_layout.addWidget(self.output_text_box)
        main_layout.addLayout(output_layout)

        # Individual Folder Selection
        individual_layout = QHBoxLayout()
        self.select_individual_folder_button = QPushButton("Select Individual Folder", self)
        self.select_individual_folder_button.clicked.connect(self.select_individual_folder)
        self.individual_text_box = QLineEdit(self)
        self.individual_text_box.setPlaceholderText("Selected Individual Folder...")
        self.individual_text_box.setReadOnly(True)
        individual_layout.addWidget(self.select_individual_folder_button)
        individual_layout.addWidget(self.individual_text_box)
        main_layout.addLayout(individual_layout)

        # Combine Excel Sheets Button
        self.combine_button = QPushButton("Combine Excel Sheets", self)
        self.combine_button.clicked.connect(self.combine_excel_sheets)
        main_layout.addWidget(self.combine_button, alignment=Qt.AlignCenter)

        # Status text area
        self.text_area = QTextEdit(self)
        self.text_area.setPlaceholderText("Status messages will appear here...")
        self.text_area.setReadOnly(True)
        main_layout.addWidget(self.text_area)

        # Exit Button
        self.exit_button = QPushButton("Exit", self)
        self.exit_button.clicked.connect(self.exit_application)
        main_layout.addWidget(self.exit_button, alignment=Qt.AlignRight)

    def select_input_directory(self):
        self.input_directory = QFileDialog.getExistingDirectory(self, "Select Input Directory")
        self.input_text_box.setText(self.input_directory)

    def select_output_file(self):
        self.output_combined_file, _ = QFileDialog.getSaveFileName(self, "Save Combined File As", "", "Excel Files (*.xlsx)")
        self.output_text_box.setText(self.output_combined_file)

    def select_individual_folder(self):
        self.output_individual_folder = QFileDialog.getExistingDirectory(self, "Select Output Individual Folder")
        self.individual_text_box.setText(self.output_individual_folder)

    def combine_excel_sheets(self):
        if not (self.input_directory and self.output_combined_file and self.output_individual_folder):
            QMessageBox.warning(self, "Error", "Please select all directories and files.")
            return
        
        combined_df = pd.DataFrame()

        for file_name in os.listdir(self.input_directory):
            if file_name.endswith('.xlsx') and not file_name.startswith('~$'):
                file_path = os.path.join(self.input_directory, file_name)
                try:
                    df = pd.read_excel(file_path)
                    df.columns = df.columns.str.replace('\n', ' ', regex=True)
                    combined_df = pd.concat([combined_df, df], ignore_index=True)
                    self.text_area.append(f"Added data from {file_name}")
                except Exception as e:
                    self.text_area.append(f"Failed to read {file_name}. Error: {e}")
        
        if not combined_df.empty:
            if 'Reg No.' in combined_df.columns:
                combined_df['Reg No.'] = combined_df['Reg No.'].astype(str)

            combined_df.sort_values(by='Reg No.', inplace=True)
            self.calculate_and_save_results(combined_df)
        else:
            self.text_area.append("No data to combine.")

    def calculate_and_save_results(self, combined_df):
        if 'Reg No.' in combined_df.columns:
            combined_df.sort_values(by='Reg No.', inplace=True)

        if 'Credit Hours' in combined_df.columns and 'Credit Point' in combined_df.columns:
            grouped = combined_df.groupby('Reg No.').agg(
                total_credit_hours=('Credit Hours', 'sum'),
                total_credit_points=('Credit Point', 'sum')
            ).reset_index()

            grouped['CGPA'] = grouped['total_credit_points'] / grouped['total_credit_hours']

            student_details = pd.merge(
                combined_df[['Reg No.', 'Student Name']].drop_duplicates(),
                grouped,
                on='Reg No.',
                how='inner'
            )

            result_df = pd.DataFrame()
            for reg_no, student_data in combined_df.groupby('Reg No.'):
                result_df = pd.concat([result_df, student_data])
                student_totals = grouped[grouped['Reg No.'] == reg_no]
                summary_row = pd.DataFrame({
                    'Reg No.': [reg_no],
                    'Credit Hours': [student_totals['total_credit_hours'].values[0]],
                    'Credit Point': [student_totals['total_credit_points'].values[0]],
                    'CGPA': [student_totals['CGPA'].values[0]]
                })
                result_df = pd.concat([result_df, summary_row], ignore_index=True)

            result_df.to_excel(self.output_combined_file, index=False)
            self.text_area.append(f"Combined data saved into {self.output_combined_file}")

            self.create_subject_summary(combined_df, grouped)
            self.create_topper_list(grouped, combined_df)

    def create_subject_summary(self, combined_df, grouped):
        mini_summary_list = []
        if 'Subject Name' in combined_df.columns and 'Grade' in combined_df.columns:
            for subject, subject_data in combined_df.groupby('Subject Name'):
                total_students = subject_data['Reg No.'].nunique()
                total_fails = (subject_data['Grade'] == 'Fail').sum()
                total_passes = total_students - total_fails

                passing_percentage = (total_passes / total_students) * 100 if total_students > 0 else 0

                mini_summary_list.append({
                    'Subject Name': subject,
                    'Total Passes': total_passes,
                    'Total Fails': total_fails,
                    'Passing Percentage': passing_percentage,
                })

                try:
                    self.save_subject_summary(subject, subject_data, total_passes, total_fails, self.output_individual_folder)
                except Exception as e:
                    self.text_area.append(f"Error saving summary for {subject}: {e}")

        mini_summary_df = pd.DataFrame(mini_summary_list)
        mini_summary_file_name = os.path.join(self.output_individual_folder, 'Mini_Summary.xlsx')
        mini_summary_df.to_excel(mini_summary_file_name, index=False)
        self.text_area.append(f"Mini summary saved as a separate file: {mini_summary_file_name}")
        self.plot_pass_fail_summary_graph(mini_summary_file_name)

    def create_topper_list(self, grouped, combined_df):
        topper_list = pd.merge(grouped, 
                               combined_df[['Reg No.', 'Student Name']].drop_duplicates(),
                               on='Reg No.',
                               how='inner')

        topper_list_sorted = topper_list.sort_values(by='CGPA', ascending=False)
        topper_list_sorted['Rank'] = range(1, len(topper_list_sorted) + 1)

        topper_list_top_5 = topper_list_sorted.head(5)
        topper_list_file_name = os.path.join(self.output_individual_folder, 'Topper_List.xlsx')
        topper_list_top_5 = topper_list_top_5[['Student Name', 'Reg No.', 'CGPA', 'Rank']]
        topper_list_top_5.to_excel(topper_list_file_name, index=False)
        self.text_area.append(f"Topper list saved as a separate file: {topper_list_file_name}")

    def save_subject_summary(self, subject_name, subject_data, total_passes, total_fails, user_defined_dir):
        # Ensure the user-defined directory exists
        os.makedirs(user_defined_dir, exist_ok=True)

        subject_summary_df = pd.DataFrame({
            'Reg No.': subject_data['Reg No.'],
            'Student Name': subject_data['Student Name'],
            'Grade': subject_data['Grade']
        })
        
        summary_row = pd.DataFrame({
            'Reg No.': [''],
            'Student Name': ['Subject Summary:'],
            'Grade': [f'Passes: {total_passes}, Fails: {total_fails}']
        })

        subject_summary_df = pd.concat([subject_summary_df, summary_row], ignore_index=True)

        # Save the summary file directly in the user-defined directory with the subject name
        subject_summary_file = os.path.join(user_defined_dir, f"{subject_name}.xlsx")
        
        try:
            subject_summary_df.to_excel(subject_summary_file, index=False)
            self.text_area.append(f"Summary for {subject_name} saved as a separate file: {subject_summary_file}")
        except Exception as e:
            self.text_area.append(f"Failed to save summary for {subject_name}. Error: {e}")

    def plot_pass_fail_summary_graph(self, mini_summary_file_name):
        mini_summary_df = pd.read_excel(mini_summary_file_name)

        # Check required columns in the mini summary DataFrame
        if 'Subject Name' not in mini_summary_df.columns or 'Total Passes' not in mini_summary_df.columns or 'Total Fails' not in mini_summary_df.columns:
            self.text_area.append("Required columns not found in the mini summary file.")
            return

        # Create a larger figure
        fig, ax = plt.subplots(figsize=(14, 7))  # Increased size

        # Bar width
        bar_width = 0.35
        index = range(len(mini_summary_df))

        # Plotting passing percentage and number of fails as bar graphs
        bars1 = ax.bar(index, mini_summary_df['Total Passes'], bar_width, color='skyblue', label='Total Passes', alpha=0.7)
        bars2 = ax.bar([i + bar_width for i in index], mini_summary_df['Total Fails'], bar_width, color='red', label='Total Fails', alpha=0.7)

        ax.set_xlabel('Subject Name')
        ax.set_ylabel('Count')
        ax.set_title('Total Passes and Total Fails by Subject')
        ax.set_xticks([i + bar_width / 2 for i in index])  # Set ticks in the center of grouped bars
        ax.set_xticklabels(mini_summary_df['Subject Name'], rotation=45, ha='right')  # Rotate labels

        # Adding the values on top of the bars
        for bar in bars1:
            yval = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, yval, f'{int(yval)}', ha='center', va='bottom')

        for bar in bars2:
            yval = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, yval, f'{int(yval)}', ha='center', va='bottom')

        ax.legend()

        # Use tight layout
        plt.tight_layout()

        # Save the graph as an image with high DPI
        graph_image_path = os.path.join(self.output_individual_folder, 'Pass_Fail_Summary_Graph.png')
        plt.savefig(graph_image_path, dpi=300)  # Increased DPI for better resolution
        plt.show()
        self.text_area.append(f"Graph saved as {graph_image_path}")
    def exit_application(self):
        self.close()

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    window = ExcelCombiner()
    window.show()
    sys.exit(app.exec_())
