import sys
import os
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from button import *  # Ensure no circular imports

class Overalldata(QMainWindow):
    
    def __init__(self):
        super().__init__()

        # Initialize the button helper
        self.bb = Button()
        
        self.setWindowTitle("EXAM RESULT AUTOMATION ENGINE")
        self.setGeometry(100, 100, 600, 400)

        # Set up the central widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # Main layout to stack GIF and buttons
        self.main_layout = QVBoxLayout(self.central_widget)

        # Set up the GIF background with a relative path
        gif_path = os.path.join(os.getcwd(), 'bg.gif')
        self.gif_label = QLabel(self)
        self.gif_movie = QMovie(gif_path)
        self.gif_label.setMovie(self.gif_movie)
        self.gif_movie.start()

        # Add the GIF as a background (not part of the layout)
        self.gif_label.setGeometry(self.rect())
        self.gif_label.setScaledContents(True)
        self.gif_label.lower()

        # Create a spacer to push buttons to the center vertically
        self.main_layout.addStretch(1)

        # Create a horizontal layout for buttons
        self.button_layout = QHBoxLayout()

        # Set up buttons
        self.extract_excel_button = QPushButton("Extract Excel", self.central_widget)
        self.extract_excel_button.clicked.connect(self.open_extract_excel)
        self.bb.b(self.extract_excel_button)
        self.button_layout.addWidget(self.extract_excel_button)

        self.combine_excel_button = QPushButton("Combine Excel", self.central_widget)
        self.combine_excel_button.clicked.connect(self.open_combine_excel)
        self.bb.b(self.combine_excel_button)
        self.button_layout.addWidget(self.combine_excel_button)

        self.calculate_cgpa_button = QPushButton("Calculate CGPA", self.central_widget)
        self.calculate_cgpa_button.clicked.connect(self.open_calculate_cgpa)
        self.bb.b(self.calculate_cgpa_button)
        self.button_layout.addWidget(self.calculate_cgpa_button)

        # Add the Back button before the Exit button
        self.back_button = QPushButton("Back", self.central_widget)
        self.back_button.clicked.connect(self.go_back_to_main_form)
        self.bb.b(self.back_button)
        self.button_layout.addWidget(self.back_button)

        # Add the Exit button
        self.exit_button = QPushButton("Exit", self.central_widget)
        self.exit_button.clicked.connect(self.close_application)
        self.bb.b(self.exit_button)
        self.button_layout.addWidget(self.exit_button)

        # Add the button layout to the main layout in the center
        self.main_layout.addLayout(self.button_layout)

        # Create a spacer to push buttons to the center vertically
        self.main_layout.addStretch(1)

    def resizeEvent(self, event):
        # Resize the GIF label to cover the entire window on resize
        self.gif_label.setGeometry(self.rect())
        super().resizeEvent(event)

    def open_extract_excel(self):
        from excel_extracter import Extract  # Move import here
        self.result_app = Extract()
        self.result_app.show()
        
    def open_combine_excel(self):
        from excel_combine import ExcelCombiner  # Move import here
        self.result_fetcher_app = ExcelCombiner()
        self.result_fetcher_app.show()
        
    def open_calculate_cgpa(self):
        from calculate_CGPA import CGPAApp  # Move import here
        self.folder_selector_app = CGPAApp()
        self.folder_selector_app.show()
        
    def close_application(self):
        self.close()

    def go_back_to_main_form(self):
        from main_form import MainForm  # Import the main form class here
        self.main_form = MainForm()
        self.main_form.showMaximized()
        self.close()  # Close the current window

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = Overalldata()
    main_window.showMaximized()
    sys.exit(app.exec_())
