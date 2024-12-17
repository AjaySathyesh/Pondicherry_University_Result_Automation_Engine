import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from button import Button
from overall import Overalldata
from result_app import ResultApp
from result_fetcher_app import ResultFetcherApp
from folder_selector_app import FolderSelectorApp
import os

class Worker(QThread):
    result_ready = pyqtSignal(str)  # Signal to send results back to the GUI

    def run(self):
        # Simulate a time-consuming operation (e.g., result analysis)
        import time
        time.sleep(5)  # Simulating a long-running process
        self.result_ready.emit("Result analysis completed!")  # Emit signal when done

class MainForm(QMainWindow):
    
    def __init__(self):
        super().__init__()
        
        # Initialize the button helper
        self.bb = Button()
        
        self.setWindowTitle("EXAM RESULT AUTOMATION ENGINE")
        self.setGeometry(100, 100, 600, 400)
        
        # Set up the central widget
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # Set up the GIF background
        gif_path = os.path.join(os.getcwd(), 'bg.gif')
        self.gif_label = QLabel(self)
        self.gif_movie = QMovie(gif_path)
        self.gif_label.setMovie(self.gif_movie)
        self.gif_movie.start()
        
        # Make the GIF label cover the entire window
        self.gif_label.setGeometry(self.rect())
        self.gif_label.setScaledContents(True)
        self.gif_label.lower()
        
        # Set up buttons
        self.print_results_button = QPushButton("Print the Results", self.central_widget)
        self.print_results_button.clicked.connect(self.open_print_results)
        self.bb.b(self.print_results_button)  # Apply button styling
        
        self.filter_results_button = QPushButton("Filter the Results", self.central_widget)
        self.filter_results_button.clicked.connect(self.open_filter_results)
        self.bb.b(self.filter_results_button)  # Apply button styling
        
        self.merge_pdfs_button = QPushButton("Merging PDFs", self.central_widget)
        self.merge_pdfs_button.clicked.connect(self.open_merge_pdfs)
        self.bb.b(self.merge_pdfs_button)  # Apply button styling
        
        # Add the RESULT ANALYSIS button
        self.result_analysis_button = QPushButton("RESULT ANALYSIS", self.central_widget)
        self.result_analysis_button.clicked.connect(self.start_result_analysis)
        self.bb.b(self.result_analysis_button)  # Apply button styling
        
        self.update_button_positions()
        
    def resizeEvent(self, event):
        # Resize the GIF label to cover the entire window on resize
        self.gif_label.setGeometry(self.rect())
        self.update_button_positions()
        super().resizeEvent(event)
        
    def update_button_positions(self):
        # Get the dimensions of the central widget
        widget_width = self.central_widget.width()
        widget_height = self.central_widget.height()
        
        # Get the dimensions of the buttons
        button_width = self.print_results_button.width()
        button_height = self.print_results_button.height()
        
        # Calculate positions to center the buttons
        button_spacing = 20  # Adjusted for better spacing
        total_width = 4 * button_width + 3 * button_spacing  # Adjusted for the new button
        x_start = (widget_width - total_width) // 2
        y_position = (widget_height - button_height) // 2
        
        # Position all buttons
        self.print_results_button.move(x_start, y_position)
        self.filter_results_button.move(x_start + button_width + button_spacing, y_position)
        self.merge_pdfs_button.move(x_start + 2 * (button_width + button_spacing), y_position)
        self.result_analysis_button.move(x_start + 3 * (button_width + button_spacing), y_position)  # New button position
        
    def open_print_results(self):
        self.result_app = ResultApp()
        self.result_app.show()
        
    def open_filter_results(self):
        self.result_fetcher_app = ResultFetcherApp()
        self.result_fetcher_app.show()
        
    def open_merge_pdfs(self):
        self.folder_selector_app = FolderSelectorApp()
        self.folder_selector_app.show()
        
    def start_result_analysis(self):
        # Start the result analysis in a separate thread
        self.worker = Worker()
        self.worker.result_ready.connect(self.show_result_analysis_message)  # Connect signal
        self.worker.start()  # Start the worker thread
        
    def show_result_analysis_message(self, message):
        # Display a message box with the result of the analysis
        QMessageBox.information(self, "Analysis Result", message)
        self.result_analysis_app = Overalldata()  
        self.hide()  # Hide the current window
        self.result_analysis_app.showMaximized()  
    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainForm()
    main_window.showMaximized()
    sys.exit(app.exec_())
