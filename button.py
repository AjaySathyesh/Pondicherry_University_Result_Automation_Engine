from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QPushButton

class Button:
    def b(self, button):
        # Set up the font
        font = QFont()
        font.setPointSize(15)
        font.setBold(True)
        button.setFont(font)

        
        # Set the button size
        button.setFixedSize(300, 70)
        button.setStyleSheet("""
            QPushButton {
                border: 7px solid #04C4FF;
                padding: 10px;
                color: white;
                background-color: transparent;
                border-radius: 20px;
            }
            QPushButton:hover {
                background-color: #04C4FF;
            }
        """)
        
        # # Set the button border (style sheet)
        # button.setStyleSheet("border: 2px solid black;")
