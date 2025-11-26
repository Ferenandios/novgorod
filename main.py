from components.convert_excel_to_JSON import convert_excel_to_JSON
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import pyqtSlot

def window():
    # Initialize PyQT enviroment
    app = QApplication(sys.argv)
    widget = QWidget()

    # Create label
    textLabel = QLabel(widget)
    textLabel.setText("works!")
    textLabel.move(8, 8)

    # Set global font
    font = QFont('Arial', 16)
    textLabel.setFont(font)

    # Set default values for spawn the window
    widget.setGeometry(960, 540, 320, 200)
    widget.setWindowTitle("Excel Extracter")
    widget.show()

    widget.setStyleSheet("QWidget {background: #90EE90;}")

    sys.exit(app.exec_())

if __name__ == '__main__':
    window()

convert_excel_to_JSON('../Elenents.xlsx')