import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from widgets.table import get_widget_table
from components.get_JSON_data import get_JSON_data


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("JSON Table Viewer")
        self.setGeometry(100, 100, 1000, 600)
        
        # Get JSON data
        json_data = get_JSON_data()
        
        # Create the table widget using the function
        widget_table = get_widget_table(json_data)
        self.setCentralWidget(widget_table)


def main():
    app = QApplication(sys.argv)
    
    # Set application-wide styles
    app.setStyleSheet("""
        QTableView {
            gridline-color: #d0d0d0;
            background-color: white;
            alternate-background-color: #f6f6f6;
        }
        QHeaderView::section {
            background-color: #e0e0e0;
            padding: 4px;
            border: 1px solid #c0c0c0;
            font-weight: bold;
        }
    """)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()