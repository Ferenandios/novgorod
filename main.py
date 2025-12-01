import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QTabWidget
from widgets.table import get_widget_table
from widgets.pdf import get_pdf_export_widget
from components.get_JSON_data import get_JSON_data
from components.DragAndDrop import DragDropWindow


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("JSON Table Viewer with Drag & Drop")
        self.setGeometry(100, 100, 1200, 700)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Create tab widget
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)
        
        # Check if data.json file exists
        data_json_exists = os.path.exists('data.json')
        json_data = None
        
        # Tab 1: JSON Table (only if data.json exists)
        if data_json_exists:
            try:
                json_data = get_JSON_data()
                if json_data:  # Check if we actually got valid data
                    widget_table = get_widget_table(json_data)
                    tab_widget.addTab(widget_table, "JSON Table")
                    self.widget_table = widget_table
                else:
                    data_json_exists = False  # File exists but no valid data
            except Exception as e:
                print(f"Error loading JSON data: {e}")
                data_json_exists = False
        
        # Tab 2: Drag & Drop (always available)
        drag_drop_widget = DragDropWindow()
        tab_widget.addTab(drag_drop_widget, "Drag & Drop Files")
        self.drag_drop_widget = drag_drop_widget
        
        # Tab 3: PDF Export (only if we have JSON data)
        if json_data:
            # Auto-save as data.pdf in current directory
            pdf_export_widget = get_pdf_export_widget(json_data, auto_save=True)
            tab_widget.addTab(pdf_export_widget, "Export to PDF")
            self.pdf_export_widget = pdf_export_widget
        
        # If no data.json, set Drag & Drop as the first/only tab
        if not data_json_exists:
            tab_widget.setCurrentIndex(0)  # Focus on Drag & Drop tab


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
        QTabWidget::pane {
            border: 1px solid #c0c0c0;
        }
        QTabWidget::tab-bar {
            alignment: center;
        }
        QTabBar::tab {
            background-color: #f0f0f0;
            padding: 8px 16px;
            margin: 2px;
            border: 1px solid #c0c0c0;
            border-radius: 4px;
        }
        QTabBar::tab:selected {
            background-color: #e0e0e0;
            font-weight: bold;
        }
    """)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()