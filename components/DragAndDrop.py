import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QLabel, QVBoxLayout, QWidget
from PyQt5.QtCore import Qt
from components.convert_excel_to_JSON import convert_excel_to_JSON

class DragDropWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Label to indicate the drag and drop area
        self.label = QLabel("Drag and drop Excel files here", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("""
            border: 2px dashed #aaa; 
            padding: 40px; 
            background-color: #f8f8f8;
            font-size: 14px;
            color: #666;
        """)
        layout.addWidget(self.label)

        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.label.setText("Drop Excel files here")
            self.label.setStyleSheet("""
                border: 2px dashed #4CAF50; 
                padding: 40px; 
                background-color: #f0f8f0;
                font-size: 14px;
                color: #2E7D32;
            """)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.label.setText("Drag and drop Excel files here")
        self.label.setStyleSheet("""
            border: 2px dashed #aaa; 
            padding: 40px; 
            background-color: #f8f8f8;
            font-size: 14px;
            color: #666;
        """)

    def dropEvent(self, event):
        files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            files.append(file_path)
            # Convert Excel to JSON when file is dropped
            convert_excel_to_JSON(file_path)
     
        if files:
            self.label.setText(f"Dropped {len(files)} file(s). First: {files[0]}")
            self.label.setStyleSheet("""
                border: 2px dashed #2196F3; 
                padding: 40px; 
                background-color: #e3f2fd;
                font-size: 14px;
                color: #1976D2;
            """)
            
        event.acceptProposedAction()

# Keep the original DragDropWindow for standalone use
class DragDropWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt5 Drag and Drop File Window")
        self.setGeometry(100, 100, 400, 300)
        
        # Use the widget version
        drag_drop_widget = DragDropWidget()
        self.setCentralWidget(drag_drop_widget)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ui = DragDropWindow()
    ui.show()
    sys.exit(app.exec_())