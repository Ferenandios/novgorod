
import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QLabel, QVBoxLayout, QWidget
from PyQt5.QtCore import Qt
import convert_excel_to_JSON


class DragDropWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyQt5 Drag and Drop File Window")
        self.setGeometry(100, 100, 400, 300)
        
        # Central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        # Label to indicate the drag and drop area
        self.label = QLabel("Drag and drop files here", self)
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("border: 2px dashed #aaa; padding: 20px;")
        layout.addWidget(self.label)

        
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):

        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.label.setText("Drop files here")
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
 
        self.label.setText("Drag and drop files here")

    def dropEvent(self, event):
 
        files = []
        for url in event.mimeData().urls():
         
            file_path = url.toLocalFile()
            files.append(file_path)
            convert_excel_to_JSON(file_path)
     
        if files:
            self.label.setText(f"Dropped {len(files)} file(s). First: {files[0]}")
            
        event.acceptProposedAction()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ui = DragDropWindow()
    ui.show()
    sys.exit(app.exec_())