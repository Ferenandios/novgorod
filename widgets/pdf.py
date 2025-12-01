import os
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton
from components.convert_json_to_pdf import convert_json_to_pdf

class PdfExportWidget(QWidget):
    def __init__(self, json_data=None, auto_save=True):
        super().__init__()
        self.json_data = json_data
        self.auto_save = auto_save  # Whether to auto-save as data.pdf
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(20)
        
        # Title
        title_label = QLabel("PDF Export")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #333;
            }
        """)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # Description
        desc_text = "Export your JSON data as a PDF table"
        if self.auto_save:
            desc_text += "\n(Automatically saves as data.pdf in current directory)"
        
        desc_label = QLabel(desc_text)
        desc_label.setStyleSheet("""
            QLabel {
                font-size: 12px;
                color: #666;
                padding: 5px;
            }
        """)
        desc_label.setAlignment(Qt.AlignCenter)
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)
        
        # Button container
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.setContentsMargins(50, 0, 50, 0)
        
        # PDF Export button
        self.pdf_button = QPushButton("📄 Export to PDF")
        self.pdf_button.setMinimumHeight(50)
        self.pdf_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 4px;
                padding: 15px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
        """)
        self.pdf_button.clicked.connect(self.on_export_pdf)
        
        button_layout.addWidget(self.pdf_button)
        layout.addWidget(button_container)
        
        # Status label
        self.status_label = QLabel("Ready to export")
        self.status_label.setStyleSheet("""
            QLabel {
                font-size: 11px;
                color: #777;
                font-style: italic;
                padding: 5px;
            }
        """)
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        
        # File location info
        self.location_label = QLabel(f"Will save to: {os.path.abspath('data.pdf')}")
        self.location_label.setStyleSheet("""
            QLabel {
                font-size: 10px;
                color: #555;
                padding: 3px;
                background-color: #f5f5f5;
                border-radius: 3px;
            }
        """)
        self.location_label.setAlignment(Qt.AlignCenter)
        self.location_label.setWordWrap(True)
        layout.addWidget(self.location_label)
        
        # Add stretch to push content to center
        layout.addStretch()
        
        self.setLayout(layout)
    
    def on_export_pdf(self):
        """Handle PDF export button click"""
        self.status_label.setText("Exporting to PDF...")
        
        # Convert to PDF with auto-save
        result = convert_json_to_pdf(self.json_data, auto_save=self.auto_save)
        
        if result:
            filename = os.path.basename(result)
            self.status_label.setText(f"PDF created: {filename}")
            self.location_label.setText(f"Saved to: {result}")
        else:
            self.status_label.setText("Export failed")
    
    def update_data(self, json_data):
        """Update the JSON data for export"""
        self.json_data = json_data
        self.status_label.setText("Ready to export with updated data")

# Convenience function
def get_pdf_export_widget(json_data=None, auto_save=True):
    """
    Create and return a PdfExportWidget instance
    
    Args:
        json_data: JSON data to export (optional)
        auto_save: If True, automatically saves as data.pdf
        
    Returns:
        PdfExportWidget: Configured PDF export widget
    """
    return PdfExportWidget(json_data, auto_save)