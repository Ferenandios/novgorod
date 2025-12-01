import os
import json
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from PyQt5.QtWidgets import QPushButton, QMessageBox, QFileDialog
from datetime import datetime

def convert_json_to_pdf(json_data=None, output_path=None, auto_save=True):
    """
    Convert JSON data to PDF table
    
    Args:
        json_data: JSON data to convert (list of dictionaries)
        output_path: Optional output file path. If None, uses data.pdf in current directory
        auto_save: If True, automatically saves as data.pdf without asking
        
    Returns:
        str: Path to saved PDF file, or None if failed
    """
    try:
        # If no JSON data provided, try to load from data.json
        if json_data is None:
            if os.path.exists('data.json'):
                with open('data.json', 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
            else:
                QMessageBox.warning(None, "Error", "No JSON data found!")
                return None
        
        if not json_data:
            QMessageBox.warning(None, "Error", "No data to convert!")
            return None
        
        # Determine output path
        if output_path is None:
            if auto_save:
                # Auto-save as data.pdf in current directory
                output_path = 'data.pdf'
            else:
                # Ask user to save
                output_path, _ = QFileDialog.getSaveFileName(
                    None,
                    "Save PDF As",
                    f"data.pdf",
                    "PDF Files (*.pdf)"
                )
                
                if not output_path:  # User cancelled
                    return None
        
        # Extract headers from first data item
        headers = list(json_data[0].keys())
        
        # Create data for table (headers + rows)
        table_data = [headers]
        for item in json_data:
            row = [str(item.get(header, "")) for header in headers]
            table_data.append(row)
        
        # Create PDF document
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            rightMargin=30,
            leftMargin=30,
            topMargin=30,
            bottomMargin=30
        )
        
        # Create story (content) for PDF
        story = []
        
        # Add title
        styles = getSampleStyleSheet()
        title = Paragraph("JSON Data Table", styles['Title'])
        story.append(title)
        
        # Add timestamp
        timestamp = Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", 
                            styles['Normal'])
        story.append(timestamp)
        story.append(Paragraph("<br/><br/>", styles['Normal']))
        
        # Create table
        table = Table(table_data, repeatRows=1)
        
        # Style the table
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # Header background
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Header text color
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Align all cells left
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Header font
            ('FONTSIZE', (0, 0), (-1, 0), 10),  # Header font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Header padding
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # Row background
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),  # Row text color
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Row font
            ('FONTSIZE', (0, 1), (-1, -1), 9),  # Row font size
            ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Grid lines
            ('ROWHEIGHTS', (0, 0), (-1, -1), 20),  # Row height
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Vertical align middle
        ])
        
        # Add alternating row colors
        for i in range(1, len(table_data)):
            if i % 2 == 0:
                table_style.add('BACKGROUND', (0, i), (-1, i), colors.lightgrey)
        
        table.setStyle(table_style)
        story.append(table)
        
        # Build PDF
        doc.build(story)
        
        # Show success message
        if auto_save:
            QMessageBox.information(
                None,
                "Success",
                f"PDF successfully created as:\n{os.path.abspath(output_path)}"
            )
        
        return os.path.abspath(output_path)
        
    except Exception as e:
        QMessageBox.critical(
            None,
            "Error",
            f"Failed to create PDF:\n{str(e)}"
        )
        return None

def create_pdf_button(json_data=None, auto_save=True):
    """
    Create a button widget for converting JSON to PDF
    
    Args:
        json_data: JSON data to convert (optional)
        auto_save: If True, automatically saves as data.pdf without asking
        
    Returns:
        QPushButton: Configured button widget
    """
    button = QPushButton("Export to PDF")
    button.setToolTip("Convert JSON data to PDF table")
    button.setMinimumHeight(40)
    button.setStyleSheet("""
        QPushButton {
            background-color: #4CAF50;
            color: white;
            font-weight: bold;
            border: none;
            border-radius: 4px;
            padding: 10px;
            font-size: 14px;
        }
        QPushButton:hover {
            background-color: #45a049;
        }
        QPushButton:pressed {
            background-color: #3d8b40;
        }
    """)
    
    # Connect button click to conversion function
    button.clicked.connect(lambda: convert_json_to_pdf(json_data, auto_save=auto_save))
    
    return button