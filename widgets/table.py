from PyQt5.QtWidgets import (QTableView, QVBoxLayout, QWidget, 
                             QHeaderView, QAbstractItemView)
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt5.QtGui import QFont

class JsonTableModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data
        if data:
            self._headers = list(data[0].keys())
        else:
            self._headers = []

    def rowCount(self, parent=QModelIndex()):
        return len(self._data)

    def columnCount(self, parent=QModelIndex()):
        return len(self._headers)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if role == Qt.DisplayRole:
            key = self._headers[col]
            return str(self._data[row].get(key, ""))
        
        elif role == Qt.TextAlignmentRole:
            return Qt.AlignLeft | Qt.AlignVCenter
            
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._headers[section]
        return None

class JsonTableView(QTableView):
    def __init__(self, json_data=None):
        super().__init__()
        
        if json_data is not None:
            self.setModel(JsonTableModel(json_data))
        
        self.setup_ui()

    def setup_ui(self):
        # Configure table appearance
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setSortingEnabled(True)
        
        # Configure font
        font = QFont()
        font.setPointSize(10)
        self.setFont(font)
        
        # Configure header
        header = self.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(True)
        header.setDefaultAlignment(Qt.AlignLeft)
        
        vertical_header = self.verticalHeader()
        vertical_header.setVisible(False)
        
        # Set minimum sizes
        self.setMinimumSize(800, 400)

    def set_json_data(self, json_data):
        """Update the table with new JSON data"""
        self.setModel(JsonTableModel(json_data))
        self.resize_columns_to_content()

    def resize_columns_to_content(self):
        """Resize columns to fit their content"""
        self.resizeColumnsToContents()
        # Ensure minimum column width
        for column in range(self.model().columnCount()):
            width = self.columnWidth(column)
            if width < 100:
                self.setColumnWidth(column, 100)

class JsonTableWidget(QWidget):
    def __init__(self, json_data=None):
        super().__init__()
        self.json_data = json_data or []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        
        # Create table view
        self.table_view = JsonTableView(self.json_data)
        
        layout.addWidget(self.table_view)
        self.setLayout(layout)

    def update_data(self, json_data):
        """Update the table with new JSON data"""
        self.json_data = json_data
        self.table_view.set_json_data(json_data)


def get_widget_table(json_data=None):
    """
    Create and return a JsonTableWidget instance.
    
    Args:
        json_data: JSON data to display in the table (list of dictionaries)
        
    Returns:
        JsonTableWidget: Configured table widget
    """
    return JsonTableWidget(json_data)