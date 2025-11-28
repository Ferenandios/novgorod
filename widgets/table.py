from PyQt5.QtWidgets import (QTableView, QVBoxLayout, QWidget, 
                             QHeaderView, QAbstractItemView, QLineEdit)
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex
from PyQt5.QtGui import QFont, QKeyEvent

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

    def setData(self, index, value, role=Qt.EditRole):
        """Set data in the model when cell is edited"""
        if role == Qt.EditRole:
            row = index.row()
            col = index.column()
            
            if 0 <= row < len(self._data) and 0 <= col < len(self._headers):
                key = self._headers[col]
                self._data[row][key] = value
                self.dataChanged.emit(index, index, [role])
                return True
        return False

    def flags(self, index):
        """Make cells editable"""
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

class JsonTableView(QTableView):
    def __init__(self, json_data=None):
        super().__init__()
        
        if json_data is not None:
            self.setModel(JsonTableModel(json_data))
        
        self.setup_ui()
        self.setup_signals()

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

    def setup_signals(self):
        """Connect signals for cell editing"""
        self.doubleClicked.connect(self.on_cell_double_click)

    def keyPressEvent(self, event):
        """Handle key press events - F2 for editing"""
        if event.key() == Qt.Key_F2:
            self.edit_current_cell()
        else:
            super().keyPressEvent(event)

    def edit_current_cell(self):
        """Edit the currently selected cell (like double-click but for F2)"""
        current_index = self.currentIndex()
        if current_index.isValid():
            self.start_cell_editing(current_index)

    def on_cell_double_click(self, index):
        """Handle cell editing when user double-clicks a cell"""
        if index.isValid():
            self.start_cell_editing(index)

    def start_cell_editing(self, index):
        """Start editing a cell with current value (like HTML input value='current_value')"""
        # Get the current value from the cell
        model = self.model()
        current_value = model.data(index, Qt.DisplayRole)
        
        # Start editing - this creates the input field
        self.edit(index)
        
        # Get the editor widget (the input field) and set its value
        editor = self.findChild(QLineEdit)
        if editor:
            editor.setText(str(current_value))  # Like <input value="current_value">
            editor.selectAll()  # Select all text for easy editing

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

    def get_current_data(self):
        """Get the current JSON data from the table"""
        if hasattr(self.table_view.model(), '_data'):
            return self.table_view.model()._data
        return self.json_data


def get_widget_table(json_data=None):
    """
    Create and return a JsonTableWidget instance.
    
    Args:
        json_data: JSON data to display in the table (list of dictionaries)
        
    Returns:
        JsonTableWidget: Configured table widget
    """
    return JsonTableWidget(json_data)