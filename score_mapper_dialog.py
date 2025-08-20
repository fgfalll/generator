from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QComboBox,
                             QTableWidget, QTableWidgetItem, QCheckBox, QHeaderView,
                             QDialogButtonBox, QWidget, QMessageBox)
from PyQt6.QtCore import Qt

class ScoreMappingDialog(QDialog):
    """A dedicated dialog for selecting score columns and configuring their template keys."""
    def __init__(self, available_columns, initial_mappings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Налаштування колонок з балами")
        self.setMinimumSize(600, 400)

        self.available_columns = available_columns
        self.mappings = initial_mappings  # Expects a list of dicts

        main_layout = QVBoxLayout(self)

        # --- Top selection part ---
        add_group_layout = QHBoxLayout()
        add_group_layout.addWidget(QLabel("Додати колонку:"))
        self.columns_combo = QComboBox()
        self.columns_combo.addItems(self.available_columns)
        self.add_button = QPushButton("Додати")
        add_group_layout.addWidget(self.columns_combo, 1)
        add_group_layout.addWidget(self.add_button)
        main_layout.addLayout(add_group_layout)

        # --- Table for mappings ---
        self.mappings_table = QTableWidget()
        self.mappings_table.setColumnCount(4)
        self.mappings_table.setHorizontalHeaderLabels([
            "Колонка з файлу", "Ключ (число)", "Додати прописом?", "Ключ (прописом)"
        ])
        header = self.mappings_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        main_layout.addWidget(self.mappings_table)
        
        # --- Bottom buttons ---
        bottom_layout = QHBoxLayout()
        self.remove_button = QPushButton("Видалити вибране")
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.remove_button)
        main_layout.addLayout(bottom_layout)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        # --- Connections ---
        self.add_button.clicked.connect(self._add_row)
        self.remove_button.clicked.connect(self._remove_row)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)

        self._populate_table()

    def _populate_table(self):
        """Fills the table with the initial mappings."""
        self.mappings_table.setRowCount(0)
        for mapping in self.mappings:
            self._add_row(mapping=mapping, startup=True)

    def _add_row(self, *, mapping=None, startup=False):
        """Adds a new row to the mapping table, either from the combo box or from existing data."""
        if mapping:
            source_col = mapping['source']
        else:
            source_col = self.columns_combo.currentText()
            if not source_col: return

        # Prevent duplicates
        for row in range(self.mappings_table.rowCount()):
            if self.mappings_table.item(row, 0).text() == source_col:
                if not startup:
                    QMessageBox.warning(self, "Дублікат", f"Колонка '{source_col}' вже додана.")
                return

        row_position = self.mappings_table.rowCount()
        self.mappings_table.insertRow(row_position)

        # Column 0: Source Column Name (read-only)
        source_item = QTableWidgetItem(source_col)
        source_item.setFlags(source_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.mappings_table.setItem(row_position, 0, source_item)

        # --- KEY CHANGE: Removed transliteration for default key generation ---
        # Column 1: Template Key (numeric)
        default_key = source_col.lower().replace(' ', '_').replace('.', '_')
        key = mapping.get('key', default_key) if mapping else default_key
        self.mappings_table.setItem(row_position, 1, QTableWidgetItem(key))

        # Column 2: CheckBox for written form
        checkbox_widget = QWidget()
        chk_layout = QHBoxLayout(checkbox_widget)
        chk_layout.setContentsMargins(0, 0, 0, 0)
        chk_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        checkbox = QCheckBox()
        chk_layout.addWidget(checkbox)
        self.mappings_table.setCellWidget(row_position, 2, checkbox_widget)
        
        # Column 3: Template Key (written)
        default_written_key = f"{key}_slova"
        written_key = mapping.get('written_key', default_written_key) if mapping else default_written_key
        written_key_item = QTableWidgetItem(written_key)
        self.mappings_table.setItem(row_position, 3, written_key_item)

        # Set state based on mapping data
        add_written = mapping.get('add_written', True) if mapping else True
        checkbox.setChecked(add_written)
        written_key_item.setFlags(written_key_item.flags() | Qt.ItemFlag.ItemIsEnabled if add_written else written_key_item.flags() & ~Qt.ItemFlag.ItemIsEnabled)
        
        # Connect checkbox signal
        checkbox.toggled.connect(lambda checked, r=row_position: self._toggle_written_key_cell(r, checked))

    def _toggle_written_key_cell(self, row, checked):
        """Enable or disable the written key item based on the checkbox."""
        item = self.mappings_table.item(row, 3)
        if item:
            if checked:
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEnabled)
                item.setBackground(Qt.GlobalColor.transparent) 
            else:
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEnabled)

    def _remove_row(self):
        current_row = self.mappings_table.currentRow()
        if current_row > -1:
            self.mappings_table.removeRow(current_row)

    def get_score_mappings(self):
        """Parses the table and returns the configuration."""
        updated_mappings = []
        for row in range(self.mappings_table.rowCount()):
            checkbox = self.mappings_table.cellWidget(row, 2).findChild(QCheckBox)
            mapping = {
                'source': self.mappings_table.item(row, 0).text(),
                'key': self.mappings_table.item(row, 1).text(),
                'add_written': checkbox.isChecked(),
                'written_key': self.mappings_table.item(row, 3).text()
            }
            updated_mappings.append(mapping)
        return updated_mappings