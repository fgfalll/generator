# ═══════════════════════════════════════════════════════════════════════════════
# score_mapper_dialog.py  –  Redesigned with MS Office-style UI
# ═══════════════════════════════════════════════════════════════════════════════
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QComboBox,
    QTableWidget, QTableWidgetItem, QCheckBox, QHeaderView,
    QWidget, QMessageBox, QFrame, QSizePolicy,
)
from PyQt6.QtCore import Qt
import qtawesome as qta


def _ok_style():
    return ("QPushButton{background:#217346;color:white;border:none;border-radius:3px;"
            "padding:6px 20px;font-size:12px;}"
            "QPushButton:hover{background:#1a5c38;}")

def _btn_style():
    return ("QPushButton{background:#f0f0f0;color:#222;border:1px solid #bbb;"
            "border-radius:3px;padding:6px 14px;font-size:12px;}"
            "QPushButton:hover{background:#e0e0e0;}")


class ScoreMappingDialog(QDialog):
    """Configure score columns and their template keys."""

    def __init__(self, available_columns, initial_mappings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Налаштування колонок з балами")
        self.setMinimumSize(680, 460)
        self.resize(720, 520)

        self.available_columns = available_columns
        self.mappings = list(initial_mappings)

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 12, 14, 10)
        root.setSpacing(10)

        # ── Title bar ────────────────────────────────────────────────────────
        title = QLabel("Налаштування колонок з балами")
        title.setStyleSheet("font-size:14px;font-weight:bold;color:#2c3e50;")
        root.addWidget(title)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#d0d0d0;")
        root.addWidget(sep)

        # ── Add row bar ──────────────────────────────────────────────────────
        add_row = QFrame()
        add_row.setStyleSheet("background:#f5f5f5;border:1px solid #e0e0e0;border-radius:4px;")
        add_lay = QHBoxLayout(add_row)
        add_lay.setContentsMargins(8, 6, 8, 6)
        add_lay.setSpacing(8)

        lbl = QLabel("Додати колонку:")
        lbl.setStyleSheet("font-size:12px;color:#555;")
        self.columns_combo = QComboBox()
        self.columns_combo.setStyleSheet("font-size:12px;")
        self.columns_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.columns_combo.addItems(self.available_columns)

        self.add_button = QPushButton(qta.icon('fa5s.plus', color='#27ae60'), " Додати")
        self.add_button.setStyleSheet(_ok_style())
        self.add_button.clicked.connect(self._add_row)

        add_lay.addWidget(lbl)
        add_lay.addWidget(self.columns_combo, 1)
        add_lay.addWidget(self.add_button)
        root.addWidget(add_row)

        # ── Mappings table ───────────────────────────────────────────────────
        self.mappings_table = QTableWidget()
        self.mappings_table.setColumnCount(4)
        self.mappings_table.setHorizontalHeaderLabels([
            "Колонка з файлу", "Ключ (число)", "Прописом?", "Ключ (прописом)"
        ])
        self.mappings_table.setAlternatingRowColors(True)
        self.mappings_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows)
        self.mappings_table.verticalHeader().setVisible(False)
        self.mappings_table.setShowGrid(True)
        self.mappings_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #d0d0d0;
                border-radius: 3px;
                background: white;
                font-size: 12px;
                gridline-color: #e8e8e8;
            }
            QTableWidget::item { padding: 4px 6px; }
            QTableWidget::item:selected { background: #1976d2; color: white; }
            QTableWidget::item:alternate { background: #f9f9f9; }
            QHeaderView::section {
                background: #34495e; color: white;
                padding: 5px 8px; border: none;
                border-right: 1px solid #4a6278;
                font-weight: bold; font-size: 12px;
            }
        """)
        hdr = self.mappings_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        root.addWidget(self.mappings_table, 1)

        # ── Bottom buttons ───────────────────────────────────────────────────
        bot = QHBoxLayout()
        bot.setSpacing(8)

        self.remove_button = QPushButton(qta.icon('fa5s.trash', color='#e74c3c'), " Видалити вибране")
        self.remove_button.setStyleSheet(_btn_style())
        self.remove_button.clicked.connect(self._remove_row)
        bot.addWidget(self.remove_button)
        bot.addStretch()

        ok_btn = QPushButton("Зберегти")
        ok_btn.setDefault(True)
        ok_btn.setStyleSheet(_ok_style())
        ok_btn.clicked.connect(self.accept)

        cancel_btn = QPushButton("Скасувати")
        cancel_btn.setStyleSheet(_btn_style())
        cancel_btn.clicked.connect(self.reject)

        bot.addWidget(ok_btn)
        bot.addWidget(cancel_btn)
        root.addLayout(bot)

        self._populate_table()

    # ─────────────────────────────────────────────────────────────────────────
    def _populate_table(self):
        self.mappings_table.setRowCount(0)
        for mapping in self.mappings:
            self._add_row(mapping=mapping, startup=True)

    def _add_row(self, *, mapping=None, startup=False):
        source_col = mapping['source'] if mapping else self.columns_combo.currentText()
        if not source_col:
            return

        # Prevent duplicates
        for row in range(self.mappings_table.rowCount()):
            item = self.mappings_table.item(row, 0)
            if item and item.text() == source_col:
                if not startup:
                    QMessageBox.warning(self, "Дублікат", f"Колонка '{source_col}' вже додана.")
                return

        r = self.mappings_table.rowCount()
        self.mappings_table.insertRow(r)
        self.mappings_table.setRowHeight(r, 32)

        # Col 0: source (read-only)
        src_item = QTableWidgetItem(source_col)
        src_item.setFlags(src_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.mappings_table.setItem(r, 0, src_item)

        # Col 1: key
        default_key = source_col.lower().replace(' ', '_').replace('.', '_')
        key = mapping.get('key', default_key) if mapping else default_key
        self.mappings_table.setItem(r, 1, QTableWidgetItem(key))

        # Col 2: checkbox (centered)
        chk_w = QWidget()
        chk_l = QHBoxLayout(chk_w)
        chk_l.setContentsMargins(0, 0, 0, 0)
        chk_l.setAlignment(Qt.AlignmentFlag.AlignCenter)
        chk = QCheckBox()
        chk.setChecked(mapping.get('add_written', True) if mapping else True)
        chk_l.addWidget(chk)
        self.mappings_table.setCellWidget(r, 2, chk_w)

        # Col 3: written key
        default_wk = f"{key}_slova"
        wk = mapping.get('written_key', default_wk) if mapping else default_wk
        wk_item = QTableWidgetItem(wk)
        self.mappings_table.setItem(r, 3, wk_item)

        # Toggle written key editability based on checkbox
        self._set_written_key_enabled(r, chk.isChecked())
        chk.toggled.connect(lambda checked, row=r: self._set_written_key_enabled(row, checked))

    def _set_written_key_enabled(self, row, enabled):
        item = self.mappings_table.item(row, 3)
        if not item:
            return
        if enabled:
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsEnabled)
            item.setForeground(Qt.GlobalColor.black)
        else:
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable & ~Qt.ItemFlag.ItemIsEnabled)
            item.setForeground(Qt.GlobalColor.gray)

    def _remove_row(self):
        row = self.mappings_table.currentRow()
        if row > -1:
            self.mappings_table.removeRow(row)

    def get_score_mappings(self):
        result = []
        for row in range(self.mappings_table.rowCount()):
            chk_w = self.mappings_table.cellWidget(row, 2)
            chk   = chk_w.findChild(QCheckBox) if chk_w else None
            result.append({
                'source':      self.mappings_table.item(row, 0).text(),
                'key':         self.mappings_table.item(row, 1).text(),
                'add_written': chk.isChecked() if chk else False,
                'written_key': self.mappings_table.item(row, 3).text(),
            })
        return result