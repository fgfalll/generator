# ═══════════════════════════════════════════════════════════════════════════════
# column_mapper.py  –  Redesigned with MS Office-style UI
# ═══════════════════════════════════════════════════════════════════════════════
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QComboBox,
    QTableWidget, QTableWidgetItem, QMessageBox, QScrollArea, QWidget,
    QGroupBox, QGridLayout, QLineEdit, QInputDialog, QFileDialog, QCheckBox,
    QFrame, QSizePolicy, QHeaderView, QAbstractItemView,
)
from PyQt6.QtCore import Qt
import qtawesome as qta
import pandas as pd


# ─────────────────────────────────────────────────────────────────────────────
# SHARED STYLE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _ok_style():
    return ("QPushButton{background:#217346;color:white;border:none;border-radius:3px;"
            "padding:6px 18px;font-size:12px;}"
            "QPushButton:hover{background:#1a5c38;}")

def _btn_style():
    return ("QPushButton{background:#f0f0f0;color:#222;border:1px solid #bbb;"
            "border-radius:3px;padding:6px 14px;font-size:12px;}"
            "QPushButton:hover{background:#e0e0e0;}")

def _warn_style():
    return ("QPushButton{background:#e74c3c;color:white;border:none;border-radius:3px;"
            "padding:6px 14px;font-size:12px;}"
            "QPushButton:hover{background:#c0392b;}")

def _row_style():
    return ("QFrame{border:1px solid #e0e0e0;border-radius:3px;background:#f8f9fa;}"
            "QFrame:hover{border-color:#90caf9;background:#fff;}")

TABLE_STYLE = """
    QTableWidget {
        border: 1px solid #d0d0d0; border-radius: 3px;
        background: white; font-size: 12px; gridline-color: #e8e8e8;
    }
    QTableWidget::item { padding: 4px 6px; }
    QTableWidget::item:selected { background: #1976d2; color: white; }
    QTableWidget::item:alternate { background: #f9f9f9; }
    QHeaderView::section {
        background: #34495e; color: white; padding: 5px 8px;
        border: none; border-right: 1px solid #4a6278;
        font-weight: bold; font-size: 12px;
    }
"""


# ─────────────────────────────────────────────────────────────────────────────
# PREVIEW WINDOW
# ─────────────────────────────────────────────────────────────────────────────
class PreviewWindow(QDialog):
    """Non-modal window for previewing and editing the DataFrame."""

    def __init__(self, main_dialog):
        super().__init__(main_dialog)
        self.setWindowTitle("Попередній перегляд та редагування даних")
        self.setGeometry(150, 150, 1200, 600)
        self.main_dialog = main_dialog

        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        # Table
        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.preview_table.setStyleSheet(TABLE_STYLE)
        root.addWidget(self.preview_table, 1)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        add_col_btn = QPushButton(qta.icon('fa5s.plus', color='#27ae60'), " Додати колонку")
        add_col_btn.setStyleSheet(_btn_style())
        add_col_btn.clicked.connect(self._add_column)

        rem_col_btn = QPushButton(qta.icon('fa5s.trash', color='#e74c3c'), " Видалити колонку")
        rem_col_btn.setStyleSheet(_btn_style())
        rem_col_btn.clicked.connect(self._remove_column)

        save_as_btn = QPushButton(qta.icon('fa5s.save', color='#217346'), " Зберегти як...")
        save_as_btn.setStyleSheet(_ok_style())
        save_as_btn.clicked.connect(self._save_as)

        btn_row.addWidget(add_col_btn)
        btn_row.addWidget(rem_col_btn)
        btn_row.addStretch()
        btn_row.addWidget(save_as_btn)
        root.addLayout(btn_row)

        self.preview_table.itemChanged.connect(self._handle_item_changed)
        self.refresh_table()

    def refresh_table(self):
        self.preview_table.blockSignals(True)
        self.preview_table.clear()
        df = self.main_dialog.df
        self.preview_table.setRowCount(df.shape[0])
        self.preview_table.setColumnCount(df.shape[1])
        self.preview_table.setHorizontalHeaderLabels(df.columns)
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                v = df.iloc[i, j]
                self.preview_table.setItem(i, j, QTableWidgetItem(str(v) if pd.notna(v) else ""))
        self.preview_table.resizeColumnsToContents()
        self.preview_table.blockSignals(False)

    def _handle_item_changed(self, item):
        try:
            self.main_dialog.df.iloc[item.row(), item.column()] = item.text()
        except Exception:
            pass

    def _add_column(self):
        name, ok = QInputDialog.getText(self, "Додати колонку", "Введіть назву нової колонки:")
        if ok and name:
            if name in self.main_dialog.df.columns:
                QMessageBox.warning(self, "Помилка", f"Колонка '{name}' вже існує.")
                return
            self.main_dialog.df[name] = ''
            self.main_dialog._invalidate_undo_state()
            self.refresh_table()
            self.main_dialog._update_all_combos()

    def _remove_column(self):
        col_idx = self.preview_table.currentColumn()
        if col_idx == -1:
            QMessageBox.warning(self, "Помилка", "Клацніть на комірку в колонці для видалення.")
            return
        col_name = self.main_dialog.df.columns[col_idx]
        reply = QMessageBox.question(self, "Видалити колонку", f"Видалити колонку '{col_name}'?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.main_dialog.df.drop(columns=[col_name], inplace=True)
            self.main_dialog._invalidate_undo_state()
            self.refresh_table()
            self.main_dialog._update_all_combos()

    def _save_as(self):
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти файл як...", "", "Excel Files (*.xlsx)")
        if path:
            try:
                self.main_dialog.df.to_excel(path, index=False)
                QMessageBox.information(self, "Успіх", f"Файл збережено:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Помилка збереження", f"Не вдалося зберегти: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN MAPPING DIALOG
# ─────────────────────────────────────────────────────────────────────────────
class ColumnMappingDialog(QDialog):
    """MS Office-style dialog for column verification, cleaning and mapping."""

    IDEAL_COLUMN_ORDER = [
        "Тип пропозиції", "Назва КП", "Освітній ступінь", "Вступ на основі", "Спеціальність", "Форма навчання",
        "Курс", "Структурний підрозділ", "Чи скорочений термін навчання", "Прізвище", "Ім'я", "По батькові",
        "Контактний номер", "Електронна адреса", "Статус заявки", "Назва групи", "Реєстраційний номер",
        "Номер групи", "Конкурсний бал", "Бютжет чи контракт", "Чи подано оригінал",
        "Чи подано довідку про місцезнаходження оригіналів",
        "Особа претендує на застосування сільського коефіцієнту",
        "Номер протоколу", "Дата протоколу", "Наказ про зарахування", "Дата наказу", "Дата вступу",
        "Тип документа", "Додаток до типу документу", "Серія документа", "Номер документа",
        "Дата видачі документа", "Ким видано", "Відзнака", "Дата подачі заяви", "Номер зно", "Рік зно",
        "ЗНО.Українська мова", "ЗНО.Українська мова та література", "ЗНО.Історія України",
        "ЗНО.Математика", "ЗНО.Біологія", "ЗНО.Географія", "ЗНО.Англійська мова",
        "ДПО", "Адреса", "ДПО.Серія", "ДПО.Номер", "ДПО.Ким виданий",
        "ДПО.Дата видачі", "ДПО.Дійсний до", "РНОКПП",
    ]

    def __init__(self, df, required_columns, column_mappings=None,
                 backend=None, parent=None, is_custom_keyword_dialog=False):
        super().__init__(parent)

        title = ("Співставлення користувацьких полів шаблону"
                 if is_custom_keyword_dialog
                 else "Перевірка, очищення та співставлення даних")
        self.setWindowTitle(title)
        self.setGeometry(100, 100, 1060, 780)

        self.original_df        = df.copy()
        self.df                 = df.copy()
        self.required_columns   = required_columns
        self.column_mappings    = column_mappings if column_mappings else {}
        self.combo_boxes        = {}
        self.backend            = backend
        self.filter_widgets     = []
        self.preview_window     = None
        self.is_custom          = is_custom_keyword_dialog
        self.delimiter_map      = {
            ", (Кома)": ",", "; (Крапка з комою)": ";",
            "\t (Табуляція)": "\t", " (Пробіл)": " ",
        }
        self.pre_split_df = None

        self._build_ui()
        if not is_custom_keyword_dialog:
            self._add_filter_row(is_deletable=False)
        self._update_all_widgets()
        self._automap_columns()

    # ─────────────────────────────────────────────────────────────────────────
    # UI BUILD
    # ─────────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(14, 12, 14, 10)
        root.setSpacing(10)

        # ── Section 1: Mapping ───────────────────────────────────────────────
        map_section_title = ("Співставте нові поля з колонками Excel"
                             if self.is_custom
                             else "1. Співставлення колонок")
        map_group = QGroupBox(map_section_title)
        map_group.setStyleSheet("QGroupBox{font-size:12px;font-weight:bold;color:#2c3e50;"
                                "border:1px solid #d0d0d0;border-radius:4px;margin-top:6px;}"
                                "QGroupBox::title{subcontrol-origin:margin;left:8px;padding:0 4px;}")
        map_vlay = QVBoxLayout(map_group)
        map_vlay.setContentsMargins(8, 14, 8, 8)
        map_vlay.setSpacing(4)

        # Column header
        hdr = QHBoxLayout()
        for txt, stretch in [("Системне поле", 2), ("Колонка у файлі", 3)]:
            lb = QLabel(f"<b>{txt}</b>")
            lb.setStyleSheet("color:#555;font-size:11px;padding:2px 6px;")
            hdr.addWidget(lb, stretch)
        map_vlay.addLayout(hdr)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setMaximumHeight(220)
        scroll.setStyleSheet("QScrollArea{border:1px solid #d0d0d0;border-radius:3px;}")
        scroll_w = QWidget()
        self._map_layout = QVBoxLayout(scroll_w)
        self._map_layout.setContentsMargins(4, 4, 4, 4)
        self._map_layout.setSpacing(3)

        for req_col in self.required_columns:
            rf = QFrame()
            rf.setFixedHeight(34)
            rf.setStyleSheet(_row_style())
            rl = QHBoxLayout(rf)
            rl.setContentsMargins(8, 2, 8, 2)
            rl.setSpacing(8)

            lbl = QLabel(req_col)
            lbl.setStyleSheet("font-size:12px;color:#2c3e50;border:none;background:transparent;")
            lbl.setAlignment(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)

            combo = QComboBox()
            combo.setStyleSheet("font-size:12px;")
            combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            combo.addItem("— Пропустити —")
            self.combo_boxes[req_col] = combo

            rl.addWidget(lbl, 2)
            rl.addWidget(combo, 3)
            self._map_layout.addWidget(rf)

        self._map_layout.addStretch()
        scroll.setWidget(scroll_w)
        map_vlay.addWidget(scroll)
        root.addWidget(map_group)

        # ── Section 2: Cleanup (only for non-custom) ─────────────────────────
        if not self.is_custom:
            cleanup_group = QGroupBox("2. Очищення та трансформація даних")
            cleanup_group.setStyleSheet(map_group.styleSheet())
            cg_lay = QVBoxLayout(cleanup_group)
            cg_lay.setContentsMargins(8, 14, 8, 8)
            cg_lay.setSpacing(8)

            # Structure
            struct_row = QHBoxLayout()
            struct_lbl = QLabel("Структура таблиці:")
            struct_lbl.setStyleSheet("font-size:12px;color:#555;")
            conform_btn = QPushButton(qta.icon('fa5s.table', color='#1565c0'), " Сформувати ідеальну структуру")
            conform_btn.setStyleSheet(_btn_style())
            conform_btn.clicked.connect(self._conform_to_ideal_structure)
            struct_row.addWidget(struct_lbl)
            struct_row.addWidget(conform_btn)
            struct_row.addStretch()
            cg_lay.addLayout(struct_row)

            # Filters
            filters_box = QGroupBox("Фільтрація рядків")
            filters_box.setStyleSheet("QGroupBox{font-size:11px;color:#555;border:1px solid #e0e0e0;"
                                      "border-radius:3px;margin-top:6px;}"
                                      "QGroupBox::title{subcontrol-origin:margin;left:8px;padding:0 4px;}")
            filt_vlay = QVBoxLayout(filters_box)
            filt_vlay.setContentsMargins(6, 12, 6, 6)
            filt_vlay.setSpacing(4)
            self.filters_layout = QVBoxLayout()
            self.filters_layout.setSpacing(3)
            filt_vlay.addLayout(self.filters_layout)

            filt_btns = QHBoxLayout()
            add_filter_btn = QPushButton(qta.icon('fa5s.plus', color='#27ae60'), " Додати фільтр")
            add_filter_btn.setStyleSheet(_btn_style())
            add_filter_btn.clicked.connect(self._add_filter_row)
            apply_filters_btn = QPushButton(qta.icon('fa5s.check', color='#27ae60'), " Застосувати фільтри")
            apply_filters_btn.setStyleSheet(_ok_style())
            apply_filters_btn.clicked.connect(self._apply_all_filters)
            filt_btns.addWidget(add_filter_btn)
            filt_btns.addStretch()
            filt_btns.addWidget(apply_filters_btn)
            filt_vlay.addLayout(filt_btns)
            cg_lay.addWidget(filters_box)

            # Split
            split_box = QGroupBox("Розділення колонки")
            split_box.setStyleSheet(filters_box.styleSheet())
            split_grid = QGridLayout(split_box)
            split_grid.setContentsMargins(6, 12, 6, 6)
            split_grid.setSpacing(6)

            split_grid.addWidget(QLabel("Розділити:"), 0, 0)
            self.split_column_combo = QComboBox()
            self.split_column_combo.setStyleSheet("font-size:12px;")
            split_grid.addWidget(self.split_column_combo, 0, 1)

            split_grid.addWidget(QLabel("Роздільник:"), 0, 2)
            delim_inner = QHBoxLayout()
            self.delimiter_combo = QComboBox()
            self.delimiter_combo.setStyleSheet("font-size:12px;")
            self.delimiter_combo.addItems(list(self.delimiter_map.keys()) + ["Інший..."])
            self.custom_delimiter_edit = QLineEdit()
            self.custom_delimiter_edit.setVisible(False)
            self.custom_delimiter_edit.setStyleSheet("font-size:12px;")
            delim_inner.addWidget(self.delimiter_combo, 1)
            delim_inner.addWidget(self.custom_delimiter_edit, 1)
            split_grid.addLayout(delim_inner, 0, 3)

            split_grid.addWidget(QLabel("Нові назви:"), 1, 0)
            self.new_column_names_edit = QLineEdit()
            self.new_column_names_edit.setPlaceholderText("Ім'я,По батькові")
            self.new_column_names_edit.setStyleSheet("font-size:12px;")
            split_grid.addWidget(self.new_column_names_edit, 1, 1, 1, 3)

            split_btns = QHBoxLayout()
            apply_split_btn = QPushButton(qta.icon('fa5s.grip-lines-vertical', color='#e67e22'), " Розділити")
            apply_split_btn.setStyleSheet(_btn_style())
            apply_split_btn.clicked.connect(self._split_column)
            self.revert_split_button = QPushButton(qta.icon('fa5s.undo', color='#c0392b'), " Скасувати")
            self.revert_split_button.setStyleSheet(_warn_style())
            self.revert_split_button.setVisible(False)
            self.revert_split_button.clicked.connect(self._revert_split)
            split_btns.addWidget(apply_split_btn)
            split_btns.addWidget(self.revert_split_button)
            split_grid.addLayout(split_btns, 0, 4, 2, 1)

            self.delimiter_combo.currentTextChanged.connect(self._on_delimiter_changed)
            cg_lay.addWidget(split_box)

            # Merge
            merge_box = QGroupBox("Об'єднання колонок")
            merge_box.setStyleSheet(filters_box.styleSheet())
            merge_grid = QGridLayout(merge_box)
            merge_grid.setContentsMargins(6, 12, 6, 6)
            merge_grid.setSpacing(6)
            merge_grid.setColumnStretch(1, 1)

            merge_grid.addWidget(QLabel("Колонки для об'єднання:"), 0, 0)
            self.merge_column_names_edit = QLineEdit()
            self.merge_column_names_edit.setPlaceholderText("Прізвище, Ім'я, По батькові")
            self.merge_column_names_edit.setStyleSheet("font-size:12px;")
            merge_grid.addWidget(self.merge_column_names_edit, 0, 1)

            merge_grid.addWidget(QLabel("Нова назва колонки:"), 1, 0)
            self.new_merged_column_name_edit = QLineEdit()
            self.new_merged_column_name_edit.setPlaceholderText("ПІБ")
            self.new_merged_column_name_edit.setStyleSheet("font-size:12px;")
            merge_grid.addWidget(self.new_merged_column_name_edit, 1, 1)

            merge_grid.addWidget(QLabel("Роздільник:"), 2, 0)
            self.merge_delimiter_edit = QLineEdit()
            self.merge_delimiter_edit.setPlaceholderText(" (пробіл)")
            self.merge_delimiter_edit.setStyleSheet("font-size:12px;")
            merge_grid.addWidget(self.merge_delimiter_edit, 2, 1)

            self.delete_original_cols_check = QCheckBox("Видалити вихідні колонки")
            self.delete_original_cols_check.setChecked(True)
            merge_grid.addWidget(self.delete_original_cols_check, 3, 1)

            apply_merge_btn = QPushButton(qta.icon('fa5s.compress-arrows-alt', color='#9b59b6'), " Об'єднати")
            apply_merge_btn.setStyleSheet(_btn_style())
            apply_merge_btn.clicked.connect(self._merge_columns)
            merge_grid.addWidget(apply_merge_btn, 0, 2, 4, 1)

            cg_lay.addWidget(merge_box)
            root.addWidget(cleanup_group)

            # ── Section 3: Preview ───────────────────────────────────────────
            preview_group = QGroupBox("3. Попередній перегляд")
            preview_group.setStyleSheet(map_group.styleSheet())
            prev_lay = QHBoxLayout(preview_group)
            prev_lay.setContentsMargins(8, 14, 8, 8)
            self.preview_button = QPushButton(
                qta.icon('fa5s.eye', color='#1565c0'), " Попередній перегляд та редагування")
            self.preview_button.setStyleSheet(_btn_style())
            prev_lay.addWidget(self.preview_button)
            prev_lay.addStretch()
            root.addWidget(preview_group)
            self.preview_button.clicked.connect(self._open_preview_window)

        # ── Bottom buttons ───────────────────────────────────────────────────
        bot = QHBoxLayout()
        bot.setSpacing(8)

        if not self.is_custom:
            test_btn = QPushButton(qta.icon('fa5s.vial', color='#8e44ad'), " Згенерувати тест")
            test_btn.setStyleSheet(_btn_style())
            test_btn.clicked.connect(self._generate_test_file)
            reset_btn = QPushButton(qta.icon('fa5s.undo', color='#c0392b'), " Скинути зміни")
            reset_btn.setStyleSheet(_btn_style())
            reset_btn.clicked.connect(self._reset_data)
            bot.addWidget(test_btn)
            bot.addWidget(reset_btn)

        bot.addStretch()

        ok_btn = QPushButton("Зберегти та закрити")
        ok_btn.setDefault(True)
        ok_btn.setStyleSheet(_ok_style())
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Скасувати")
        cancel_btn.setStyleSheet(_btn_style())
        cancel_btn.clicked.connect(self.reject)

        bot.addWidget(ok_btn)
        bot.addWidget(cancel_btn)
        root.addLayout(bot)

    # ─────────────────────────────────────────────────────────────────────────
    # DATA / COMBO HELPERS
    # ─────────────────────────────────────────────────────────────────────────
    def _invalidate_undo_state(self):
        self.pre_split_df = None
        if hasattr(self, 'revert_split_button'):
            self.revert_split_button.setVisible(False)

    def _update_all_widgets(self):
        self._update_all_combos()
        self._notify_preview_window()

    def _update_all_combos(self):
        cols = self.df.columns.tolist()
        combos = list(self.combo_boxes.values())
        if hasattr(self, 'split_column_combo'):
            combos.append(self.split_column_combo)

        for combo in combos:
            curr = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            if combo in self.combo_boxes.values():
                combo.addItem("— Пропустити —")
            combo.addItems(cols)
            if curr in cols:
                combo.setCurrentText(curr)
            elif combo in self.combo_boxes.values():
                combo.setCurrentIndex(0)
            combo.blockSignals(False)

        if hasattr(self, 'filter_widgets'):
            self._update_filter_combos()

    def _update_filter_combos(self):
        cols = self.df.columns.tolist()
        for fw in self.filter_widgets:
            curr = fw["col_combo"].currentText()
            fw["col_combo"].blockSignals(True)
            fw["col_combo"].clear()
            fw["col_combo"].addItems(cols)
            if curr in cols:
                fw["col_combo"].setCurrentText(curr)
            fw["col_combo"].blockSignals(False)

    def _notify_preview_window(self):
        if self.preview_window and self.preview_window.isVisible():
            self.preview_window.refresh_table()

    def _automap_columns(self):
        for req in self.required_columns:
            if req in self.df.columns and req in self.combo_boxes:
                combo = self.combo_boxes[req]
                # Find the item index (offset by 1 for "—Пропустити—")
                idx = combo.findText(req)
                if idx >= 0:
                    combo.setCurrentIndex(idx)

    # ─────────────────────────────────────────────────────────────────────────
    # FILTER ROWS
    # ─────────────────────────────────────────────────────────────────────────
    def _add_filter_row(self, is_deletable=True):
        fw_widget = QFrame()
        fw_widget.setFixedHeight(36)
        fw_widget.setStyleSheet(_row_style())
        row = QHBoxLayout(fw_widget)
        row.setContentsMargins(6, 2, 6, 2)
        row.setSpacing(6)

        col = QComboBox()
        col.setStyleSheet("font-size:12px;")
        col.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        cond = QComboBox()
        cond.setStyleSheet("font-size:12px;")
        cond.addItems(["містить", "не містить", "дорівнює", "не дорівнює"])

        val = QLineEdit()
        val.setPlaceholderText("Значення")
        val.setStyleSheet("font-size:12px;")
        val.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        row.addWidget(col, 2)
        row.addWidget(cond, 2)
        row.addWidget(val, 3)

        fw = {"row": fw_widget, "col_combo": col, "cond_combo": cond, "val_edit": val}

        if is_deletable:
            rem_btn = QPushButton("✕")
            rem_btn.setFixedSize(24, 24)
            rem_btn.setStyleSheet("QPushButton{background:#e74c3c;color:white;border:none;"
                                  "border-radius:3px;font-size:11px;}"
                                  "QPushButton:hover{background:#c0392b;}")
            # Correct closure capture
            rem_btn.clicked.connect(lambda checked=False, f=fw: self._remove_filter_row(f))
            row.addWidget(rem_btn)

        self.filters_layout.addWidget(fw_widget)
        self.filter_widgets.append(fw)
        self._update_filter_combos()

    def _remove_filter_row(self, fw):
        if fw in self.filter_widgets:
            self.filter_widgets.remove(fw)
        fw["row"].hide()
        fw["row"].deleteLater()

    def _apply_all_filters(self):
        self._invalidate_undo_state()
        tmp = self.df.copy()
        for fw in self.filter_widgets:
            col, val = fw["col_combo"].currentText(), fw["val_edit"].text()
            if not col or not val:
                continue
            cond = fw["cond_combo"].currentText()
            s = tmp[col].astype(str)
            if   cond == "містить":    mask = s.str.contains(val, case=False, na=False)
            elif cond == "не містить": mask = ~s.str.contains(val, case=False, na=False)
            elif cond == "дорівнює":   mask = s.str.lower() == val.lower()
            else:                      mask = s.str.lower() != val.lower()
            tmp = tmp[mask]
        self.df = tmp.reset_index(drop=True)
        self._notify_preview_window()
        QMessageBox.information(self, "Фільтри застосовано", f"Залишилось рядків: {len(self.df)}")

    # ─────────────────────────────────────────────────────────────────────────
    # CONFORM STRUCTURE
    # ─────────────────────────────────────────────────────────────────────────
    def _conform_to_ideal_structure(self):
        reply = QMessageBox.question(
            self, "Сформувати ідеальну структуру?",
            "Ця дія додасть відсутні ідеальні колонки, видалить усі зайві та змінить порядок існуючих.\n"
            "Дані в зайвих колонках буде втрачено. Цю дію не можна скасувати.\n\nПродовжити?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.No:
            return
        self._invalidate_undo_state()
        new_df = pd.DataFrame(index=self.df.index)
        for col in self.IDEAL_COLUMN_ORDER:
            new_df[col] = self.df[col] if col in self.df.columns else ""
        self.df = new_df
        self._update_all_widgets()
        self._automap_columns()
        QMessageBox.information(self, "Успіх", "Структуру таблиці приведено до ідеального формату.")

    # ─────────────────────────────────────────────────────────────────────────
    # SPLIT
    # ─────────────────────────────────────────────────────────────────────────
    def _on_delimiter_changed(self, text):
        self.custom_delimiter_edit.setVisible(text == "Інший...")

    def _split_column(self):
        col_split  = self.split_column_combo.currentText()
        delim_txt  = self.delimiter_combo.currentText()
        names_str  = self.new_column_names_edit.text()
        delim      = (self.custom_delimiter_edit.text() if delim_txt == "Інший..."
                      else self.delimiter_map.get(delim_txt))
        if not all([col_split, delim is not None, names_str]):
            QMessageBox.warning(self, "Недостатньо даних", "Заповніть усі поля для розділення.")
            return
        names = [n.strip() for n in names_str.split(',') if n.strip()]
        if not names:
            QMessageBox.warning(self, "Недостатньо даних", "Вкажіть назви нових колонок.")
            return
        for n in names:
            if n in self.df.columns and n != col_split:
                QMessageBox.warning(self, "Конфлікт імен", f"Колонка '{n}' вже існує.")
                return
        try:
            self.pre_split_df = self.df.copy()
            idx        = self.df.columns.get_loc(col_split)
            split_data = self.df[col_split].astype(str).str.split(delim, n=len(names) - 1, expand=True)
            self.df.drop(columns=[col_split], inplace=True)
            for i, n in reversed(list(enumerate(names))):
                cd = split_data[i].fillna('') if i < split_data.shape[1] else [''] * len(self.df)
                self.df.insert(idx, n, cd)
            self._update_all_widgets()
            self.revert_split_button.setVisible(True)
            QMessageBox.information(self, "Успіх", f"Колонку '{col_split}' розділено.")
        except Exception as e:
            QMessageBox.critical(self, "Помилка", f"Сталася помилка: {e}")
            self.pre_split_df = None

    def _revert_split(self):
        if self.pre_split_df is not None:
            self.df = self.pre_split_df.copy()
            self._invalidate_undo_state()
            self._update_all_widgets()
            QMessageBox.information(self, "Скасовано", "Операцію розділення скасовано.")
        else:
            QMessageBox.warning(self, "Неможливо скасувати", "Немає операції для скасування.")

    # ─────────────────────────────────────────────────────────────────────────
    # MERGE
    # ─────────────────────────────────────────────────────────────────────────
    def _merge_columns(self):
        cols_str  = self.merge_column_names_edit.text()
        new_name  = self.new_merged_column_name_edit.text().strip()
        delimiter = self.merge_delimiter_edit.text()

        if not all([cols_str, new_name]):
            QMessageBox.warning(self, "Недостатньо даних",
                                "Вкажіть колонки для об'єднання та назву нової колонки.")
            return
        cols_to_merge = [c.strip() for c in cols_str.split(',') if c.strip()]
        if len(cols_to_merge) < 2:
            QMessageBox.warning(self, "Помилка", "Потрібно щонайменше дві колонки.")
            return
        for c in cols_to_merge:
            if c not in self.df.columns:
                QMessageBox.warning(self, "Помилка", f"Колонку '{c}' не знайдено.")
                return
        if new_name in self.df.columns and new_name not in cols_to_merge:
            QMessageBox.warning(self, "Конфлікт", f"Колонка '{new_name}' вже існує.")
            return
        try:
            self._invalidate_undo_state()
            self.df[new_name] = self.df[cols_to_merge].astype(str).agg(delimiter.join, axis=1)
            if self.delete_original_cols_check.isChecked():
                drop = [c for c in cols_to_merge if c != new_name]
                if drop:
                    self.df.drop(columns=drop, inplace=True)
            self._update_all_widgets()
            QMessageBox.information(self, "Успіх", f"Колонки об'єднано в '{new_name}'.")
        except Exception as e:
            QMessageBox.critical(self, "Помилка", f"Сталася помилка: {e}")

    # ─────────────────────────────────────────────────────────────────────────
    # PREVIEW
    # ─────────────────────────────────────────────────────────────────────────
    def _open_preview_window(self):
        if not self.preview_window or not self.preview_window.isVisible():
            self.preview_window = PreviewWindow(self)
            self.preview_window.show()
        else:
            self.preview_window.raise_()
            self.preview_window.activateWindow()

    # ─────────────────────────────────────────────────────────────────────────
    # TEST GENERATION
    # ─────────────────────────────────────────────────────────────────────────
    def _generate_test_file(self):
        if not self.backend:
            QMessageBox.critical(self, "Помилка", "Зв'язок з додатком втрачено.")
            return
        templates = list(self.backend.template_paths.keys())
        if not templates:
            QMessageBox.warning(self, "Немає шаблонів", "Не знайдено шаблонів.")
            return
        sel_row = -1
        if self.preview_window and self.preview_window.isVisible():
            sel_row = self.preview_window.preview_table.currentRow()
        if sel_row < 0:
            QMessageBox.warning(self, "Рядок не вибрано",
                                "Відкрийте попередній перегляд та виберіть рядок.")
            return
        template, ok = QInputDialog.getItem(
            self, "Вибір шаблону", "Виберіть шаблон:", sorted(templates), 0, False)
        if ok and template:
            row_data = self.df.iloc[sel_row]
            mappings = self._get_current_mappings_from_ui()
            self.backend.handle_test_file_generation(row_data, mappings, template)

    # ─────────────────────────────────────────────────────────────────────────
    # RESET
    # ─────────────────────────────────────────────────────────────────────────
    def _reset_data(self):
        reply = QMessageBox.question(
            self, "Скинути всі зміни",
            "Відновити дані до стану на момент відкриття вікна?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self._invalidate_undo_state()
            self.df = self.original_df.copy()
            for fw in self.filter_widgets:
                fw["row"].deleteLater()
            self.filter_widgets.clear()
            self._add_filter_row(is_deletable=False)
            self._update_all_widgets()
            self._automap_columns()

    # ─────────────────────────────────────────────────────────────────────────
    # ACCEPT / GETTERS
    # ─────────────────────────────────────────────────────────────────────────
    def _get_current_mappings_from_ui(self):
        return {r: c.currentText()
                for r, c in self.combo_boxes.items()
                if c.currentText() not in ("— Пропустити —", "Пропустити")}

    def accept(self):
        mappings = self._get_current_mappings_from_ui()
        if len(mappings.values()) != len(set(mappings.values())):
            QMessageBox.warning(self, "Дублювання колонок",
                                "Одна колонка не може бути призначена для кількох полів.")
            return
        self.column_mappings = mappings
        super().accept()

    def closeEvent(self, event):
        if self.preview_window:
            self.preview_window.close()
        super().closeEvent(event)

    def get_mapped_columns(self):
        return self.column_mappings

    def get_modified_df(self):
        return self.df