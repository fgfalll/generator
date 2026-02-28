from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                             QTreeWidget, QTreeWidgetItem, QComboBox, QTableWidget, QTableWidgetItem, QCheckBox,
                             QFileDialog, QDialog, QDialogButtonBox, QTextEdit,
                             QGroupBox, QTabWidget, QAbstractItemView, QHeaderView, QScrollArea,
                             QFormLayout, QGridLayout, QLineEdit, QMessageBox, QFrame, QLayout, QSpacerItem)
from PyQt6.QtCore import Qt, pyqtSignal, QSize, QRect, QPoint
from PyQt6.QtWidgets import QSizePolicy
from datetime import datetime
import qtawesome as qta
import pandas as pd

class CsvImportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Налаштування імпорту CSV")
        self.setMinimumWidth(320)
        
        layout = QFormLayout(self)
        
        self.sep_combo = QComboBox()
        self.sep_combo.addItems(["Автовизначення", "Кома (,)", "Крапка з комою (;)", "Табуляція (\\t)", "Пробіл"])
        
        self.enc_combo = QComboBox()
        self.enc_combo.addItems(["Автовизначення", "utf-8", "utf-8-sig", "windows-1251", "cp866", "latin1"])
        
        layout.addRow("Роздільник:", self.sep_combo)
        layout.addRow("Кодування:", self.enc_combo)
        
        btn_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_box.button(QDialogButtonBox.StandardButton.Ok).setText("ОК")
        btn_box.button(QDialogButtonBox.StandardButton.Cancel).setText("Скасувати")
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addRow(btn_box)
        
    def get_settings(self):
        sep_map = {
            "Автовизначення": None,
            "Кома (,)": ",",
            "Крапка з комою (;)": ";",
            "Табуляція (\\t)": "\t",
            "Пробіл": " "
        }
        enc = self.enc_combo.currentText()
        if enc == "Автовизначення":
            enc = None
            
        return {
            "sep": sep_map[self.sep_combo.currentText()],
            "encoding": enc
        }

class FlowLayout(QLayout):
    """Layout that wraps items to multiple lines like a flow/text wrap"""

    def __init__(self, parent=None, margin=0, spacing=-1):
        super().__init__(parent)
        self.setContentsMargins(margin, margin, margin, margin)
        self.setSpacing(spacing)
        self._item_list = []

    def addItem(self, item):
        self._item_list.append(item)

    def removeItem(self, item):
        self._item_list.remove(item)

    def count(self):
        return len(self._item_list)

    def itemAt(self, index):
        if 0 <= index < len(self._item_list):
            return self._item_list[index]
        return None

    def takeAt(self, index):
        if 0 <= index < len(self._item_list):
            return self._item_list.pop(index)
        return None

    def expandingDirections(self):
        return Qt.Orientation(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        return self._do_layout(QRect(0, 0, width, 0), False)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._do_layout(rect, True)

    def _do_layout(self, rect, dry_run):
        x = rect.x()
        y = rect.y()
        line_height = 0
        spacing = self.spacing()
        margin = self.contentsMargins().left()

        for item in self._item_list:
            widget = item.widget()
            if widget:
                size_hint = widget.sizeHint()
            else:
                size_hint = item.sizeHint()

            available_width = rect.width() - margin * 2
            next_x = x + size_hint.width() + spacing

            if next_x - spacing > rect.right() and line_height > 0:
                x = rect.x()
                y = y + line_height + spacing
                next_x = x + size_hint.width() + spacing
                line_height = 0

            if not dry_run:
                item.setGeometry(QRect(QPoint(x, y), size_hint))

            x = next_x
            line_height = max(line_height, size_hint.height())

        return y + line_height - rect.y() + self.contentsMargins().bottom()

    def minimumSize(self):
        size = QSize()
        for item in self._item_list:
            size = size.expandedTo(item.minimumSize())
        size += QSize(2 * self.contentsMargins().left(), 2 * self.contentsMargins().top())
        return size


class DataToolsPanel(QDialog):
    """Data verification, cleaning and mapping – MS Office ribbon UI"""

    data_changed = pyqtSignal(object)
    mappings_changed = pyqtSignal(dict)

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

    # ─────────────────────────────────────────────────────────────────────────
    # INIT
    # ─────────────────────────────────────────────────────────────────────────
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Перевірка, очищення та співставлення даних")
        self.setGeometry(100, 100, 1200, 800)
        self.df = None
        self.required_columns = []
        self.column_mappings = {}
        self.combo_boxes = {}          # keyed by req_col; rebuilt on every mapping dialog open
        self.filter_widgets = []       # list of {col, cond, val} text dicts
        self.pre_split_df = None
        self.delimiter_map = {
            ", (Кома)": ",",
            "; (Крапка з комою)": ";",
            "\t (Табуляція)": "\t",
            " (Пробіл)": " ",
        }
        self._setup_ui()

    # ─────────────────────────────────────────────────────────────────────────
    # SHARED STYLE HELPERS
    # ─────────────────────────────────────────────────────────────────────────
    @staticmethod
    def _ok_style():
        return ("QPushButton{background:#217346;color:white;border:none;border-radius:3px;"
                "padding:5px 18px;font-size:12px;}"
                "QPushButton:hover{background:#1a5c38;}")

    @staticmethod
    def _btn_style():
        return ("QPushButton{background:#f0f0f0;color:#222;border:1px solid #bbb;"
                "border-radius:3px;padding:5px 14px;font-size:12px;}"
                "QPushButton:hover{background:#e0e0e0;}")

    @staticmethod
    def _row_style():
        return ("QFrame{border:1px solid #e0e0e0;border-radius:3px;background:#f8f9fa;}"
                "QFrame:hover{border-color:#90caf9;background:#fff;}")

    # ─────────────────────────────────────────────────────────────────────────
    # UI SETUP  – MS Office ribbon
    # ─────────────────────────────────────────────────────────────────────────
    def _setup_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── Ribbon tab widget ────────────────────────────────────────────────
        self.ribbon_tabs = QTabWidget()
        self.ribbon_tabs.setMaximumHeight(94)
        self.ribbon_tabs.setStyleSheet("""
            QTabWidget::pane{border:none;border-bottom:1px solid #c0c0c0;background:#fff;}
            QTabBar::tab{background:#f0f0f0;padding:5px 14px;border:1px solid transparent;
                         border-bottom:none;font-size:12px;color:#333;min-width:56px;}
            QTabBar::tab:selected{background:#fff;border-color:#c0c0c0;
                                  border-bottom:2px solid #217346;color:#217346;font-weight:bold;}
            QTabBar::tab:hover:!selected{background:#e3f2fd;}
        """)

        _rb = ("QPushButton{background:transparent;border:1px solid transparent;"
               "border-radius:2px;padding:4px 8px;font-size:12px;color:#222;min-width:60px;}"
               "QPushButton:hover{background:#e3f2fd;border-color:#90caf9;}"
               "QPushButton:pressed{background:#bbdefb;border-color:#1976d2;}")

        def _group(label, *btns):
            """Return a QFrame styled as an Office ribbon group."""
            outer = QFrame()
            outer.setStyleSheet("QFrame{border:none;border-right:1px solid #d0d0d0;background:transparent;}")
            v = QVBoxLayout(outer)
            v.setContentsMargins(4, 3, 4, 2)
            v.setSpacing(2)
            row = QHBoxLayout()
            row.setSpacing(2)
            for b in btns:
                b.setStyleSheet(_rb)
                row.addWidget(b)
            v.addLayout(row)
            lbl = QLabel(label)
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet("font-size:9px;color:#888;border:none;")
            v.addWidget(lbl)
            return outer

        def _tab_row():
            w = QWidget()
            w.setStyleSheet("background:#fff;")
            lay = QHBoxLayout(w)
            lay.setContentsMargins(4, 0, 4, 0)
            lay.setSpacing(0)
            lay.setAlignment(Qt.AlignmentFlag.AlignLeft)
            return w, lay

        # Tab 1 – Home
        ft, fl = _tab_row()
        self.save_btn    = QPushButton(qta.icon('fa5s.save',         color='#217346'), " Зберегти")
        self.save_btn.clicked.connect(self._save_data)
        self.map_btn     = QPushButton(qta.icon('fa5s.exchange-alt', color='#1565c0'), " Співставлення")
        self.map_btn.clicked.connect(self._show_mapping_dialog)
        self.automap_btn = QPushButton(qta.icon('fa5s.magic',        color='#6a1b9a'), " Авто-маппінг")
        self.automap_btn.clicked.connect(self._automap_columns)
        fl.addWidget(_group("Файл",    self.save_btn))
        fl.addWidget(_group("Колонки", self.map_btn, self.automap_btn))
        fl.addStretch()
        self.ribbon_tabs.addTab(ft, "Головна")

        # Tab 2 – Data
        dt, dl = _tab_row()
        self.filter_btn  = QPushButton(qta.icon('fa5s.filter',               color='#e67e22'), " Фільтри")
        self.filter_btn.clicked.connect(self._show_filter_dialog)
        self.apply_btn   = QPushButton(qta.icon('fa5s.check',                color='#27ae60'), " Застосувати")
        self.apply_btn.clicked.connect(self._apply_filters)
        self.clear_btn   = QPushButton(qta.icon('fa5s.times',                color='#c0392b'), " Очистити")
        self.clear_btn.clicked.connect(self._clear_filters)
        self.split_btn   = QPushButton(qta.icon('fa5s.grip-lines-vertical',  color='#e67e22'), " Розділити")
        self.split_btn.clicked.connect(self._show_split_dialog)
        self.merge_btn   = QPushButton(qta.icon('fa5s.compress-arrows-alt',  color='#9b59b6'), " Об'єднати")
        self.merge_btn.clicked.connect(self._show_merge_dialog)
        self.struct_btn  = QPushButton(qta.icon('fa5s.table',                color='#1565c0'), " Структура")
        self.struct_btn.clicked.connect(self._conform_structure)
        dl.addWidget(_group("Фільтрація",    self.filter_btn, self.apply_btn, self.clear_btn))
        dl.addWidget(_group("Трансформації", self.split_btn, self.merge_btn, self.struct_btn))
        dl.addStretch()
        self.ribbon_tabs.addTab(dt, "Дані")

        # Tab 3 – Columns
        ct, cl = _tab_row()
        self.cols_mgr_btn = QPushButton(qta.icon('fa5s.columns', color='#1565c0'), " Менеджер")
        self.cols_mgr_btn.clicked.connect(self._show_columns_dialog)
        self.add_col_btn  = QPushButton(qta.icon('fa5s.plus',    color='#27ae60'), " Додати")
        self.add_col_btn.clicked.connect(self._add_column)
        self.rem_col_btn  = QPushButton(qta.icon('fa5s.trash',   color='#e74c3c'), " Видалити")
        self.rem_col_btn.clicked.connect(self._remove_column)
        cl.addWidget(_group("Менеджер",  self.cols_mgr_btn))
        cl.addWidget(_group("Швидко",    self.add_col_btn, self.rem_col_btn))
        cl.addStretch()
        self.ribbon_tabs.addTab(ct, "Колонки")

        root.addWidget(self.ribbon_tabs)

        # ── Data table ───────────────────────────────────────────────────────
        self.data_table = QTableWidget()
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.data_table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed)
        self.data_table.setStyleSheet("""
            QTableWidget{background:white;gridline-color:#e0e0e0;border:none;}
            QTableWidget::item{padding:4px;}
            QTableWidget::item:selected{background:#1976d2;color:white;}
            QTableWidget::item:alternate{background:#f9f9f9;}
            QHeaderView::section{background:#34495e;color:white;padding:5px 6px;
                                  border:none;border-right:1px solid #4a6278;font-weight:bold;}
        """)
        self.data_table.itemChanged.connect(self._on_cell_changed)

        # ── Context menus on cells and column header ────────────────────────
        self.data_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.data_table.customContextMenuRequested.connect(self._on_table_context_menu)
        hdr = self.data_table.horizontalHeader()
        hdr.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        hdr.customContextMenuRequested.connect(self._on_header_context_menu)
        hdr.sectionClicked.connect(self._on_header_clicked)

        root.addWidget(self.data_table, 1)

        # ── Status bar ───────────────────────────────────────────────────────
        sb = QFrame()
        sb.setFixedHeight(24)
        sb.setStyleSheet("background:#f5f5f5;border-top:1px solid #d0d0d0;")
        sbl = QHBoxLayout(sb)
        sbl.setContentsMargins(10, 0, 10, 0)
        self.status_label        = QLabel("Записів: 0")
        self.filter_status_label = QLabel("")
        self.status_label.setStyleSheet("font-size:12px;color:#444;")
        self.filter_status_label.setStyleSheet("font-size:12px;color:#e67e22;")
        sbl.addWidget(self.status_label)
        sbl.addStretch()
        sbl.addWidget(self.filter_status_label)
        root.addWidget(sb)

    # ─────────────────────────────────────────────────────────────────────────
    # ERROR HANDLING HELPER
    # ─────────────────────────────────────────────────────────────────────────
    def _err(self, title, exc, context=""):
        """Show a styled critical error box with full traceback details."""
        import traceback
        detail = traceback.format_exc()
        msg = QMessageBox(self)
        msg.setWindowTitle(title)
        msg.setIcon(QMessageBox.Icon.Critical)
        msg.setText(f"<b>{context}</b><br><br>{exc}" if context else str(exc))
        msg.setDetailedText(detail)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def _guard(method_name):
        """Decorator: wrap a DataToolsPanel method body in try/except → _err."""
        import functools
        def decorator(fn):
            @functools.wraps(fn)
            def wrapper(self, *a, **kw):
                try:
                    return fn(self, *a, **kw)
                except Exception as exc:
                    self._err(f"Помилка – {method_name}", exc,
                              f"Сталася непередбачена помилка у «{method_name}».")
            return wrapper
        return decorator

    # ─────────────────────────────────────────────────────────────────────────
    # DATA LOADING
    # ─────────────────────────────────────────────────────────────────────────
    def set_data(self, df, required_columns, mappings=None):
        self.df               = df.copy() if df is not None else None
        self.required_columns = required_columns or []
        self.column_mappings  = dict(mappings) if mappings else {}
        self.filter_widgets   = []
        self.combo_boxes      = {}
        self._update_filter_status()
        self._refresh_table()
        self._automap_columns()

    # ─────────────────────────────────────────────────────────────────────────
    # TABLE
    # ─────────────────────────────────────────────────────────────────────────
    def _refresh_table(self):
        try:
            if self.df is None:
                self.data_table.setRowCount(0)
                self.data_table.setColumnCount(0)
                self.status_label.setText("Записів: 0")
                return
            self.data_table.blockSignals(True)
            self.data_table.clear()
            self.data_table.setRowCount(self.df.shape[0])
            self.data_table.setColumnCount(self.df.shape[1])
            self.data_table.setHorizontalHeaderLabels(self.df.columns)
            active_cols = getattr(self, '_active_filtered_cols', set())
            for j, c in enumerate(self.df.columns):
                if c in active_cols:
                    hdr_item = self.data_table.horizontalHeaderItem(j)
                    if hdr_item:
                        hdr_item.setIcon(qta.icon('fa5s.filter', color='#f39c12'))
            for i in range(self.df.shape[0]):
                for j in range(self.df.shape[1]):
                    v = self.df.iloc[i, j]
                    self.data_table.setItem(i, j, QTableWidgetItem(str(v) if pd.notna(v) else ""))
            self.data_table.resizeColumnsToContents()
            self.data_table.blockSignals(False)
            self.status_label.setText(f"Записів: {len(self.df)}")
        except Exception as exc:
            self.data_table.blockSignals(False)
            self._err("Помилка оновлення таблиці", exc, "Не вдалося відобразити дані у таблиці.")

    def _on_cell_changed(self, item):
        if self.df is None:
            return
        try:
            self.df.iloc[item.row(), item.column()] = item.text()
            self.data_changed.emit(self.df)
        except Exception as exc:
            self._err("Помилка редагування комірки", exc,
                      f"Не вдалося оновити значення в рядку {item.row()+1}, колонці {item.column()+1}.")

    # ─────────────────────────────────────────────────────────────────────────
    # FILE ACTIONS
    # ─────────────────────────────────────────────────────────────────────────
    def _save_data(self):
        if self.df is None:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти файл", "", "Excel Files (*.xlsx)")
        if path:
            try:
                self.df.to_excel(path, index=False)
                QMessageBox.information(self, "Успіх", f"Файл збережено:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Помилка", f"Не вдалося зберегти: {e}")

    # ─────────────────────────────────────────────────────────────────────────
    # COLUMN MAPPING
    # ─────────────────────────────────────────────────────────────────────────
    def _show_mapping_dialog(self):
        try:
            self._show_mapping_dialog_impl()
        except Exception as exc:
            self._err("Помилка вікна співставлення", exc,
                      "Не вдалося відкрити або зберегти діалог співставлення колонок.")

    def _show_mapping_dialog_impl(self):
        if self.df is None:
            QMessageBox.warning(self, "Попередження", "Спочатку завантажте дані.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Співставлення колонок")
        dialog.setMinimumSize(820, 520)
        dialog.resize(900, 600)

        lay = QVBoxLayout(dialog)
        lay.setContentsMargins(12, 10, 12, 10)
        lay.setSpacing(6)

        # Column headers
        hdr = QHBoxLayout()
        for txt, stretch in [("Системне поле", 2), ("Колонка у файлі", 3)]:
            lb = QLabel(f"<b>{txt}</b>")
            lb.setStyleSheet("color:#555;font-size:12px;padding:2px 8px;")
            hdr.addWidget(lb, stretch)
        lay.addLayout(hdr)

        # Scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea{border:1px solid #d0d0d0;border-radius:3px;}")
        sw = QWidget()
        sl = QVBoxLayout(sw)
        sl.setContentsMargins(4, 4, 4, 4)
        sl.setSpacing(3)

        self.combo_boxes = {}
        cols = self.df.columns.tolist() if self.df is not None else []

        for req_col in self.required_columns:
            rf = QFrame()
            rf.setFixedHeight(36)
            rf.setStyleSheet(self._row_style())
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
            combo.addItems(cols)
            if req_col in self.column_mappings and self.column_mappings[req_col] in cols:
                combo.setCurrentText(self.column_mappings[req_col])

            combo.currentTextChanged.connect(
                lambda text, rc=req_col: self._on_mapping_changed(rc, text))
            self.combo_boxes[req_col] = combo

            rl.addWidget(lbl, 2)
            rl.addWidget(combo, 3)
            sl.addWidget(rf)

        sl.addStretch()
        scroll.setWidget(sw)
        lay.addWidget(scroll, 1)

        # Buttons
        br = QHBoxLayout()
        br.setSpacing(8)
        am = QPushButton(qta.icon('fa5s.magic', color='#6a1b9a'), " Авто-маппінг")
        am.setStyleSheet(self._btn_style())
        am.clicked.connect(self._automap_columns)
        br.addWidget(am)
        br.addStretch()

        ok  = QPushButton("Зберегти")
        ok.setDefault(True)
        ok.setStyleSheet(self._ok_style())
        ok.clicked.connect(dialog.accept)
        cn  = QPushButton("Скасувати")
        cn.setStyleSheet(self._btn_style())
        cn.clicked.connect(dialog.reject)
        br.addWidget(ok)
        br.addWidget(cn)
        lay.addLayout(br)

        dialog.exec()
        # CRITICAL: the dialog (and all QComboBox widgets inside it) is now
        # destroyed by Qt.  We MUST clear self.combo_boxes before doing
        # anything else, otherwise any access to those C++ objects crashes
        # the process with 0xC0000409 (use-after-free / stack corruption).
        self.combo_boxes = {}
        self.mappings_changed.emit(self.column_mappings)

    def _on_mapping_changed(self, req_col, text):
        if text in ("— Пропустити —", "Пропустити"):
            self.column_mappings.pop(req_col, None)
        else:
            self.column_mappings[req_col] = text
        self.mappings_changed.emit(self.column_mappings)

    def _automap_columns(self):
        if self.df is None:
            return

        cols = self.df.columns.tolist()

        for req in self.required_columns:
            if req in cols:
                self.column_mappings[req] = req
                if req in self.combo_boxes:
                    try:
                        self.combo_boxes[req].blockSignals(True)
                        self.combo_boxes[req].setCurrentText(req)
                        self.combo_boxes[req].blockSignals(False)
                    except RuntimeError:
                        self.combo_boxes = {}
                        break

        self.mappings_changed.emit(self.column_mappings)

    # ─────────────────────────────────────────────────────────────────────────
    # COLUMN ADD / REMOVE
    # ─────────────────────────────────────────────────────────────────────────
    # ─────────────────────────────────────────────────────────────────────────
    # COLUMN CONTEXT MENUS
    # ─────────────────────────────────────────────────────────────────────────
    def _on_table_context_menu(self, pos):
        """Right-click on a table cell → column actions for that column."""
        if self.df is None:
            return
        # pos is relative to data_table; map it to the viewport to find the column under the mouse
        vp_pos = self.data_table.viewport().mapFrom(self.data_table, pos)
        col_idx = self.data_table.columnAt(vp_pos.x())
        if col_idx < 0:
            return
        menu = self._build_col_context_menu(col_idx)
        menu.exec(self.data_table.viewport().mapToGlobal(pos))

    def _on_header_context_menu(self, pos):
        """Right-click on a column header section → column actions."""
        if self.df is None:
            return
        col_idx = self.data_table.horizontalHeader().logicalIndexAt(pos)
        if col_idx < 0:
            return
        menu = self._build_col_context_menu(col_idx)
        menu.exec(self.data_table.horizontalHeader().mapToGlobal(pos))

    def _on_header_clicked(self, logicalIndex):
        """Left-click on a column header section → column actions."""
        if self.df is None or logicalIndex < 0:
            return
        from PyQt6.QtCore import QPoint
        menu = self._build_col_context_menu(logicalIndex)
        hdr = self.data_table.horizontalHeader()
        # Place the menu right below the clicked header section
        x = hdr.sectionViewportPosition(logicalIndex)
        y = hdr.height()
        pos = hdr.viewport().mapToGlobal(QPoint(x, y))
        menu.exec(pos)

    def _build_col_context_menu(self, col_idx: int):
        """Build and return the shared column-actions QMenu."""
        from PyQt6.QtWidgets import QMenu, QInputDialog as _QID
        from PyQt6.QtGui import QAction

        col_name = self.df.columns[col_idx]

        _MENU_STYLE = """
            QMenu {
                background: white;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                padding: 4px 0;
                font-size: 12px;
            }
            QMenu::item {
                padding: 6px 28px 6px 12px;
                color: #2c3e50;
            }
            QMenu::item:selected { background: #e3f2fd; color: #1565c0; }
            QMenu::item:disabled { color: #aaa; }
            QMenu::separator { height: 1px; background: #e0e0e0; margin: 4px 8px; }
        """

        menu = QMenu(self)
        menu.setStyleSheet(_MENU_STYLE)

        # ── Header info ───────────────────────────────────────────────────
        title_act = QAction(f"  📋  «{col_name}»", self)
        title_act.setEnabled(False)
        menu.addAction(title_act)
        menu.addSeparator()

        # ── Rename ────────────────────────────────────────────────────────
        ren_act = QAction(qta.icon('fa5s.pencil-alt', color='#e67e22'), "  Перейменувати колонку", self)
        def _rename():
            new, ok = _QID.getText(self, "Перейменувати", "Нова назва:", text=col_name)
            if not ok or not new.strip(): return
            n = new.strip()
            if n != col_name and n in self.df.columns:
                QMessageBox.warning(self, "Конфлікт", f"Колонка «{n}» вже існує."); return
            try:
                self.df.rename(columns={col_name: n}, inplace=True)
                self.column_mappings = {k: (n if v == col_name else v)
                                        for k, v in self.column_mappings.items()}
                self._refresh_table()
                self.data_changed.emit(self.df)
            except Exception as exc:
                self._err("Перейменування", exc)
        ren_act.triggered.connect(_rename)
        menu.addAction(ren_act)

        # ── Duplicate ─────────────────────────────────────────────────────
        dup_act = QAction(qta.icon('fa5s.copy', color='#8e44ad'), "  Дублювати колонку", self)
        def _duplicate():
            try:
                base = col_name + "_copy"; n = base; i = 2
                while n in self.df.columns: n = f"{base}_{i}"; i += 1
                self.df.insert(col_idx + 1, n, self.df[col_name].copy())
                self._refresh_table()
                self.data_changed.emit(self.df)
            except Exception as exc:
                self._err("Дублювання", exc)
        dup_act.triggered.connect(_duplicate)
        menu.addAction(dup_act)

        menu.addSeparator()

        # ── Add column left / right ───────────────────────────────────────
        add_l_act = QAction(qta.icon('fa5s.arrow-left', color='#27ae60'), "  Додати колонку ліворуч", self)
        def _add_left():
            new, ok = _QID.getText(self, "Нова колонка", "Назва:")
            if not ok or not new.strip(): return
            n = new.strip()
            if n in self.df.columns:
                QMessageBox.warning(self, "Конфлікт", f"Колонка «{n}» вже існує."); return
            try:
                self.df.insert(col_idx, n, "")
                self._refresh_table(); self.data_changed.emit(self.df)
            except Exception as exc:
                self._err("Додавання", exc)
        add_l_act.triggered.connect(_add_left)
        menu.addAction(add_l_act)

        add_r_act = QAction(qta.icon('fa5s.arrow-right', color='#27ae60'), "  Додати колонку праворуч", self)
        def _add_right():
            new, ok = _QID.getText(self, "Нова колонка", "Назва:")
            if not ok or not new.strip(): return
            n = new.strip()
            if n in self.df.columns:
                QMessageBox.warning(self, "Конфлікт", f"Колонка «{n}» вже існує."); return
            try:
                self.df.insert(col_idx + 1, n, "")
                self._refresh_table(); self.data_changed.emit(self.df)
            except Exception as exc:
                self._err("Додавання", exc)
        add_r_act.triggered.connect(_add_right)
        menu.addAction(add_r_act)

        menu.addSeparator()

        # ── Reorder ───────────────────────────────────────────────────────
        move_l_act = QAction(qta.icon('fa5s.chevron-left', color='#1565c0'), "  Перемістити ліворуч", self)
        def _move_left():
            if col_idx <= 0: return
            try:
                cols = list(self.df.columns)
                cols[col_idx - 1], cols[col_idx] = cols[col_idx], cols[col_idx - 1]
                self.df = self.df[cols]
                self._refresh_table(); self.data_changed.emit(self.df)
            except Exception as exc:
                self._err("Переміщення", exc)
        move_l_act.triggered.connect(_move_left)
        move_l_act.setEnabled(col_idx > 0)
        menu.addAction(move_l_act)

        move_r_act = QAction(qta.icon('fa5s.chevron-right', color='#1565c0'), "  Перемістити праворуч", self)
        def _move_right():
            if col_idx >= len(self.df.columns) - 1: return
            try:
                cols = list(self.df.columns)
                cols[col_idx], cols[col_idx + 1] = cols[col_idx + 1], cols[col_idx]
                self.df = self.df[cols]
                self._refresh_table(); self.data_changed.emit(self.df)
            except Exception as exc:
                self._err("Переміщення", exc)
        move_r_act.triggered.connect(_move_right)
        move_r_act.setEnabled(col_idx < len(self.df.columns) - 1)
        menu.addAction(move_r_act)

        menu.addSeparator()

        # ── Delete ────────────────────────────────────────────────────────
        del_act = QAction(qta.icon('fa5s.trash', color='#e74c3c'), "  Видалити колонку", self)
        def _delete():
            reply = QMessageBox.question(
                self, "Видалити колонку",
                f"Видалити колонку «{col_name}»?\nДані в ній буде втрачено безповоротно.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply != QMessageBox.StandardButton.Yes: return
            try:
                self.df.drop(columns=[col_name], inplace=True)
                self.column_mappings = {k: v for k, v in self.column_mappings.items()
                                        if v != col_name}
                self._refresh_table(); self.data_changed.emit(self.df)
            except Exception as exc:
                self._err("Видалення", exc)
        del_act.triggered.connect(_delete)
        menu.addAction(del_act)

        menu.addSeparator()

        # ── Add filter for this column ────────────────────────────────────
        flt_act = QAction(qta.icon('fa5s.filter', color='#e67e22'), "  Додати фільтр для цієї колонки", self)
        def _add_filter(checked=False, col=col_name):
            # Store the target column; the dialog picks it up and adds a focused row
            self._pending_filter_col = col
            self._show_filter_dialog()
        flt_act.triggered.connect(_add_filter)
        menu.addAction(flt_act)

        menu.addSeparator()

        # ── Open Column Manager ───────────────────────────────────────────
        mgr_act = QAction(qta.icon('fa5s.columns', color='#1565c0'), "  Відкрити менеджер колонок…", self)
        mgr_act.triggered.connect(self._show_columns_dialog)
        menu.addAction(mgr_act)

        return menu

    def _add_column(self):
        if self.df is None:
            QMessageBox.warning(self, "Помилка", "Спочатку завантажте дані.")
            return
        try:
            from PyQt6.QtWidgets import QInputDialog
            name, ok = QInputDialog.getText(self, "Додати колонку", "Введіть назву нової колонки:")
            if ok and name:
                name = name.strip()
                if not name:
                    QMessageBox.warning(self, "Помилка", "Назва колонки не може бути порожньою.")
                    return
                if name in self.df.columns:
                    QMessageBox.warning(self, "Помилка", f"Колонка «{name}» вже існує.")
                    return
                self.df[name] = ""
                self._refresh_table()
                self.data_changed.emit(self.df)
        except Exception as exc:
            self._err("Помилка додавання колонки", exc)

    def _remove_column(self):
        if self.df is None:
            QMessageBox.warning(self, "Помилка", "Спочатку завантажте дані.")
            return
        col_idx = self.data_table.currentColumn()
        if col_idx < 0:
            QMessageBox.warning(self, "Помилка", "Клацніть на комірку в колонці, яку хочете видалити.")
            return
        try:
            col_name = self.df.columns[col_idx]
            reply = QMessageBox.question(
                self, "Видалити колонку", f"Видалити колонку «{col_name}»?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.df.drop(columns=[col_name], inplace=True)
                self.column_mappings = {k: v for k, v in self.column_mappings.items() if v != col_name}
                self._refresh_table()
                self.data_changed.emit(self.df)
        except Exception as exc:
            self._err("Помилка видалення колонки", exc)

    # ─────────────────────────────────────────────────────────────────────────
    # FILTER DIALOG  – closure bug fixed via default-arg capture
    # ─────────────────────────────────────────────────────────────────────────
    def _show_filter_dialog(self):
        try:
            self._show_filter_dialog_impl()
        except Exception as exc:
            self._err("Помилка фільтрів", exc,
                      "Не вдалося відкрити діалог фільтрації.")

    def _show_filter_dialog_impl(self):
        if self.df is None:
            QMessageBox.warning(self, "Попередження", "Спочатку завантажте дані.")
            return

        ALL_CONDITIONS = [
            "містить", "не містить",
            "дорівнює", "не дорівнює",
            "починається з", "закінчується на",
            "порожнє", "не порожнє",
            "> (більше)", "< (менше)", ">= (більше рівне)", "<= (менше рівне)",
            "регулярний вираз",
        ]
        NO_VALUE_CONDS = {"порожнє", "не порожнє"}

        dialog = QDialog(self)
        dialog.setWindowTitle("Фільтрація рядків")
        dialog.setMinimumSize(820, 520)
        dialog.resize(880, 580)

        root = QVBoxLayout(dialog)
        root.setContentsMargins(14, 12, 14, 10)
        root.setSpacing(10)

        # ── Title + live count ─────────────────────────────────────────────
        title_row = QHBoxLayout()
        title_lbl = QLabel("Фільтрація рядків")
        title_lbl.setStyleSheet("font-size:14px;font-weight:bold;color:#2c3e50;")
        title_row.addWidget(title_lbl); title_row.addStretch()
        count_lbl = QLabel("")
        count_lbl.setStyleSheet(
            "font-size:11px;color:#555;background:#f0f4f8;"
            "border:1px solid #d0d0d0;border-radius:3px;padding:3px 8px;")
        title_row.addWidget(count_lbl)
        root.addLayout(title_row)
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#e0e0e0;"); root.addWidget(sep)

        # ── AND / OR logic toggle ─────────────────────────────────────────
        logic_row = QHBoxLayout(); logic_row.setSpacing(6)
        logic_row.addWidget(QLabel("Логіка між фільтрами:"))
        logic_and_btn = QPushButton("І (AND)"); logic_or_btn = QPushButton("АБО (OR)")
        for b in (logic_and_btn, logic_or_btn):
            b.setCheckable(True)
            b.setStyleSheet(
                "QPushButton{background:#fff;border:1px solid #bbb;border-radius:3px;"
                "padding:3px 14px;font-size:11px;}"
                "QPushButton:checked{background:#217346;color:white;border-color:#217346;}"
                "QPushButton:hover:!checked{background:#e8f5e9;}")
        logic_and_btn.setChecked(True)
        def _set_and(): logic_or_btn.setChecked(False);  logic_and_btn.setChecked(True);  _update_preview()
        def _set_or():  logic_and_btn.setChecked(False); logic_or_btn.setChecked(True);   _update_preview()
        logic_and_btn.clicked.connect(_set_and); logic_or_btn.clicked.connect(_set_or)
        logic_row.addWidget(logic_and_btn); logic_row.addWidget(logic_or_btn)
        logic_row.addStretch()
        root.addLayout(logic_row)

        # ── Column headers ────────────────────────────────────────────────
        hdr = QFrame(); hdr.setFixedHeight(26)
        hdr.setStyleSheet("background:#f5f5f5;border:1px solid #d0d0d0;border-radius:3px;")
        hl = QHBoxLayout(hdr); hl.setContentsMargins(8, 0, 8, 0)
        for txt, stretch in [("Колонка", 3), ("Умова", 3), ("Значення", 4), ("", 1)]:
            lb = QLabel(f"<b>{txt}</b>" if txt else "")
            lb.setStyleSheet("font-size:11px;color:#555;border:none;background:transparent;")
            hl.addWidget(lb, stretch)
        root.addWidget(hdr)

        # ── Scroll area ───────────────────────────────────────────────────
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea{border:1px solid #d0d0d0;border-radius:3px;}")
        sw = QWidget(); sl = QVBoxLayout(sw)
        sl.setContentsMargins(4, 4, 4, 4); sl.setSpacing(3)
        sl.addStretch()
        filter_rows = []

        def _compute_mask(tmp_df):
            use_and = logic_and_btn.isChecked()
            mask = None
            for r in filter_rows:
                col  = r["col"].currentText()
                cond = r["cond"].currentText()
                val  = r["val"].text()
                if col not in tmp_df.columns: continue
                if cond not in NO_VALUE_CONDS and not val.strip(): continue
                s = tmp_df[col].astype(str)
                if   cond == "містить":            m = s.str.contains(val, case=False, na=False)
                elif cond == "не містить":         m = ~s.str.contains(val, case=False, na=False)
                elif cond == "дорівнює":           m = s.str.lower() == val.lower()
                elif cond == "не дорівнює":        m = s.str.lower() != val.lower()
                elif cond == "починається з":      m = s.str.lower().str.startswith(val.lower())
                elif cond == "закінчується на":    m = s.str.lower().str.endswith(val.lower())
                elif cond == "порожнє":            m = s.str.strip() == ""
                elif cond == "не порожнє":         m = s.str.strip() != ""
                elif cond == "регулярний вираз":
                    try:   m = s.str.contains(val, regex=True, na=False)
                    except: continue
                elif cond == "> (більше)":
                    try:   m = pd.to_numeric(tmp_df[col], errors='coerce') > float(val)
                    except: continue
                elif cond == "< (менше)":
                    try:   m = pd.to_numeric(tmp_df[col], errors='coerce') < float(val)
                    except: continue
                elif cond == ">= (більше рівне)":
                    try:   m = pd.to_numeric(tmp_df[col], errors='coerce') >= float(val)
                    except: continue
                elif cond == "<= (менше рівне)":
                    try:   m = pd.to_numeric(tmp_df[col], errors='coerce') <= float(val)
                    except: continue
                else: continue
                mask = m if mask is None else (mask & m if use_and else mask | m)
            return mask

        def _update_preview(*_):
            if self.df is None: count_lbl.setText(""); return
            try:
                mask = _compute_mask(self.df)
                result_n = int(mask.sum()) if mask is not None else len(self.df)
                total    = len(self.df)
                pct      = 100 * result_n / max(total, 1)
                count_lbl.setText(f"Залишиться: {result_n} з {total} ({pct:.0f}%)")
                color = '#c0392b' if result_n < total else '#27ae60'
                count_lbl.setStyleSheet(
                    f"font-size:11px;color:{color};background:#f0f4f8;"
                    "border:1px solid #d0d0d0;border-radius:3px;padding:3px 8px;")
            except Exception:
                count_lbl.setText("–")

        def _add_row(col_text="", cond_text="містить", val_text=""):
            rf = QFrame(); rf.setFixedHeight(38); rf.setStyleSheet(self._row_style())
            rl = QHBoxLayout(rf); rl.setContentsMargins(6, 2, 6, 2); rl.setSpacing(6)

            cc = QComboBox(); cc.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            cc.setStyleSheet("font-size:12px;"); cc.addItems(self.df.columns)
            if col_text in self.df.columns: cc.setCurrentText(col_text)

            cd = QComboBox(); cd.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            cd.setStyleSheet("font-size:12px;"); cd.addItems(ALL_CONDITIONS)
            if cond_text in ALL_CONDITIONS: cd.setCurrentText(cond_text)

            ve = QLineEdit(val_text); ve.setPlaceholderText("Значення")
            ve.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            ve.setStyleSheet("font-size:12px;")

            db = QPushButton("✕"); db.setFixedSize(24, 24)
            db.setStyleSheet("QPushButton{background:#e74c3c;color:white;border:none;"
                             "border-radius:3px;font-size:11px;}"
                             "QPushButton:hover{background:#c0392b;}")

            rl.addWidget(cc, 3); rl.addWidget(cd, 3); rl.addWidget(ve, 4); rl.addWidget(db, 1)
            row_data = {"frame": rf, "col": cc, "cond": cd, "val": ve}
            filter_rows.append(row_data)
            sl.insertWidget(sl.count() - 1, rf)

            def _on_cond(text, v=ve):
                v.setVisible(text not in NO_VALUE_CONDS)
                _update_preview()
            cd.currentTextChanged.connect(_on_cond)
            cc.currentTextChanged.connect(lambda *_: _update_preview())
            ve.textChanged.connect(lambda *_: _update_preview())
            # Initial hide if needed
            ve.setVisible(cond_text not in NO_VALUE_CONDS)

            def _remove(checked=False, rd=row_data):
                rd["frame"].hide(); rd["frame"].deleteLater()
                if rd in filter_rows: filter_rows.remove(rd)
                _update_preview()
            db.clicked.connect(_remove)

        # Pre-select column if opened via right-click context menu
        pending_col = getattr(self, '_pending_filter_col', None)
        if pending_col:
            self._pending_filter_col = None

        # Restore saved filters
        for fw in self.filter_widgets:
            _add_row(fw.get("col", ""), fw.get("cond", "містить"), fw.get("val", ""))


        # Add the pending column row if requested
        if pending_col:
            _add_row(col_text=pending_col, cond_text="містить", val_text="")
            # Scroll the scroll area to the bottom so the new row is visible
            scroll.verticalScrollBar().setValue(scroll.verticalScrollBar().maximum())

        scroll.setWidget(sw)
        root.addWidget(scroll, 1)

        # ── Bottom bar ────────────────────────────────────────────────────
        br = QHBoxLayout(); br.setSpacing(6)
        ab = QPushButton(qta.icon('fa5s.plus',  color='#27ae60'), " Додати фільтр")
        ab.setStyleSheet(self._btn_style())
        ab.clicked.connect(lambda: (_add_row(), _update_preview()))
        clr_all = QPushButton(qta.icon('fa5s.times', color='#c0392b'), " Очистити всі")
        clr_all.setStyleSheet(self._btn_style())
        def _clear_all():
            for rd in list(filter_rows):
                rd["frame"].hide(); rd["frame"].deleteLater()
            filter_rows.clear(); _update_preview()
        clr_all.clicked.connect(_clear_all)
        br.addWidget(ab); br.addWidget(clr_all); br.addStretch()

        def _save_and_close():
            self.filter_widgets = [
                {"col": r["col"].currentText(),
                 "cond": r["cond"].currentText(),
                 "val":  r["val"].text()}
                for r in filter_rows
            ]
            self._filter_logic_and = logic_and_btn.isChecked()
            self._update_filter_status()
            dialog.accept()

        ok = QPushButton("Зберегти фільтри"); ok.setDefault(True)
        ok.setStyleSheet(self._ok_style()); ok.clicked.connect(_save_and_close)
        cn = QPushButton("Скасувати"); cn.setStyleSheet(self._btn_style())
        cn.clicked.connect(dialog.reject)
        br.addWidget(ok); br.addWidget(cn)
        root.addLayout(br)

        _update_preview()
        dialog.exec()

    # ─────────────────────────────────────────────────────────────────────────
    # COLUMNS MANAGER DIALOG
    # ─────────────────────────────────────────────────────────────────────────
    def _show_columns_dialog(self):
        if self.df is None:
            QMessageBox.warning(self, "Колонки", "Спочатку завантажте дані."); return
        try:
            self._show_columns_dialog_impl()
        except Exception as exc:
            self._err("Помилка менеджера колонок", exc)

    def _show_columns_dialog_impl(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Менеджер колонок")
        dialog.setMinimumSize(560, 560)
        dialog.resize(600, 620)

        root = QVBoxLayout(dialog)
        root.setContentsMargins(14, 12, 14, 10)
        root.setSpacing(10)

        title = QLabel("Керування колонками таблиці")
        title.setStyleSheet("font-size:14px;font-weight:bold;color:#2c3e50;")
        root.addWidget(title)
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#e0e0e0;"); root.addWidget(sep)

        body = QHBoxLayout(); body.setSpacing(8)

        # ── Left: column list ─────────────────────────────────────────────
        list_frame = QFrame()
        list_frame.setStyleSheet("QFrame{border:1px solid #d0d0d0;border-radius:4px;background:white;}")
        lf_lay = QVBoxLayout(list_frame); lf_lay.setContentsMargins(6,6,6,6); lf_lay.setSpacing(4)
        lf_lay.addWidget(QLabel("<b>Колонки (клацніть для вибору)</b>"))

        col_table = QTableWidget()
        col_table.setColumnCount(1); col_table.setShowGrid(False)
        col_table.verticalHeader().setVisible(False)
        col_table.horizontalHeader().setVisible(False)
        col_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        col_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        col_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        col_table.setAlternatingRowColors(True)
        col_table.setStyleSheet(
            "QTableWidget{border:none;font-size:12px;}"
            "QTableWidget::item:selected{background:#1976d2;color:white;}"
            "QTableWidget::item:alternate{background:#f5f5f5;}")

        col_names = list(self.df.columns)

        def _refresh_list(select_idx=None):
            col_table.setRowCount(0)
            for i, c in enumerate(col_names):
                col_table.insertRow(i)
                col_table.setItem(i, 0, QTableWidgetItem(c))
                col_table.setRowHeight(i, 26)
            if select_idx is not None and 0 <= select_idx < len(col_names):
                col_table.selectRow(select_idx)

        _refresh_list()
        lf_lay.addWidget(col_table, 1)
        body.addWidget(list_frame, 1)

        # ── Right: action buttons ─────────────────────────────────────────
        act_w = QWidget(); act_w.setFixedWidth(145)
        act_l = QVBoxLayout(act_w); act_l.setContentsMargins(0,0,0,0); act_l.setSpacing(5)
        act_l.addStretch()

        def _abtn(icon_name, color, label):
            b = QPushButton(qta.icon(icon_name, color=color), f"  {label}")
            b.setStyleSheet(
                "QPushButton{background:#f8f9fa;border:1px solid #d0d0d0;border-radius:3px;"
                "padding:6px 8px;font-size:11px;text-align:left;}"
                "QPushButton:hover{background:#e3f2fd;border-color:#90caf9;}")
            return b

        btn_up  = _abtn('fa5s.arrow-up',          '#1565c0', "Вгору")
        btn_dn  = _abtn('fa5s.arrow-down',         '#1565c0', "Вниз")
        btn_top = _abtn('fa5s.angle-double-up',    '#1565c0', "На початок")
        btn_bot = _abtn('fa5s.angle-double-down',  '#1565c0', "В кінець")
        btn_add = _abtn('fa5s.plus',               '#27ae60', "Додати")
        btn_ren = _abtn('fa5s.pencil-alt',         '#e67e22', "Перейменувати")
        btn_dup = _abtn('fa5s.copy',               '#8e44ad', "Дублювати")
        btn_del = _abtn('fa5s.trash',              '#e74c3c', "Видалити")

        for b in (btn_up, btn_dn, btn_top, btn_bot):
            act_l.addWidget(b)
        act_l.addSpacing(8)
        for b in (btn_add, btn_ren, btn_dup, btn_del):
            act_l.addWidget(b)
        act_l.addStretch()
        body.addWidget(act_w)
        root.addLayout(body, 1)

        def _cur(): return col_table.currentRow()

        def _move(delta):
            r = _cur(); nr = r + delta
            if r < 0 or not (0 <= nr < len(col_names)): return
            col_names[r], col_names[nr] = col_names[nr], col_names[r]; _refresh_list(nr)

        btn_up.clicked.connect(lambda: _move(-1))
        btn_dn.clicked.connect(lambda: _move(1))
        def _to_top():
            r = _cur()
            if r < 0: return
            col_names.insert(0, col_names.pop(r)); _refresh_list(0)
        def _to_bot():
            r = _cur()
            if r < 0: return
            col_names.append(col_names.pop(r)); _refresh_list(len(col_names)-1)
        btn_top.clicked.connect(_to_top); btn_bot.clicked.connect(_to_bot)

        def _add():
            from PyQt6.QtWidgets import QInputDialog
            name, ok = QInputDialog.getText(dialog, "Додати колонку", "Назва нової колонки:")
            if ok and name.strip():
                n = name.strip()
                if n in col_names:
                    QMessageBox.warning(dialog, "Конфлікт", f"Колонка «{n}» вже існує."); return
                col_names.append(n); _refresh_list(len(col_names)-1)

        def _rename():
            r = _cur()
            if r < 0: QMessageBox.warning(dialog, "Помилка", "Виберіть колонку."); return
            from PyQt6.QtWidgets import QInputDialog
            old = col_names[r]
            new, ok = QInputDialog.getText(dialog, "Перейменувати", "Нова назва:", text=old)
            if ok and new.strip():
                n = new.strip()
                if n != old and n in col_names:
                    QMessageBox.warning(dialog, "Конфлікт", f"Колонка «{n}» вже існує."); return
                col_names[r] = n; _refresh_list(r)

        def _duplicate():
            r = _cur()
            if r < 0: QMessageBox.warning(dialog, "Помилка", "Виберіть колонку."); return
            base = col_names[r] + "_copy"; n = base; idx = 2
            while n in col_names: n = f"{base}_{idx}"; idx += 1
            col_names.insert(r + 1, n); _refresh_list(r + 1)

        def _delete():
            r = _cur()
            if r < 0: QMessageBox.warning(dialog, "Помилка", "Виберіть колонку."); return
            if len(col_names) <= 1:
                QMessageBox.warning(dialog, "Помилка", "Не можна видалити єдину колонку."); return
            name = col_names[r]
            reply = QMessageBox.question(dialog, "Видалити",
                f"Видалити колонку «{name}»? Дані в ній буде втрачено.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                col_names.pop(r); _refresh_list(min(r, len(col_names)-1))

        btn_add.clicked.connect(_add); btn_ren.clicked.connect(_rename)
        btn_dup.clicked.connect(_duplicate); btn_del.clicked.connect(_delete)

        # ── Buttons ───────────────────────────────────────────────────────
        br = QHBoxLayout(); br.setSpacing(8); br.addStretch()

        def _apply():
            try:
                old_cols = list(self.df.columns)
                new_df = pd.DataFrame(index=self.df.index)
                for i, new_name in enumerate(col_names):
                    if new_name in self.df.columns:
                        new_df[new_name] = self.df[new_name]
                    elif i < len(old_cols) and old_cols[i] not in col_names:
                        # Same position, different name → treat as rename
                        new_df[new_name] = self.df[old_cols[i]]
                    else:
                        # Newly added or duplicated column
                        src = new_name[:-5] if new_name.endswith("_copy") else None
                        new_df[new_name] = self.df[src].copy() if src and src in self.df.columns else ""
                self.df = new_df
                self.column_mappings = {k: v for k, v in self.column_mappings.items() if v in col_names}
                self._refresh_table()
                self.data_changed.emit(self.df)
                dialog.accept()
            except Exception as exc:
                self._err("Помилка застосування", exc)

        ok_btn = QPushButton(qta.icon('fa5s.check', color='white'), " Застосувати")
        ok_btn.setDefault(True); ok_btn.setStyleSheet(self._ok_style()); ok_btn.clicked.connect(_apply)
        cn_btn = QPushButton("Скасувати"); cn_btn.setStyleSheet(self._btn_style())
        cn_btn.clicked.connect(dialog.reject)
        br.addWidget(ok_btn); br.addWidget(cn_btn)
        root.addLayout(br)
        dialog.exec()

    def _update_filter_status(self):
        n = len([fw for fw in self.filter_widgets if fw.get("val", "")])
        if n > 0:
            self.filter_btn.setText(f" Фільтри ({n})")
            self.filter_status_label.setText(f"Активних фільтрів: {n}")
        else:
            self.filter_btn.setText(" Фільтри")
            self.filter_status_label.setText("")

    def _apply_filters(self):
        if self.df is None:
            QMessageBox.warning(self, "Фільтри", "Спочатку завантажте дані.")
            return

        NO_VALUE_CONDS = {"порожнє", "не порожнє"}
        active = [fw for fw in self.filter_widgets
                  if fw.get("col") and (fw.get("val") or fw.get("cond") in NO_VALUE_CONDS)]
        if not active:
            QMessageBox.information(
                self, "Фільтри",
                "Немає активних фільтрів.\nДодайте фільтри через менеджер або контекстне меню.")
            return

        try:
            use_and = getattr(self, '_filter_logic_and', True)
            tmp = self.df.copy()
            mask = None
            active_cols = set()

            for fw in active:
                col = fw["col"]
                cond = fw.get("cond", "містить")
                val = fw.get("val", "")
                if col not in tmp.columns: continue

                s = tmp[col].astype(str)
                if   cond == "містить":            m = s.str.contains(val, case=False, na=False)
                elif cond == "не містить":         m = ~s.str.contains(val, case=False, na=False)
                elif cond == "дорівнює":           m = s.str.lower() == val.lower()
                elif cond == "не дорівнює":        m = s.str.lower() != val.lower()
                elif cond == "починається з":      m = s.str.lower().str.startswith(val.lower())
                elif cond == "закінчується на":    m = s.str.lower().str.endswith(val.lower())
                elif cond == "порожнє":            m = s.str.strip() == ""
                elif cond == "не порожнє":         m = s.str.strip() != ""
                elif cond == "регулярний вираз":
                    try:   m = s.str.contains(val, regex=True, na=False)
                    except Exception: continue
                elif cond == "> (більше)":
                    try:   m = pd.to_numeric(tmp[col], errors='coerce') > float(val)
                    except: continue
                elif cond == "< (менше)":
                    try:   m = pd.to_numeric(tmp[col], errors='coerce') < float(val)
                    except: continue
                elif cond == ">= (більше рівне)":
                    try:   m = pd.to_numeric(tmp[col], errors='coerce') >= float(val)
                    except: continue
                elif cond == "<= (менше рівне)":
                    try:   m = pd.to_numeric(tmp[col], errors='coerce') <= float(val)
                    except: continue
                else: continue

                mask = m if mask is None else (mask & m if use_and else mask | m)
                active_cols.add(col)

            before = len(tmp)
            self.df = (tmp[mask] if mask is not None else tmp).reset_index(drop=True)
            self._active_filtered_cols = active_cols
            self._refresh_table()
            self.data_changed.emit(self.df)
            QMessageBox.information(self, "Фільтри застосовано", f"Залишилось рядків: {len(self.df)} з {before}")
        except Exception as e:
            QMessageBox.critical(self, "Помилка", f"Помилка при застосуванні фільтрів: {e}")

    def _clear_filters(self):
        self.filter_widgets = []
        self._active_filtered_cols = set()
        self._update_filter_status()
        # Optionally refresh table if we want the 🔍 icons to disappear
        self._refresh_table()

    # ─────────────────────────────────────────────────────────────────────────
    # SPLIT DIALOG  (redesigned)
    # ─────────────────────────────────────────────────────────────────────────
    def _show_split_dialog(self):
        try:
            self._show_split_dialog_impl()
        except Exception as exc:
            self._err("Помилка розділення", exc,
                      "Не вдалося відкрити діалог розділення колонки.")

    def _show_split_dialog_impl(self):
        if self.df is None:
            QMessageBox.warning(self, "Попередження", "Спочатку завантажте дані.")
            return
        cols = self.df.columns.tolist()
        if not cols:
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Розділити колонку")
        dialog.setMinimumSize(620, 520)
        dialog.resize(660, 560)

        root = QVBoxLayout(dialog)
        root.setContentsMargins(16, 14, 16, 12)
        root.setSpacing(12)

        # ── Title ─────────────────────────────────────────────────────────
        title = QLabel("Розділити колонку на кілька")
        title.setStyleSheet("font-size:14px;font-weight:bold;color:#2c3e50;")
        root.addWidget(title)
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#e0e0e0;"); root.addWidget(sep)

        # ── Step 1 – source column ─────────────────────────────────────────
        s1 = QFrame()
        s1.setStyleSheet("QFrame{background:#f5f5f5;border:1px solid #e0e0e0;border-radius:4px;}")
        s1l = QHBoxLayout(s1); s1l.setContentsMargins(10, 8, 10, 8); s1l.setSpacing(10)
        s1l.addWidget(QLabel("Колонка для розділення:"))
        col_combo = QComboBox()
        col_combo.addItems(cols)
        col_combo.setStyleSheet("font-size:12px;min-width:180px;")
        s1l.addWidget(col_combo, 1)

        remove_cb = QCheckBox("Видалити вихідну")
        remove_cb.setChecked(True)
        remove_cb.setStyleSheet("font-size:11px;")
        s1l.addWidget(remove_cb)

        root.addWidget(s1)

        # ── Step 2 – delimiter ────────────────────────────────────────────
        s2 = QFrame()
        s2.setStyleSheet(s1.styleSheet())
        s2l = QHBoxLayout(s2); s2l.setContentsMargins(10, 8, 10, 8); s2l.setSpacing(10)
        s2l.addWidget(QLabel("Роздільник:"))
        delim_combo = QComboBox()
        delim_combo.setStyleSheet("font-size:12px;")
        delim_combo.addItems(list(self.delimiter_map.keys()) + ["Інший..."])
        s2l.addWidget(delim_combo)
        custom_edit = QLineEdit()
        custom_edit.setPlaceholderText("Введіть символ…")
        custom_edit.setVisible(False)
        custom_edit.setMaximumWidth(140)
        custom_edit.setStyleSheet("font-size:12px;")
        s2l.addWidget(custom_edit)
        delim_combo.currentTextChanged.connect(lambda t: custom_edit.setVisible(t == "Інший..."))
        root.addWidget(s2)

        # ── Step 3 – output column names ──────────────────────────────────
        names_group = QFrame()
        names_group.setStyleSheet("QFrame{border:1px solid #d0d0d0;border-radius:4px;background:white;}")
        ng_lay = QVBoxLayout(names_group); ng_lay.setContentsMargins(10, 8, 10, 8); ng_lay.setSpacing(6)

        ng_hdr = QHBoxLayout()
        ng_hdr.addWidget(QLabel("<b>Назви нових колонок</b>"))
        ng_hdr.addStretch()
        add_part_btn = QPushButton(qta.icon('fa5s.plus', color='#27ae60'), " Додати")
        add_part_btn.setStyleSheet(self._btn_style())
        add_part_btn.setFixedHeight(26)
        ng_hdr.addWidget(add_part_btn)
        ng_lay.addLayout(ng_hdr)

        parts_widget = QWidget()
        parts_lay = QVBoxLayout(parts_widget)
        parts_lay.setContentsMargins(0, 0, 0, 0)
        parts_lay.setSpacing(4)
        name_edits = []   # list of QLineEdit

        def _add_name_row(text=""):
            idx = len(name_edits)
            row_f = QFrame()
            row_f.setFixedHeight(34)
            row_f.setStyleSheet("QFrame{border:1px solid #e8e8e8;border-radius:3px;background:#f8f9fa;}")
            row_l = QHBoxLayout(row_f); row_l.setContentsMargins(6, 2, 6, 2); row_l.setSpacing(6)

            num_lbl = QLabel(f"#{idx + 1}")
            num_lbl.setFixedWidth(24)
            num_lbl.setStyleSheet("color:#888;font-size:11px;border:none;background:transparent;")

            edit = QLineEdit(text)
            edit.setPlaceholderText(f"Назва частини {idx + 1}")
            edit.setStyleSheet("font-size:12px;border:none;background:transparent;")
            # IMPORTANT: use lambda so _update_preview is looked up at CALL TIME
            # (not at connection time – it isn't defined yet when _add_name_row is first called)
            edit.textChanged.connect(lambda *_: _update_preview())

            del_btn = QPushButton("✕")
            del_btn.setFixedSize(22, 22)
            del_btn.setStyleSheet("QPushButton{background:#e74c3c;color:white;border:none;"
                                  "border-radius:3px;font-size:10px;}"
                                  "QPushButton:hover{background:#c0392b;}")

            row_l.addWidget(num_lbl)
            row_l.addWidget(edit, 1)
            row_l.addWidget(del_btn)
            parts_lay.addWidget(row_f)
            name_edits.append(edit)

            def _remove(checked=False, rf=row_f, ed=edit):
                if ed in name_edits:
                    name_edits.remove(ed)
                rf.hide(); rf.deleteLater()
                # Same deferred lookup pattern
                _update_preview()

            del_btn.clicked.connect(_remove)

        _add_name_row("Частина 1")
        _add_name_row("Частина 2")
        add_part_btn.clicked.connect(lambda: (_add_name_row(), _update_preview()))

        ng_lay.addWidget(parts_widget)
        root.addWidget(names_group)

        # ── Live preview ──────────────────────────────────────────────────
        prev_group = QFrame()
        prev_group.setStyleSheet("QFrame{border:1px solid #d0d0d0;border-radius:4px;background:white;}")
        pg_lay = QVBoxLayout(prev_group); pg_lay.setContentsMargins(10, 8, 10, 8); pg_lay.setSpacing(6)
        pg_lay.addWidget(QLabel("<b>Попередній перегляд (перші 3 рядки)</b>"))
        preview_table = QTableWidget()
        preview_table.setMaximumHeight(120)
        preview_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        preview_table.setAlternatingRowColors(True)
        preview_table.verticalHeader().setVisible(False)
        preview_table.setStyleSheet(
            "QTableWidget{border:none;font-size:11px;gridline-color:#eee;}"
            "QHeaderView::section{background:#34495e;color:white;padding:3px 6px;"
            "border:none;border-right:1px solid #4a6278;font-size:11px;font-weight:bold;}")
        pg_lay.addWidget(preview_table)
        root.addWidget(prev_group)

        def _get_delim():
            t = delim_combo.currentText()
            return custom_edit.text() if t == "Інший..." else self.delimiter_map.get(t, "")

        def _update_preview(*_):
            col   = col_combo.currentText()
            delim = _get_delim()
            names = [e.text() or f"Частина {i+1}" for i, e in enumerate(name_edits)]
            if not col or not delim or not names:
                preview_table.setRowCount(0); preview_table.setColumnCount(0)
                return
            try:
                sample = self.df[col].head(3).astype(str)
                split_data = sample.str.split(delim, n=len(names) - 1, expand=True)
                preview_table.setColumnCount(len(names))
                preview_table.setHorizontalHeaderLabels(names)
                preview_table.setRowCount(min(3, len(sample)))
                header = preview_table.horizontalHeader()
                for j in range(len(names)):
                    header.setSectionResizeMode(j, QHeaderView.ResizeMode.Stretch)
                for i in range(min(3, len(sample))):
                    for j, name in enumerate(names):
                        val = split_data.iloc[i, j] if j < split_data.shape[1] else ""
                        preview_table.setItem(i, j, QTableWidgetItem(str(val) if pd.notna(val) else ""))
            except Exception:
                preview_table.setRowCount(0); preview_table.setColumnCount(0)

        col_combo.currentTextChanged.connect(_update_preview)
        delim_combo.currentTextChanged.connect(_update_preview)
        custom_edit.textChanged.connect(_update_preview)
        _update_preview()

        # ── Buttons ───────────────────────────────────────────────────────
        br = QHBoxLayout(); br.setSpacing(8); br.addStretch()

        def _do():
            col_split = col_combo.currentText()
            delim     = _get_delim()
            names     = [e.text().strip() for e in name_edits if e.text().strip()]
            if not delim:
                QMessageBox.warning(dialog, "Помилка", "Вкажіть роздільник."); return
            if not names:
                QMessageBox.warning(dialog, "Помилка", "Додайте щонайменше одну назву."); return
            if len(names) != len(set(names)):
                QMessageBox.warning(dialog, "Дублікати", "Назви нових колонок повинні бути унікальними."); return
            for n in names:
                if n in self.df.columns and n != col_split:
                    QMessageBox.warning(dialog, "Конфлікт", f"Колонка «{n}» вже існує."); return
            try:
                self.pre_split_df = self.df.copy()
                idx        = self.df.columns.get_loc(col_split)
                split_data = self.df[col_split].astype(str).str.split(delim, n=len(names) - 1, expand=True)
                
                if remove_cb.isChecked():
                    self.df.drop(columns=[col_split], inplace=True)
                    self.column_mappings = {k: v for k, v in self.column_mappings.items() if v != col_split}
                else:
                    idx += 1  # Insert AFTER the original column if we keep it

                for i, n in reversed(list(enumerate(names))):
                    col_data = split_data[i].fillna('') if i < split_data.shape[1] else pd.Series([''] * len(self.df))
                    self.df.insert(idx, n, col_data.values)
                self._refresh_table()
                self.data_changed.emit(self.df)
                dialog.accept()
                QMessageBox.information(
                    self, "Успіх",
                    f"Колонку «{col_split}» розділено на {len(names)} частини:\n" +
                    "\n".join(f"  • {n}" for n in names))
            except Exception as e:
                QMessageBox.critical(dialog, "Помилка", f"Сталася помилка: {e}")

        ok_btn = QPushButton(qta.icon('fa5s.grip-lines-vertical', color='white'), " Розділити")
        ok_btn.setDefault(True); ok_btn.setStyleSheet(self._ok_style()); ok_btn.clicked.connect(_do)
        cn_btn = QPushButton("Скасувати"); cn_btn.setStyleSheet(self._btn_style()); cn_btn.clicked.connect(dialog.reject)
        br.addWidget(ok_btn); br.addWidget(cn_btn)
        root.addLayout(br)
        dialog.exec()


    # ─────────────────────────────────────────────────────────────────────────
    # MERGE DIALOG  (redesigned)
    # ─────────────────────────────────────────────────────────────────────────
    def _show_merge_dialog(self):
        try:
            self._show_merge_dialog_impl()
        except Exception as exc:
            self._err("Помилка об'єднання", exc,
                      "Не вдалося відкрити діалог об'єднання колонок.")

    def _show_merge_dialog_impl(self):
        if self.df is None:
            QMessageBox.warning(self, "Попередження", "Спочатку завантажте дані.")
            return
        cols = self.df.columns.tolist()

        dialog = QDialog(self)
        dialog.setWindowTitle("Об'єднати колонки")
        dialog.setMinimumSize(660, 560)
        dialog.resize(700, 600)

        root = QVBoxLayout(dialog)
        root.setContentsMargins(16, 14, 16, 12)
        root.setSpacing(12)

        # ── Title ──────────────────────────────────────────────────────────
        title = QLabel("Об'єднати кілька колонок в одну")
        title.setStyleSheet("font-size:14px;font-weight:bold;color:#2c3e50;")
        root.addWidget(title)
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#e0e0e0;"); root.addWidget(sep)

        # ── Column picker panel ────────────────────────────────────────────
        picker_frame = QFrame()
        picker_frame.setStyleSheet("QFrame{border:1px solid #d0d0d0;border-radius:4px;background:white;}")
        pf_lay = QVBoxLayout(picker_frame); pf_lay.setContentsMargins(10, 8, 10, 8); pf_lay.setSpacing(6)
        pf_lay.addWidget(QLabel("<b>Оберіть колонки для об'єднання:</b>"))

        # Two-column layout: available | selected
        picker_body = QHBoxLayout(); picker_body.setSpacing(8)

        # Left – available
        avail_w = QWidget()
        avail_l = QVBoxLayout(avail_w); avail_l.setContentsMargins(0,0,0,0); avail_l.setSpacing(4)
        avail_l.addWidget(QLabel("Доступні колонки:"))
        avail_list = QTableWidget()
        avail_list.setColumnCount(1); avail_list.setShowGrid(False)
        avail_list.verticalHeader().setVisible(False)
        avail_list.horizontalHeader().setVisible(False)
        avail_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        avail_list.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        avail_list.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        avail_list.setAlternatingRowColors(True)
        avail_list.setStyleSheet("QTableWidget{border:1px solid #d0d0d0;border-radius:3px;font-size:12px;}"
                                 "QTableWidget::item:alternate{background:#f5f5f5;}")
        avail_list.setRowCount(len(cols))
        for i, c in enumerate(cols):
            avail_list.setItem(i, 0, QTableWidgetItem(c))
            avail_list.setRowHeight(i, 26)
        avail_l.addWidget(avail_list, 1)

        # Middle – arrow buttons
        mid_w = QWidget(); mid_w.setFixedWidth(44)
        mid_l = QVBoxLayout(mid_w); mid_l.setContentsMargins(0,0,0,0); mid_l.setSpacing(6)
        mid_l.addStretch()
        add_arrow   = QPushButton("→"); add_arrow.setFixedSize(36, 28)
        rem_arrow   = QPushButton("←"); rem_arrow.setFixedSize(36, 28)
        for b in (add_arrow, rem_arrow):
            b.setStyleSheet("QPushButton{background:#217346;color:white;border:none;"
                            "border-radius:3px;font-size:14px;font-weight:bold;}"
                            "QPushButton:hover{background:#1a5c38;}")
        mid_l.addWidget(add_arrow); mid_l.addWidget(rem_arrow)
        mid_l.addStretch()

        # Right – selected + up/down reorder
        sel_w = QWidget()
        sel_l = QVBoxLayout(sel_w); sel_l.setContentsMargins(0,0,0,0); sel_l.setSpacing(4)
        sel_l.addWidget(QLabel("Колонки для об'єднання (у порядку):"))

        sel_body = QHBoxLayout(); sel_body.setSpacing(4)
        sel_list = QTableWidget()
        sel_list.setColumnCount(1); sel_list.setShowGrid(False)
        sel_list.verticalHeader().setVisible(False)
        sel_list.horizontalHeader().setVisible(False)
        sel_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        sel_list.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        sel_list.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        sel_list.setAlternatingRowColors(True)
        sel_list.setStyleSheet(avail_list.styleSheet())
        sel_body.addWidget(sel_list, 1)

        ord_btns = QWidget(); ord_btns.setFixedWidth(30)
        ord_l = QVBoxLayout(ord_btns); ord_l.setContentsMargins(0,0,0,0); ord_l.setSpacing(4)
        ord_l.addStretch()
        up_btn  = QPushButton("↑"); up_btn.setFixedSize(26, 26)
        dn_btn  = QPushButton("↓"); dn_btn.setFixedSize(26, 26)
        for b in (up_btn, dn_btn):
            b.setStyleSheet("QPushButton{background:#ecf0f1;border:1px solid #bbb;"
                            "border-radius:3px;font-size:13px;}"
                            "QPushButton:hover{background:#d0d0d0;}")
        ord_l.addWidget(up_btn); ord_l.addWidget(dn_btn)
        ord_l.addStretch()
        sel_body.addWidget(ord_btns)
        sel_l.addLayout(sel_body, 1)

        picker_body.addWidget(avail_w, 1)
        picker_body.addWidget(mid_w)
        picker_body.addWidget(sel_w, 1)
        pf_lay.addLayout(picker_body, 1)
        root.addWidget(picker_frame, 2)

        selected_cols = []   # ordered list of column names

        def _refresh_sel_list():
            sel_list.setRowCount(0)
            for c in selected_cols:
                r = sel_list.rowCount(); sel_list.insertRow(r)
                sel_list.setItem(r, 0, QTableWidgetItem(c))
                sel_list.setRowHeight(r, 26)
            _update_preview()

        def _add_col():
            row = avail_list.currentRow()
            if row < 0: return
            name = avail_list.item(row, 0).text()
            if name not in selected_cols:
                selected_cols.append(name)
                _refresh_sel_list()

        def _rem_col():
            row = sel_list.currentRow()
            if row < 0: return
            selected_cols.pop(row)
            _refresh_sel_list()

        def _move_up():
            row = sel_list.currentRow()
            if row > 0:
                selected_cols[row - 1], selected_cols[row] = selected_cols[row], selected_cols[row - 1]
                _refresh_sel_list()
                sel_list.selectRow(row - 1)

        def _move_down():
            row = sel_list.currentRow()
            if row >= 0 and row < len(selected_cols) - 1:
                selected_cols[row], selected_cols[row + 1] = selected_cols[row + 1], selected_cols[row]
                _refresh_sel_list()
                sel_list.selectRow(row + 1)

        add_arrow.clicked.connect(_add_col)
        rem_arrow.clicked.connect(_rem_col)
        up_btn.clicked.connect(_move_up)
        dn_btn.clicked.connect(_move_down)
        # Double-click to add/remove
        avail_list.cellDoubleClicked.connect(lambda r, c: _add_col())
        sel_list.cellDoubleClicked.connect(lambda r, c: _rem_col())

        # ── Settings row ───────────────────────────────────────────────────
        cfg_frame = QFrame()
        cfg_frame.setStyleSheet("QFrame{background:#f5f5f5;border:1px solid #e0e0e0;border-radius:4px;}")
        cfg_lay = QHBoxLayout(cfg_frame); cfg_lay.setContentsMargins(10, 8, 10, 8); cfg_lay.setSpacing(16)

        # Delimiter
        cfg_lay.addWidget(QLabel("Роздільник:"))
        delim_btns = QHBoxLayout(); delim_btns.setSpacing(3)
        _delims = [("_  пробіл", " "), (",  кома", ", "), (";  крапка", "; "), ("Інший…", None)]
        delim_group = []
        custom_delim_edit = QLineEdit(); custom_delim_edit.setFixedWidth(54)
        custom_delim_edit.setStyleSheet("font-size:12px;"); custom_delim_edit.setVisible(False)
        custom_delim_edit.setPlaceholderText("…")
        # NOTE: actual textChanged connection is made below after _update_preview is defined

        def _make_delim_btn(label, val):
            btn = QPushButton(label); btn.setCheckable(True)
            btn.setStyleSheet("QPushButton{background:#fff;border:1px solid #bbb;border-radius:3px;"
                              "padding:2px 7px;font-size:11px;min-width:0px;}"
                              "QPushButton:checked{background:#217346;color:white;border-color:#217346;}"
                              "QPushButton:hover:!checked{background:#e8f5e9;}")
            btn.clicked.connect(lambda checked, v=val, b=btn: _select_delim(b, v))
            delim_group.append((btn, val))
            delim_btns.addWidget(btn)
            return btn

        current_delim = [" "]   # mutable cell
        def _select_delim(clicked_btn, val):
            for b, v in delim_group:
                b.setChecked(b is clicked_btn)
            if val is None:
                custom_delim_edit.setVisible(True)
                current_delim[0] = custom_delim_edit.text()
            else:
                custom_delim_edit.setVisible(False)
                current_delim[0] = val
            _update_preview()

        custom_delim_edit.textChanged.connect(lambda t: current_delim.__setitem__(0, t) or _update_preview())

        first_btn = None
        for label, val in _delims:
            b = _make_delim_btn(label, val)
            if first_btn is None:
                first_btn = b
                b.setChecked(True)
        delim_btns.addWidget(custom_delim_edit)
        cfg_lay.addLayout(delim_btns)
        cfg_lay.addStretch()

        # New name
        cfg_lay.addWidget(QLabel("Нова назва:"))
        new_name_edit = QLineEdit(); new_name_edit.setPlaceholderText("ПІБ")
        new_name_edit.setFixedWidth(140); new_name_edit.setStyleSheet("font-size:12px;")
        new_name_edit.textChanged.connect(lambda *a: _update_preview())
        cfg_lay.addWidget(new_name_edit)

        # Delete originals
        del_chk = QCheckBox("Видалити вихідні"); del_chk.setChecked(True)
        cfg_lay.addWidget(del_chk)
        root.addWidget(cfg_frame)

        # ── Live preview ───────────────────────────────────────────────────
        prev_frame = QFrame()
        prev_frame.setStyleSheet("QFrame{border:1px solid #d0d0d0;border-radius:4px;background:white;}")
        prev_l = QVBoxLayout(prev_frame); prev_l.setContentsMargins(10, 8, 10, 8); prev_l.setSpacing(4)
        prev_l.addWidget(QLabel("<b>Попередній перегляд (перший рядок)</b>"))
        preview_label = QLabel("—")
        preview_label.setStyleSheet("font-size:12px;color:#2c3e50;background:#f0f4f8;"
                                    "border:1px solid #d0d0d0;border-radius:3px;padding:6px 10px;")
        preview_label.setWordWrap(True)
        prev_l.addWidget(preview_label)
        root.addWidget(prev_frame)

        def _update_preview(*_):
            if not selected_cols or self.df is None or self.df.empty:
                preview_label.setText("—"); return
            delim = current_delim[0]
            try:
                row = self.df[selected_cols].iloc[0].astype(str)
                result = delim.join(row.values)
                name   = new_name_edit.text().strip() or "нова_колонка"
                preview_label.setText(f"<b>{name}:</b>  {result}")
            except Exception as e:
                preview_label.setText(f"<i>Помилка: {e}</i>")

        # Connect custom_delim_edit now that _update_preview is defined
        custom_delim_edit.textChanged.connect(
            lambda t: (current_delim.__setitem__(0, t), _update_preview()))

        # ── Buttons ────────────────────────────────────────────────────────
        br = QHBoxLayout(); br.setSpacing(8); br.addStretch()

        def _do():
            if len(selected_cols) < 2:
                QMessageBox.warning(dialog, "Помилка",
                                    "Оберіть щонайменше дві колонки для об'єднання."); return
            new_name = new_name_edit.text().strip()
            if not new_name:
                QMessageBox.warning(dialog, "Помилка", "Введіть назву нової колонки."); return
            if new_name in self.df.columns and new_name not in selected_cols:
                QMessageBox.warning(dialog, "Конфлікт",
                                    f"Колонка «{new_name}» вже існує у таблиці."); return
            delim = current_delim[0]
            try:
                self.df[new_name] = self.df[selected_cols].astype(str).agg(delim.join, axis=1)
                if del_chk.isChecked():
                    drop = [c for c in selected_cols if c != new_name]
                    if drop:
                        self.df.drop(columns=drop, inplace=True)
                self._refresh_table()
                self.data_changed.emit(self.df)
                dialog.accept()
                QMessageBox.information(
                    self, "Успіх",
                    f"Колонки об'єднано в «{new_name}»:\n" +
                    "\n".join(f"  • {c}" for c in selected_cols))
            except Exception as e:
                QMessageBox.critical(dialog, "Помилка", f"Сталася помилка: {e}")

        ok_btn = QPushButton(qta.icon('fa5s.compress-arrows-alt', color='white'), " Об'єднати")
        ok_btn.setDefault(True); ok_btn.setStyleSheet(self._ok_style()); ok_btn.clicked.connect(_do)
        cn_btn = QPushButton("Скасувати"); cn_btn.setStyleSheet(self._btn_style()); cn_btn.clicked.connect(dialog.reject)
        br.addWidget(ok_btn); br.addWidget(cn_btn)
        root.addLayout(br)
        dialog.exec()

    # ─────────────────────────────────────────────────────────────────────────
    # CONFORM STRUCTURE
    # ─────────────────────────────────────────────────────────────────────────
    def _conform_structure(self):
        if self.df is None:
            QMessageBox.warning(self, "Структура", "Спочатку завантажте дані.")
            return

        # ── Compute diff ────────────────────────────────────────────────────
        current = set(self.df.columns)
        ideal   = set(self.IDEAL_COLUMN_ORDER)
        matched = [c for c in self.IDEAL_COLUMN_ORDER if c in current]
        missing = [c for c in self.IDEAL_COLUMN_ORDER if c not in current]
        extra   = [c for c in self.df.columns if c not in ideal]

        # ── Dialog ─────────────────────────────────────────────────────────
        dialog = QDialog(self)
        dialog.setWindowTitle("Сформувати ідеальну структуру таблиці")
        dialog.setMinimumSize(720, 560)
        dialog.resize(780, 600)

        root = QVBoxLayout(dialog)
        root.setContentsMargins(16, 14, 16, 12)
        root.setSpacing(12)

        # Title
        title = QLabel("Привести таблицю до ідеальної структури")
        title.setStyleSheet("font-size:14px;font-weight:bold;color:#2c3e50;")
        root.addWidget(title)
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#e0e0e0;"); root.addWidget(sep)

        # Summary bar
        summary = QLabel(
            f"<b>Поточних колонок:</b> {len(self.df.columns)} &nbsp;│&nbsp; "
            f"<b style='color:#27ae60'>Збережено:</b> {len(matched)} &nbsp;│&nbsp; "
            f"<b style='color:#2980b9'>Додається (порожні):</b> {len(missing)} &nbsp;│&nbsp; "
            f"<b style='color:#c0392b'>Видаляється (дані буде втрачено!):</b> {len(extra)}"
        )
        summary.setWordWrap(True)
        summary.setStyleSheet("font-size:12px;background:#f8f9fa;border:1px solid #d0d0d0;"
                              "border-radius:4px;padding:8px 10px;")
        root.addWidget(summary)

        # 3-panel diff layout
        panels_row = QHBoxLayout(); panels_row.setSpacing(10)

        def _make_panel(title_text, color, items, icon_char, item_style=""):
            frame = QFrame()
            frame.setStyleSheet(f"QFrame{{border:1px solid {color}33;border-radius:5px;"
                                f"background:{color}08;}}")
            lay = QVBoxLayout(frame); lay.setContentsMargins(8, 8, 8, 8); lay.setSpacing(4)

            hdr = QLabel(f"<b>{icon_char} {title_text}</b> ({len(items)})")
            hdr.setStyleSheet(f"font-size:12px;color:{color};border:none;background:transparent;")
            lay.addWidget(hdr)

            sep2 = QFrame(); sep2.setFrameShape(QFrame.Shape.HLine)
            sep2.setStyleSheet(f"color:{color}44;"); lay.addWidget(sep2)

            list_w = QTableWidget()
            list_w.setColumnCount(1); list_w.setShowGrid(False)
            list_w.verticalHeader().setVisible(False)
            list_w.horizontalHeader().setVisible(False)
            list_w.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            list_w.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            list_w.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
            list_w.setAlternatingRowColors(True)
            list_w.setStyleSheet(
                f"QTableWidget{{border:none;font-size:11px;background:transparent;}}"
                f"QTableWidget::item{{padding:2px 4px;color:#2c3e50;}}"
                f"QTableWidget::item:alternate{{background:{color}0d;}}")
            list_w.setRowCount(len(items))
            for i, col in enumerate(items):
                item = QTableWidgetItem(col)
                if item_style:
                    item.setForeground(__import__('PyQt6.QtGui', fromlist=['QColor']).QColor(color))
                list_w.setItem(i, 0, item)
                list_w.setRowHeight(i, 22)
            lay.addWidget(list_w, 1)
            return frame

        panels_row.addWidget(_make_panel("Збережено", "#27ae60", matched, "✓"), 1)
        panels_row.addWidget(_make_panel("Додається (порожні)", "#2980b9", missing, "+", "colored"), 1)
        panels_row.addWidget(_make_panel("Видаляється ⚠", "#c0392b", extra, "✕", "colored"), 1)
        root.addLayout(panels_row, 1)

        # Warning if data loss
        if extra:
            warn = QLabel(
                f"⚠️  <b>Увага!</b>  Буде видалено {len(extra)} колонк{'у' if len(extra)==1 else 'и' if len(extra)<5 else 'ок'} "
                f"з даними. Цю дію <b>не можна скасувати</b>."
            )
            warn.setWordWrap(True)
            warn.setStyleSheet("font-size:12px;color:#c0392b;background:#fdf3f3;"
                               "border:1px solid #f5c6c6;border-radius:4px;padding:8px 10px;")
            root.addWidget(warn)

        # Buttons
        br = QHBoxLayout(); br.setSpacing(8); br.addStretch()

        def _apply():
            try:
                new_df = pd.DataFrame(index=self.df.index)
                for col in self.IDEAL_COLUMN_ORDER:
                    new_df[col] = self.df[col] if col in self.df.columns else ""
                self.df = new_df
                self._refresh_table()
                self._automap_columns()
                self.data_changed.emit(self.df)
                dialog.accept()
                QMessageBox.information(
                    self, "Структура оновлена",
                    f"Таблицю приведено до ідеального формату.\n"
                    f"• Збережено:  {len(matched)} колонок\n"
                    f"• Додано:     {len(missing)} порожніх колонок\n"
                    f"• Видалено:   {len(extra)} колонок")
            except Exception as exc:
                self._err("Помилка структури", exc, "Не вдалося застосувати ідеальну структуру.")

        apply_btn = QPushButton(qta.icon('fa5s.check', color='white'), " Застосувати")
        apply_btn.setDefault(True)
        apply_btn.setStyleSheet(self._ok_style())
        apply_btn.clicked.connect(_apply)

        cancel_btn = QPushButton("Скасувати")
        cancel_btn.setStyleSheet(self._btn_style())
        cancel_btn.clicked.connect(dialog.reject)

        br.addWidget(apply_btn)
        br.addWidget(cancel_btn)
        root.addLayout(br)
        dialog.exec()

    # ─────────────────────────────────────────────────────────────────────────
    # GETTERS
    # ─────────────────────────────────────────────────────────────────────────
    def get_mappings(self):
        return self.column_mappings

    def get_data(self):
        return self.df


class AddTemplatesDialog(QDialog):
    """MS Office-style dialog for adding templates – card-style choice buttons."""

    _CARD_STYLE = """
        QPushButton#card {
            background: #f8f9fa;
            border: 2px solid #d0d0d0;
            border-radius: 8px;
            min-width: 190px;
            min-height: 150px;
            padding: 16px;
        }
        QPushButton#card:hover {
            background: #e8f5e9;
            border-color: #217346;
        }
        QPushButton#card:pressed {
            background: #c8e6c9;
            border-color: #155724;
        }
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Додати шаблони")
        self.setMinimumWidth(480)
        self.setMinimumHeight(280)
        self.choice = None

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 14)
        root.setSpacing(14)

        # Title
        title = QLabel("Як ви хочете додати нові шаблони?")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size:14px;font-weight:bold;color:#2c3e50;")
        root.addWidget(title)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#e0e0e0;")
        root.addWidget(sep)

        # Card buttons
        cards_row = QHBoxLayout()
        cards_row.setSpacing(14)

        def _make_card(icon_name, icon_color, title_text, desc_text):
            btn = QPushButton()
            btn.setObjectName("card")
            btn.setStyleSheet(self._CARD_STYLE)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            lay = QVBoxLayout(btn)
            lay.setSpacing(6)
            lay.setAlignment(Qt.AlignmentFlag.AlignCenter)

            icon_lbl = QLabel()
            icon_lbl.setPixmap(qta.icon(icon_name, color=icon_color).pixmap(QSize(48, 48)))
            icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            icon_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

            t_lbl = QLabel(title_text)
            t_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            t_lbl.setStyleSheet("font-size:13px;font-weight:bold;color:#2c3e50;")
            t_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

            d_lbl = QLabel(desc_text)
            d_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            d_lbl.setWordWrap(True)
            d_lbl.setStyleSheet("font-size:11px;color:#666;")
            d_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

            lay.addWidget(icon_lbl)
            lay.addWidget(t_lbl)
            lay.addWidget(d_lbl)
            return btn

        btn_files = _make_card('fa5s.file-alt', '#217346',
                               "Додати файли",
                               "Вибрати один або декілька .docx файлів")
        btn_folder = _make_card('fa5s.folder-open', '#e67e22',
                                "Сканувати папку",
                                "Імпортувати всі .docx з папки та підпапок")

        btn_files.clicked.connect(self._select_files)
        btn_folder.clicked.connect(self._select_folder)

        cards_row.addWidget(btn_files, 1)
        cards_row.addWidget(btn_folder, 1)
        root.addLayout(cards_row, 1)

        # Cancel
        cancel_row = QHBoxLayout()
        cancel_row.addStretch()
        cancel_btn = QPushButton("Скасувати")
        cancel_btn.setStyleSheet(
            "QPushButton{background:#f0f0f0;color:#222;border:1px solid #bbb;"
            "border-radius:3px;padding:6px 20px;font-size:12px;}"
            "QPushButton:hover{background:#e0e0e0;}")
        cancel_btn.clicked.connect(self.reject)
        cancel_row.addWidget(cancel_btn)
        root.addLayout(cancel_row)

    def _select_files(self):
        self.choice = 'files'
        self.accept()

    def _select_folder(self):
        self.choice = 'folder'
        self.accept()

class MainWindow(QMainWindow):
    select_excel_file_signal = pyqtSignal()
    sheet_changed_signal = pyqtSignal(str)
    map_columns_signal = pyqtSignal()
    refresh_templates_signal = pyqtSignal()
    template_category_changed_signal = pyqtSignal(str)
    choose_custom_templates_signal = pyqtSignal()
    unselect_all_templates_signal = pyqtSignal()
    configure_scores_signal = pyqtSignal()
    generate_documents_signal = pyqtSignal()
    data_tools_updated_signal = pyqtSignal(object)  # Emits dataframe
    open_template_signal = pyqtSignal(str)
    open_template_dir_signal = pyqtSignal(str)

    def changeEvent(self, event):
        from PyQt6.QtCore import QEvent
        if event.type() == QEvent.Type.ActivationChange and self.isActiveWindow():
            self.set_active_redacting_template("") # clear highlight when returning to app
        super().changeEvent(event)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Document Generator")
        self.setGeometry(100, 100, 1200, 750)
        self.log_color_map = {'INFO': '#2ecc71', 'WARNING': '#f39c12', 'ERROR': '#e74c3c', 'PROCESS': '#3498db', 'DEFAULT': '#ecf0f1'}
        self._setup_ui()
        self._initial_ui_state()
        self.update_configured_scores_display([]) # Set initial state

    def _initial_ui_state(self):
        self.sheet_dropdown.setEnabled(False)
        self.set_sheet_loaded_state(False)

    def set_file_loaded_state(self, loaded):
        self.sheet_dropdown.setEnabled(loaded)
        if not loaded: self.update_sheet_dropdown([])

    def set_sheet_loaded_state(self, loaded):
        self.map_columns_button.setEnabled(loaded)
        self.template_group.setEnabled(loaded)
        self.settings_group.setEnabled(loaded)
        is_scores_checked = self.include_scores_checkbox.isChecked()
        self.configure_scores_button.setEnabled(loaded and is_scores_checked)

    def _setup_ui(self):
        central_widget = QWidget()
        central_widget.setObjectName("MainContentWidget")
        self.setCentralWidget(central_widget)
        
        root_layout = QVBoxLayout(central_widget)
        root_layout.setContentsMargins(8, 8, 8, 8)
        root_layout.setSpacing(8)

        main_panels_layout = QHBoxLayout()
        root_layout.addLayout(main_panels_layout)

        left_panel = QWidget(); right_panel = QWidget()
        main_panels_layout.addWidget(left_panel, 3) 
        main_panels_layout.addWidget(right_panel, 5)
        
        left_layout = QVBoxLayout(left_panel)
        left_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        left_layout.setSpacing(8)

        self.source_group = QGroupBox("Крок 1: Джерело даних")
        source_layout = QVBoxLayout(self.source_group); source_layout.setSpacing(6)
        self.select_excel_button = QPushButton(qta.icon('fa5s.file-excel', color='#27ae60'), " Виберіть файл (Excel/CSV)")
        self.sheet_dropdown = QComboBox()
        self.map_columns_button = QPushButton(qta.icon('fa5s.columns', color='#1565c0'), " Співставлення колонок")
        source_layout.addWidget(self.select_excel_button)
        source_layout.addWidget(QLabel("Виберіть аркуш:"))
        source_layout.addWidget(self.sheet_dropdown)
        source_layout.addWidget(self.map_columns_button)
        left_layout.addWidget(self.source_group)
        
        self.template_group = QGroupBox("Крок 2: Шаблони документів")
        template_layout = QVBoxLayout(self.template_group); template_layout.setSpacing(6)
        
        self.template_tree = QTreeWidget()
        self.template_tree.setHeaderHidden(True)
        self.template_tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.template_tree.setIndentation(16)
        
        category_layout = QHBoxLayout()
        self.template_category_dropdown = QComboBox()
        self.refresh_templates_button = QPushButton(qta.icon('fa5s.sync-alt'), "")
        self.refresh_templates_button.setObjectName("RefreshButton")
        category_layout.addWidget(QLabel("Категорія:")); category_layout.addWidget(self.template_category_dropdown, 1); category_layout.addWidget(self.refresh_templates_button)
        template_buttons_layout = QHBoxLayout()
        self.choose_custom_templates_button = QPushButton(qta.icon('fa5s.plus', color='#27ae60'), " Додати")
        self.unselect_all_templates_button = QPushButton(qta.icon('fa5s.square', color='#7f8c8d'), " Зняти виділення")
        template_buttons_layout.addWidget(self.choose_custom_templates_button)
        template_buttons_layout.addWidget(self.unselect_all_templates_button)
        template_layout.addLayout(category_layout)
        template_layout.addWidget(self.template_tree)
        template_layout.addLayout(template_buttons_layout)
        left_layout.addWidget(self.template_group)

        self.settings_group = QGroupBox("Крок 3: Налаштування та Генерація")
        settings_layout = QVBoxLayout(self.settings_group); settings_layout.setSpacing(8)
        
        self.include_scores_checkbox = QCheckBox("Додати бали з таблиці")
        self.configure_scores_button = QPushButton(qta.icon('fa5s.cog', color='#8e44ad'), " Налаштувати колонки балів")
        
        self.configured_scores_table = QTableWidget()
        self.configured_scores_table.setColumnCount(3)
        self.configured_scores_table.setHorizontalHeaderLabels(["Колонка", "Ключ (число)", "Ключ (прописом)"])
        self.configured_scores_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.configured_scores_table.verticalHeader().setVisible(False)
        self.configured_scores_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.configured_scores_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.configured_scores_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.configured_scores_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        
        self.generate_button = QPushButton(qta.icon('fa5s.cogs', color='#217346'), " Створити документи"); self.generate_button.setObjectName("GenerateButton")
        
        settings_layout.addWidget(self.include_scores_checkbox)
        settings_layout.addWidget(self.configure_scores_button)
        settings_layout.addWidget(self.configured_scores_table)
        settings_layout.addStretch()
        settings_layout.addWidget(self.generate_button)
        left_layout.addWidget(self.settings_group)
        
        right_layout = QVBoxLayout(right_panel)
        self.tabs = QTabWidget()
        self.preview_table = QTableWidget()
        self.log_window = QTextEdit(); self.log_window.setReadOnly(True)

        self.tabs.addTab(self.preview_table, qta.icon('fa5s.table'), "Попередній перегляд")
        self.tabs.addTab(self.log_window, qta.icon('fa5s.terminal'), "Консоль")
        right_layout.addWidget(self.tabs)

        self.select_excel_button.clicked.connect(self.select_excel_file_signal.emit)
        self.sheet_dropdown.currentTextChanged.connect(self.sheet_changed_signal.emit)
        self.map_columns_button.clicked.connect(self.map_columns_signal.emit)
        self.refresh_templates_button.clicked.connect(self.refresh_templates_signal.emit)
        self.template_category_dropdown.currentTextChanged.connect(self.template_category_changed_signal.emit)
        self.choose_custom_templates_button.clicked.connect(self.choose_custom_templates_signal.emit)
        self.unselect_all_templates_button.clicked.connect(self.template_tree.clearSelection)
        
        self.template_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.template_tree.customContextMenuRequested.connect(self._on_template_context_menu)
        self.include_scores_checkbox.toggled.connect(self.configure_scores_button.setEnabled)
        self.configure_scores_button.clicked.connect(self.configure_scores_signal.emit)
        self.generate_button.clicked.connect(self.generate_documents_signal.emit)

    def show_add_templates_dialog(self):
        dialog = AddTemplatesDialog(self)
        return dialog.choice if dialog.exec() else None

    def _on_template_context_menu(self, pos):
        item = self.template_tree.itemAt(pos)
        if not item: return
        path_data = item.data(0, Qt.ItemDataRole.UserRole)
        if not path_data: return # It's a folder, no context menu
        
        from PyQt6.QtWidgets import QMenu
        from PyQt6.QtGui import QAction
        import qtawesome as qta
        menu = QMenu(self)
        menu.setStyleSheet("QMenu { background: white; border: 1px solid #d0d0d0; } QMenu::item { padding: 4px 20px; }")
        
        open_act = QAction(qta.icon('fa5s.external-link-alt', color='#217346'), " Відкрити/Редагувати шаблон", self)
        open_act.triggered.connect(lambda: self.open_template_signal.emit(path_data))
        menu.addAction(open_act)
        
        dir_act = QAction(qta.icon('fa5s.folder-open', color='#e67e22'), " Відкрити папку з шаблоном", self)
        dir_act.triggered.connect(lambda: self.open_template_dir_signal.emit(path_data))
        menu.addAction(dir_act)
        
        menu.exec(self.template_tree.viewport().mapToGlobal(pos))
        
    def set_active_redacting_template(self, template_name: str):
        """Highlight the currently editing template in the tree."""
        if not template_name:
            # Quick clear for all
            self.template_tree.clearSelection() # just in case, but usually we just want to reset icons/bg
        from PyQt6.QtGui import QColor
        
        def traverse(node):
            for i in range(node.childCount()):
                item = node.child(i)
                path_data = item.data(0, Qt.ItemDataRole.UserRole)
                if path_data:
                    if path_data == template_name:
                        item.setBackground(0, QColor('#fff3cd'))
                        item.setForeground(0, QColor('#856404'))
                        item.setIcon(0, qta.icon('fa5s.edit', color='#856404'))
                    else:
                        item.setBackground(0, Qt.GlobalColor.transparent)
                        item.setForeground(0, Qt.GlobalColor.black)
                        item.setIcon(0, qta.icon('fa5s.file-word', color='#217346'))
                traverse(item)
                
        traverse(self.template_tree.invisibleRootItem())
    
    def log_message(self, message, level='INFO'):
        color = self.log_color_map.get(level.upper(), self.log_color_map['DEFAULT'])
        timestamp = datetime.now().strftime("%H:%M:%S")
        html_message = f'<font color="{color}"><b>[{timestamp} - {level.upper()}]:</b> {message}</font>'
        self.log_window.append(html_message)

    def update_sheet_dropdown(self, sheet_names):
        self.sheet_dropdown.blockSignals(True)
        self.sheet_dropdown.clear()
        self.sheet_dropdown.addItem("...")
        if sheet_names: self.sheet_dropdown.addItems(sheet_names)
        self.sheet_dropdown.blockSignals(False)

    def update_template_categories_dropdown(self, category_names):
        self.template_category_dropdown.blockSignals(True)
        current_selection = self.template_category_dropdown.currentText()
        self.template_category_dropdown.clear()
        if category_names:
            self.template_category_dropdown.addItems(sorted(category_names))
        if current_selection in category_names:
            self.template_category_dropdown.setCurrentText(current_selection)
        self.template_category_dropdown.blockSignals(False)
        self.template_category_changed_signal.emit(self.template_category_dropdown.currentText())

    def update_preview_table(self, df):
        self.preview_table.setRowCount(0); self.preview_table.setColumnCount(0)
        if df is None or df.empty: return
        self.preview_table.setRowCount(df.shape[0])
        self.preview_table.setColumnCount(df.shape[1])
        self.preview_table.setHorizontalHeaderLabels(df.columns)
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                self.preview_table.setItem(i, j, QTableWidgetItem(str(df.iloc[i, j])))
        self.preview_table.resizeColumnsToContents()

    def update_configured_scores_display(self, mappings):
        self.configured_scores_table.setRowCount(0)

        if not mappings:
            self.configured_scores_table.horizontalHeader().setVisible(False)
            self.configured_scores_table.setRowCount(1)
            self.configured_scores_table.setColumnCount(1)
            placeholder = QTableWidgetItem("Не налаштовано жодної колонки.")
            placeholder.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.configured_scores_table.setSpan(0, 0, 1, 3)
            self.configured_scores_table.setItem(0, 0, placeholder)
            self.configured_scores_table.setColumnCount(3)
            return

        self.configured_scores_table.horizontalHeader().setVisible(True)
        self.configured_scores_table.setColumnCount(3)
        self.configured_scores_table.setRowCount(len(mappings))
        
        for row, m in enumerate(mappings):
            source_item = QTableWidgetItem(m['source'])
            key_item = QTableWidgetItem(m['key'])
            written_key_text = m['written_key'] if m.get('add_written', False) else "---"
            written_key_item = QTableWidgetItem(written_key_text)
            if not m.get('add_written', False):
                written_key_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            self.configured_scores_table.setItem(row, 0, source_item)
            self.configured_scores_table.setItem(row, 1, key_item)
            self.configured_scores_table.setItem(row, 2, written_key_item)

    def populate_template_list(self, template_paths):
        self.template_tree.clear()
        folders = {}
        for path_str in template_paths:
            parts = path_str.replace('\\', '/').split('/')
            if len(parts) == 1:
                item = QTreeWidgetItem(self.template_tree)
                item.setText(0, parts[0])
                item.setIcon(0, qta.icon('fa5s.file-word', color='#217346'))
                item.setData(0, Qt.ItemDataRole.UserRole, path_str)
            else:
                parent_node = self.template_tree.invisibleRootItem()
                for i in range(len(parts)-1):
                    folder_name = parts[i]
                    folder_path = "/".join(parts[:i+1])
                    if folder_path not in folders:
                        f_item = QTreeWidgetItem(parent_node)
                        f_item.setText(0, folder_name)
                        f_item.setIcon(0, qta.icon('fa5s.folder', color='#e67e22'))
                        f_item.setFlags(f_item.flags() & ~Qt.ItemFlag.ItemIsSelectable) # No select
                        folders[folder_path] = f_item
                    parent_node = folders[folder_path]
                item = QTreeWidgetItem(parent_node)
                item.setText(0, parts[-1])
                item.setIcon(0, qta.icon('fa5s.file-word', color='#217346'))
                item.setData(0, Qt.ItemDataRole.UserRole, path_str)
        self.template_tree.expandAll()

    # === Data Tools Panel Methods ===
    def _on_data_tools_data_changed(self, df=None):
        """Handle data changes from the integrated data tools panel"""
        if df is None:
            df = self.data_tools_panel.get_data()
        if df is not None:
            self.update_preview_table(df)
            # Emit signal for backend to update its df
            self.data_tools_updated_signal.emit(df)

    def _on_data_tools_mappings_changed(self, mappings):
        """Handle mapping changes from the integrated data tools panel"""
        # Could emit signal to backend if needed
        pass

    def show_data_tools_panel(self, df, required_columns, mappings=None):
        """Open data tools window with provided data"""
        try:
            # Create new window instance (non-modal)
            self.data_tools_window = DataToolsPanel(self)
            self.data_tools_window.data_changed.connect(self._on_data_tools_data_changed)
            self.data_tools_window.mappings_changed.connect(self._on_data_tools_mappings_changed)
            self.data_tools_window.set_data(df, required_columns, mappings)
            self.data_tools_window.show()
            self.data_tools_window.raise_()
            self.data_tools_window.activateWindow()
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Помилка", f"Не вдалося відкрити вікно: {e}")