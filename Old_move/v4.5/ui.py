from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
                             QListWidget, QComboBox, QTableWidget, QTableWidgetItem, QCheckBox,
                             QFileDialog, QDialog, QDialogButtonBox, QTextEdit,
                             QGroupBox, QTabWidget, QAbstractItemView, QHeaderView)
from PyQt6.QtCore import Qt, pyqtSignal, QSize
from datetime import datetime
import qtawesome as qta
from column_mapper import ColumnMappingDialog

class AddTemplatesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Додати шаблони")
        self.setMinimumWidth(500)
        self.choice = None
        main_layout = QVBoxLayout(self)
        title_label = QLabel("Як ви хочете додати нові шаблони?")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setObjectName("DialogTitle")
        buttons_layout = QHBoxLayout()
        main_layout.addWidget(title_label)
        main_layout.addLayout(buttons_layout)
        btn_add_files = QPushButton()
        btn_add_files.setObjectName("ChoiceButton")
        btn_add_files.setCursor(Qt.CursorShape.PointingHandCursor)
        layout_files = QVBoxLayout(btn_add_files)
        icon_files_qicon = qta.icon('fa5s.file-import', color='#ecf0f1')
        icon_label_files = QLabel()
        icon_label_files.setPixmap(icon_files_qicon.pixmap(QSize(64, 64)))
        label_title_files = QLabel("Додати файли")
        label_desc_files = QLabel("Вибрати один або декілька .docx файлів.")
        label_title_files.setStyleSheet("color: #ecf0f1; font-size: 11pt; font-weight: bold; background-color: transparent;")
        label_desc_files.setStyleSheet("color: #bdc3c7; font-size: 9pt; background-color: transparent;")
        layout_files.addWidget(icon_label_files, alignment=Qt.AlignmentFlag.AlignCenter)
        layout_files.addWidget(label_title_files, alignment=Qt.AlignmentFlag.AlignCenter)
        layout_files.addWidget(label_desc_files, alignment=Qt.AlignmentFlag.AlignCenter)
        btn_scan_folder = QPushButton()
        btn_scan_folder.setObjectName("ChoiceButton")
        btn_scan_folder.setCursor(Qt.CursorShape.PointingHandCursor)
        layout_folder = QVBoxLayout(btn_scan_folder)
        icon_folder_qicon = qta.icon('fa5s.folder-plus', color='#ecf0f1')
        icon_label_folder = QLabel()
        icon_label_folder.setPixmap(icon_folder_qicon.pixmap(QSize(64, 64)))
        label_title_folder = QLabel("Сканувати папку")
        label_desc_folder = QLabel("Імпортувати всі .docx файли з папки та її підпапок.")
        label_title_folder.setStyleSheet("color: #ecf0f1; font-size: 11pt; font-weight: bold; background-color: transparent;")
        label_desc_folder.setStyleSheet("color: #bdc3c7; font-size: 9pt; background-color: transparent;")
        layout_folder.addWidget(icon_label_folder, alignment=Qt.AlignmentFlag.AlignCenter)
        layout_folder.addWidget(label_title_folder, alignment=Qt.AlignmentFlag.AlignCenter)
        layout_folder.addWidget(label_desc_folder, alignment=Qt.AlignmentFlag.AlignCenter)
        buttons_layout.addWidget(btn_add_files)
        buttons_layout.addWidget(btn_scan_folder)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel)
        button_box.rejected.connect(self.reject)
        main_layout.addWidget(button_box)
        btn_add_files.clicked.connect(self._select_files)
        btn_scan_folder.clicked.connect(self._select_folder)
    def _select_files(self): self.choice = 'files'; self.accept()
    def _select_folder(self): self.choice = 'folder'; self.accept()

class MainWindow(QMainWindow):
    open_settings_signal = pyqtSignal()
    select_excel_file_signal = pyqtSignal()
    sheet_changed_signal = pyqtSignal(str)
    map_columns_signal = pyqtSignal()
    refresh_templates_signal = pyqtSignal()
    template_category_changed_signal = pyqtSignal(str)
    choose_custom_templates_signal = pyqtSignal()
    unselect_all_templates_signal = pyqtSignal()
    configure_scores_signal = pyqtSignal()
    generate_documents_signal = pyqtSignal()

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

        top_bar = QHBoxLayout()
        top_bar.addStretch()
        self.settings_button = QPushButton(qta.icon('fa5s.cog'), "")
        self.settings_button.setObjectName("SettingsButton")
        top_bar.addWidget(self.settings_button)
        root_layout.addLayout(top_bar)

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
        self.select_excel_button = QPushButton(qta.icon('fa5s.file-excel', color='white'), " Виберіть Excel файл")
        self.sheet_dropdown = QComboBox()
        self.map_columns_button = QPushButton(qta.icon('fa5s.columns', color='white'), " Співставлення колонок")
        source_layout.addWidget(self.select_excel_button)
        source_layout.addWidget(QLabel("Виберіть аркуш:"))
        source_layout.addWidget(self.sheet_dropdown)
        source_layout.addWidget(self.map_columns_button)
        left_layout.addWidget(self.source_group)
        
        self.template_group = QGroupBox("Крок 2: Шаблони документів")
        template_layout = QVBoxLayout(self.template_group); template_layout.setSpacing(6)
        self.template_listbox = QListWidget(); self.template_listbox.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        category_layout = QHBoxLayout()
        self.template_category_dropdown = QComboBox()
        self.refresh_templates_button = QPushButton(qta.icon('fa5s.sync-alt'), "")
        self.refresh_templates_button.setObjectName("RefreshButton")
        category_layout.addWidget(QLabel("Категорія:")); category_layout.addWidget(self.template_category_dropdown, 1); category_layout.addWidget(self.refresh_templates_button)
        template_buttons_layout = QHBoxLayout()
        self.choose_custom_templates_button = QPushButton(qta.icon('fa5s.plus', color='white'), " Додати")
        self.unselect_all_templates_button = QPushButton(qta.icon('fa5s.square', color='white'), " Зняти виділення")
        template_buttons_layout.addWidget(self.choose_custom_templates_button)
        template_buttons_layout.addWidget(self.unselect_all_templates_button)
        template_layout.addLayout(category_layout)
        template_layout.addWidget(self.template_listbox)
        template_layout.addLayout(template_buttons_layout)
        left_layout.addWidget(self.template_group)

        self.settings_group = QGroupBox("Крок 3: Налаштування та Генерація")
        settings_layout = QVBoxLayout(self.settings_group); settings_layout.setSpacing(8)
        
        self.include_scores_checkbox = QCheckBox("Додати бали з таблиці")
        self.configure_scores_button = QPushButton(qta.icon('fa5s.check-square', color='white'), " Налаштувати колонки балів")
        
        # --- KEY CHANGE: Replaced QListWidget with QTableWidget ---
        self.configured_scores_table = QTableWidget()
        self.configured_scores_table.setColumnCount(3)
        self.configured_scores_table.setHorizontalHeaderLabels(["Колонка", "Ключ (число)", "Ключ (прописом)"])
        self.configured_scores_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.configured_scores_table.verticalHeader().setVisible(False)
        self.configured_scores_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.configured_scores_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.configured_scores_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.configured_scores_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        # --- End of change ---
        
        self.generate_button = QPushButton(qta.icon('fa5s.cogs', color='white'), " Створити документи"); self.generate_button.setObjectName("GenerateButton")
        
        settings_layout.addWidget(self.include_scores_checkbox)
        settings_layout.addWidget(self.configure_scores_button)
        settings_layout.addWidget(self.configured_scores_table) # Add the new table
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
        
        self.settings_button.clicked.connect(self.open_settings_signal.emit)
        self.select_excel_button.clicked.connect(self.select_excel_file_signal.emit)
        self.sheet_dropdown.currentTextChanged.connect(self.sheet_changed_signal.emit)
        self.map_columns_button.clicked.connect(self.map_columns_signal.emit)
        self.refresh_templates_button.clicked.connect(self.refresh_templates_signal.emit)
        self.template_category_dropdown.currentTextChanged.connect(self.template_category_changed_signal.emit)
        self.choose_custom_templates_button.clicked.connect(self.choose_custom_templates_signal.emit)
        self.unselect_all_templates_button.clicked.connect(self.unselect_all_templates_signal.emit)
        self.include_scores_checkbox.toggled.connect(self.configure_scores_button.setEnabled)
        self.configure_scores_button.clicked.connect(self.configure_scores_signal.emit)
        self.generate_button.clicked.connect(self.generate_documents_signal.emit)

    def show_add_templates_dialog(self):
        dialog = AddTemplatesDialog(self)
        return dialog.choice if dialog.exec() else None
    
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
            placeholder = QTableWidgetItem("Не налаштовано жодної колонки.")
            placeholder.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.configured_scores_table.setSpan(0, 0, 1, 3)
            self.configured_scores_table.setItem(0, 0, placeholder)
            return

        self.configured_scores_table.horizontalHeader().setVisible(True)
        self.configured_scores_table.setRowCount(len(mappings))
        for row, m in enumerate(mappings):
            self.configured_scores_table.setSpan(row, 0, 1, 1) # Reset spans
            self.configured_scores_table.setSpan(row, 1, 1, 1)
            self.configured_scores_table.setSpan(row, 2, 1, 1)

            source_item = QTableWidgetItem(m['source'])
            key_item = QTableWidgetItem(m['key'])
            
            written_key_text = m['written_key'] if m.get('add_written', False) else "---"
            written_key_item = QTableWidgetItem(written_key_text)
            if not m.get('add_written', False):
                written_key_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            self.configured_scores_table.setItem(row, 0, source_item)
            self.configured_scores_table.setItem(row, 1, key_item)
            self.configured_scores_table.setItem(row, 2, written_key_item)

    def populate_template_list(self, template_names):
        self.template_listbox.clear()
        self.template_listbox.addItems(template_names)