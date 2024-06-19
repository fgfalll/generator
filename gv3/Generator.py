import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, \
    QListWidget, \
    QComboBox, QTableWidget, QTableWidgetItem, QCheckBox, QMessageBox, QFileDialog, QDialog, QDialogButtonBox, \
    QListWidgetItem, QTextEdit, QScrollArea, QFormLayout
from PyQt5.QtCore import Qt
from docxtpl import DocxTemplate
from num2words import num2words
from babel.dates import format_datetime
from datetime import datetime
import pandas as pd
import pathlib

class ScoreColumnSelectorDialog(QDialog):
    def __init__(self, excel_path, selected_sheet, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Виберіть поля з оцінками")
        self.selected_columns = []

        # Read Excel file into a DataFrame
        try:
            df = pd.read_excel(excel_path, sheet_name=selected_sheet)
            numeric_columns = df.select_dtypes(include=['int', 'float']).columns.tolist()

            layout = QVBoxLayout(self)

            self.list_widget = QListWidget(self)
            self.list_widget.setSelectionMode(QListWidget.MultiSelection)

            for column in numeric_columns:
                item = QListWidgetItem(column)
                self.list_widget.addItem(item)

            layout.addWidget(self.list_widget)

            button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
            button_box.accepted.connect(self.accept)
            button_box.rejected.connect(self.reject)
            layout.addWidget(button_box)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Під час читання файлу Excel сталася помилка: {e}")

    def accept(self):
        self.selected_columns = [item.text() for item in self.list_widget.selectedItems()]
        super().accept()

class ColumnMappingDialog(QDialog):
    def __init__(self, excel_columns, required_columns, parent=None):
        super(ColumnMappingDialog, self).__init__(parent)
        self.setWindowTitle("Map Columns")
        self.mappings = {}

        # Main layout
        main_layout = QVBoxLayout()

        # Scroll Area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scrollContent = QWidget(scroll)
        scrollLayout = QVBoxLayout(scrollContent)

        # Create form layout for mappings
        form_layout = QFormLayout()

        self.combo_boxes = {}
        for required_column in required_columns:
            label = QLabel(required_column)
            combo_box = QComboBox()
            combo_box.addItem("Skip")  # Add Skip option
            combo_box.addItems(excel_columns)
            form_layout.addRow(label, combo_box)
            self.combo_boxes[required_column] = combo_box

        scrollLayout.addLayout(form_layout)
        scroll.setWidget(scrollContent)
        main_layout.addWidget(scroll)

        # OK and Cancel buttons
        buttons_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Відмінити")
        buttons_layout.addWidget(ok_button)
        buttons_layout.addWidget(cancel_button)
        main_layout.addLayout(buttons_layout)

        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)

        self.setLayout(main_layout)

    def accept(self):
        try:
            for required_column, combo_box in self.combo_boxes.items():
                selected_column = combo_box.currentText()
                if selected_column != "Skip":
                    self.mappings[required_column] = selected_column
            super(ColumnMappingDialog, self).accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while saving mappings: {e}")

class DocumentGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Generator V3")
        self.setGeometry(100, 100, 700, 600)  # Set initial window size

        self.excel_path = None
        self.word_templates = []  # List to store multiple selected templates
        self.sheet_name = None
        self.include_scores = False
        self.selected_score_columns = []
        self.column_mappings = {}
        self.output_dir = pathlib.Path(__file__).resolve().parents[1] / "Вихід"
        self.example_dir = pathlib.Path(__file__).resolve().parents[0] / "Приклади"
        self.expected_sheets = []  # Initially empty

        self.setup_gui()

    def setup_gui(self):
        central_widget = QWidget()
        central_widget.setAcceptDrops(True)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        # Excel file selection
        select_excel_button = QPushButton("Виберіть Excel файл", self)
        select_excel_button.clicked.connect(self.select_excel_file)
        layout.addWidget(select_excel_button)

        # Dropdown for sheet selection
        sheet_label = QLabel("Виберіть лист", self)
        layout.addWidget(sheet_label)
        self.sheet_dropdown = QComboBox(self)
        self.sheet_dropdown.addItem("Виберіть лист")
        self.sheet_dropdown.currentIndexChanged.connect(self.update_preview_from_sheet)
        layout.addWidget(self.sheet_dropdown)

        # Preview label
        preview_label = QLabel("Попередній перегляд вибраного файлу Excel:", self)
        layout.addWidget(preview_label)

        # Table widget for preview
        self.preview_table = QTableWidget(self)
        layout.addWidget(self.preview_table)

        # Button to open column mapping dialog
        map_columns_button = QPushButton("Співставлення колонок", self)
        map_columns_button.clicked.connect(self.map_columns)
        layout.addWidget(map_columns_button)

        # Template selection
        template_label = QLabel("Виберіть Шаблони документів", self)
        layout.addWidget(template_label)

        self.template_listbox = QListWidget(self)
        self.template_listbox.setSelectionMode(QListWidget.MultiSelection)
        self.template_listbox.setMaximumHeight(150)
        templates = [
            "Анкета для банку",
            "Аркуш випробувань",
            "Витяг з наказу",
            "Опис справи",
            "Повідомлення",
        ]
        self.template_listbox.addItems(templates)
        layout.addWidget(self.template_listbox)

        # Horizontal layout for "Choose Custom Templates" and "Unselect All Templates" buttons
        buttons_layout = QHBoxLayout()
        self.choose_custom_templates_button = QPushButton("Choose Custom Templates", self)
        self.choose_custom_templates_button.clicked.connect(self.choose_custom_templates)
        buttons_layout.addWidget(self.choose_custom_templates_button)

        self.unselect_all_button = QPushButton("Unselect All Templates", self)
        self.unselect_all_button.clicked.connect(self.unselect_all_templates)
        buttons_layout.addWidget(self.unselect_all_button)

        layout.addLayout(buttons_layout)

        # Include scores checkbox and button for selecting columns
        self.include_scores_checkbox = QCheckBox("Додати бали", self)
        self.include_scores_checkbox.stateChanged.connect(self.toggle_include_scores)
        layout.addWidget(self.include_scores_checkbox)

        self.select_score_columns_button = QPushButton("Вибір стовпців балів", self)
        self.select_score_columns_button.setEnabled(False)
        self.select_score_columns_button.clicked.connect(self.select_score_columns)
        layout.addWidget(self.select_score_columns_button)

        # Generate documents
        generate_button = QPushButton("Створення документів", self)
        generate_button.clicked.connect(self.generate_documents)
        layout.addWidget(generate_button)

        # Log window
        self.log_window = QTextEdit()
        self.log_window.setReadOnly(True)
        layout.addWidget(self.log_window)

    def log_message(self, message):
        self.log_window.append(message)

    def select_excel_file(self):
        try:
            file, _ = QFileDialog.getOpenFileName(self, "Вибір Excel файлу", str(pathlib.Path.home() / 'Desktop'),
                                                  "Excel Files (*.xlsx)")
            if file:
                self.excel_path = file
                self.log_message(f"Вибраний Excel файл: {self.excel_path}")
                if not self.check_expected_sheets():
                    self.excel_path = None  # Reset excel_path if expected sheets are not found
                    return
                self.show_excel_preview()
                self.update_sheet_dropdown()
        except Exception as e:
            if 'Worksheet named "' in str(e) and ' not found' in str(e):
                pass
            else:
                QMessageBox.critical(self, "Error", f"Під час читання файлу Excel сталася помилка: {e}")

    def check_expected_sheets(self):
        try:
            # Fetch sheet names from the selected Excel file
            self.expected_sheets = pd.ExcelFile(self.excel_path).sheet_names
            return True
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Під час читання файлу Excel сталася помилка: {e}")
            return False

    def update_sheet_dropdown(self):
        if self.excel_path:
            try:
                self.sheet_dropdown.clear()
                self.sheet_dropdown.addItem("Select Sheet")
                self.sheet_dropdown.addItems(self.expected_sheets)
                self.sheet_dropdown.setCurrentIndex(1)  # Set to the first available sheet
                self.update_preview_from_sheet()  # Update preview for the initial sheet
            except Exception as e:
                    QMessageBox.critical(self, "Error", f"Під час читання файлу Excel сталася помилка: {e}")

    def show_excel_preview(self):
        # This method is no longer directly needed for preview since update_preview_from_sheet handles it
        pass

    def update_preview_from_sheet(self):
        selected_sheet = self.sheet_dropdown.currentText()
        if selected_sheet != "Виберіть лист":
            try:
                df = pd.read_excel(self.excel_path, sheet_name=selected_sheet)
                self.preview_table.setRowCount(df.shape[0])
                self.preview_table.setColumnCount(df.shape[1])
                self.preview_table.setHorizontalHeaderLabels(df.columns)

                for i in range(df.shape[0]):
                    for j in range(df.shape[1]):
                        item = QTableWidgetItem(str(df.iloc[i, j]))
                        self.preview_table.setItem(i, j, item)

            except Exception as e:
                if 'Worksheet named' in str(e):
                    self.log_message(str(e))
                else:
                    self.log_message(str(e))

    def map_columns(self):
        selected_sheet = self.sheet_dropdown.currentText()
        if selected_sheet == "Виберіть лист":
            QMessageBox.warning(self, "Аркуш не вибрано", "Спочатку виберіть аркуш.")
            return

        try:
            df = pd.read_excel(self.excel_path, sheet_name=selected_sheet)
            excel_columns = df.columns.tolist()
            required_columns = [
                "Назва групи", "Реєстраційни номер", "Прізвище", "Ім'я", "По батькові", "Адреса", "Контактний номер",
                "Бютжет чи контракт", "Номер групи", "Освітній ступінь", "Спеціальність", "ДПО.Номер", "ДПО.Серія",
                "ДПО.Ким виданий", "Наказ про зарахування", "Серія документа", "Номер документа", "Ким видано",
                "Номер зно", "Рік зно", "Форма навчання", "ДПО", "Тип документа", "Додаток до типу документу",
                "Номер протоколу", "Дата видачі документа", "Дата протоколу", "Дата подачі заяви", "ДПО.Дата видачі",
                "Дата вступу", "Дата наказу"
            ]
            dialog = ColumnMappingDialog(excel_columns, required_columns, self)
            if dialog.exec_():
                self.column_mappings = dialog.mappings
                self.log_message("Column mappings updated.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Під час зіставлення стовпців сталася помилка: {e}")

    def toggle_include_scores(self, state):
        self.include_scores = state == Qt.Checked
        self.select_score_columns_button.setEnabled(self.include_scores)
        if self.include_scores:
            self.select_score_columns()

    def select_score_columns(self):
        if not self.excel_path or self.sheet_dropdown.currentText() == "Виберіть Лист":
            self.log_message("Спочатку виберіть файл Excel і аркуш.")
            return

        selected_sheet = self.sheet_dropdown.currentText()
        try:
            dialog = ScoreColumnSelectorDialog(self.excel_path, selected_sheet, self)
            if dialog.exec_() == QDialog.Accepted:
                self.selected_score_columns = dialog.selected_columns
                self.log_message(f"Selected score columns: {', '.join(self.selected_score_columns)}")
        except Exception as e:
            self.log_message(f"Спочатку виберіть файл Excel і аркуш: {e}")
            return

    def choose_custom_templates(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Виберіть приклад документу",
                                                filter="Word документ (*.docx)")
        if files:
            # Replace the available templates with the selected custom templates
            self.template_listbox.clear()
            for file in files:
                template_name = os.path.splitext(os.path.basename(file))[0]  # Strip extension
                self.template_listbox.addItem(template_name)
            self.log_message("Вибрано користувацький шаблон.")
            return

    def unselect_all_templates(self):
        self.template_listbox.clearSelection()

    def generate_documents(self):
        try:
            self.log_message("Генерація документів...")
            if not all([self.excel_path, self.sheet_dropdown.currentText() != "Виберіть лист"]):
                self.log_message("Відсутня інформація: переконайтеся, що ви вибрали файл Excel і аркуш.")
                return

            selected_items = self.template_listbox.selectedItems()
            if not selected_items and not self.word_templates:
                self.log_message("Шаблон не вибрано: виберіть принаймні один шаблон документа.")
                return

            selected_templates = [item.text() for item in selected_items]
            self.word_templates.extend([str(self.example_dir / f"{template}.docx") for template in selected_templates if
                                        template != "Свій приклад документу"])

            if not self.word_templates:
                self.log_message("Шаблон не вибрано: виберіть дійсний шаблон документа.")
                return

            selected_sheet = self.sheet_dropdown.currentText()
            df = pd.read_excel(self.excel_path, sheet_name=selected_sheet).fillna(' ')

            if self.include_scores and not self.selected_score_columns:
                self.log_message("Не вибрано стовпців оцінок: виберіть стовпці оцінок для включення.")
                return

            # Additional data processing and document generation
            self.process_data(df)
            output_dir = QFileDialog.getExistingDirectory(self, "Виберіть папку для збереження документів")
            if output_dir:
                self.create_documents(df, output_dir)
                self.log_message(f"Документи успішно створено в {output_dir}")
        except Exception as e:
            self.log_message(f"Виникла помилка: {e}")

    def process_data(self, df):
        # Apply column mappings
        for required_column, mapped_column in self.column_mappings.items():
            if mapped_column in df.columns:
                df[required_column] = df[mapped_column]

            # Column mappings as specified
        df["kod1"] = df.get("Назва групи", '')
        df["nomer"] = df.get("Реєстраційни номер", '')
        df["name1"] = df.get("Прізвище", '')
        df["name2"] = df.get("Ім'я", '')
        df["name3"] = df.get("По батькові", '')
        df["adresa"] = df.get("Адреса", '')
        df["mob_number"] = df.get("Контактний номер", '')
        df["form_b"] = df.get("Бютжет чи контракт", '')
        df["gr_num"] = df.get("Номер групи", '')
        df["stupen"] = df.get("Освітній ступінь", '')
        df["spc"] = df.get("Спеціальність", '')
        df["num_pass"] = df.get("ДПО.Номер", '')
        df["seria_pass"] = df.get("ДПО.Серія", '')
        df["vydan"] = df.get("ДПО.Ким виданий", '')
        df["nakaz"] = df.get("Наказ про зарахування", '')
        df["ser_sv"] = df.get("Серія документа", '')
        df["num_sv"] = df.get("Номер документа", '')
        df["kym_vydany"] = df.get("Ким видано", '')
        df["zno_num"] = df.get("Номер зно", '')
        df["zno_rik"] = df.get("Рік зно", '')
        df["forma_nav"] = df.get("Форма навчання", '')
        df["doc_of"] = df.get("ДПО", '')
        df["typ_doc"] = df.get("Тип документа", '')
        df["typ_doc_dod"] = df.get("Додаток до типу документу", '')
        df["prot_num"] = df.get("Номер протоколу", '')

        # Nested function to format date
        def format_date(column_name):
            if column_name in df:
                return pd.to_datetime(df[column_name], errors='coerce').dt.strftime('%d.%m.%Y')
            else:
                return [''] * len(df)

        # Date formatting
        df["data_sv"] = format_date("Дата видачі документа")
        df["data_prot"] = format_date("Дата протоколу")
        df["zayava_vid"] = format_date("Дата подачі заяви")
        df["data"] = format_date("ДПО.Дата видачі")
        df["data_vstup"] = format_date("Дата вступу")
        df["data_nakaz"] = format_date("Дата наказу")

        # Adding today's date in specific formats
        df["d"] = datetime.today().strftime("%d")
        df["m"] = format_datetime(datetime.today(), "MMMM", locale='uk_UA')
        df["Y"] = datetime.today().strftime("%Y")

        # Handling score columns if included
        if self.include_scores:
            try:
                for column in self.selected_score_columns:
                    df[f"{column}_slova"] = df[column].apply(
                        lambda x: num2words(x, lang='uk') if pd.notnull(x) and isinstance(x, (int, float)) else '')
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Сталася помилка під час перетворення балів: {e}")

    def create_documents(self, df, output_dir):
        try:
            for record in df.to_dict(orient="records"):
                for template in self.word_templates:
                    doc = DocxTemplate(template)
                    doc.render(record)
                    template_name = pathlib.Path(template).stem
                    output_path = pathlib.Path(output_dir) / f'{record["Прізвище"]}_{template_name}.docx'
                    doc.save(output_path)
                    self.log_message(f"Документ збережено в: {output_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Під час створення документів сталася помилка: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DocumentGeneratorApp()
    window.show()
    sys.exit(app.exec_())