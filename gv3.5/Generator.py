import shutil
import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, \
    QListWidget, \
    QComboBox, QTableWidget, QTableWidgetItem, QCheckBox, QMessageBox, QFileDialog, QDialog, QDialogButtonBox, \
    QListWidgetItem, QTextEdit, QScrollArea, QFormLayout, QInputDialog
from PyQt5.QtCore import Qt
from docxtpl import DocxTemplate
from num2words import num2words
from babel.dates import format_datetime
from datetime import datetime
import pandas as pd
import pathlib
from transliterate import translit

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

    def get_transliterated_columns(self):
        try:
            transliterated_columns = {}
            for column in self.selected_columns:
                transliterated_name = translit(column, 'uk', reversed=True).lower()
                transliterated_columns[column] = transliterated_name
            return transliterated_columns
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Під час транслітерації сталася помилка: {e}")
            return {}


class ColumnMappingDialog(QDialog):
    def __init__(self, df, required_columns, column_mappings=None, parent=None):
        super(ColumnMappingDialog, self).__init__(parent)
        self.setWindowTitle("Map Columns")
        self.setGeometry(100, 100, 1500, 600)
        self.df = df.copy()  # Make a copy of the DataFrame to avoid modifying the original
        self.required_columns = required_columns.copy()  # Copy the required columns
        self.column_mappings = column_mappings if column_mappings else {}

        # Initialize combo_boxes attribute to store references to combo boxes
        self.combo_boxes = {}

        # Main layout
        main_layout = QVBoxLayout()

        # Scroll Area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scrollContent = QWidget(scroll)
        scrollLayout = QVBoxLayout(scrollContent)

        # Create form layout for mappings
        form_layout = QFormLayout()

        for required_column in self.required_columns:
            label = QLabel(required_column)
            combo_box = QComboBox()
            combo_box.addItem("Skip")  # Add Skip option
            combo_box.addItems(self.df.columns)  # Add columns from the DataFrame
            combo_box.setCurrentText(self.column_mappings.get(required_column, "Skip"))  # Set current mapping if exists
            form_layout.addRow(label, combo_box)
            self.combo_boxes[required_column] = combo_box

        scrollLayout.addLayout(form_layout)
        scroll.setWidget(scrollContent)
        main_layout.addWidget(scroll)

        # Excel preview table
        self.preview_table = QTableWidget(self)
        self.update_preview_table()
        main_layout.addWidget(self.preview_table)

        # OK and Cancel buttons
        buttons_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        cancel_button = QPushButton("Cancel")
        buttons_layout.addWidget(ok_button)
        buttons_layout.addWidget(cancel_button)
        main_layout.addLayout(buttons_layout)

        ok_button.clicked.connect(self.accept)
        cancel_button.clicked.connect(self.reject)

        self.setLayout(main_layout)

        # Automap required columns if present in df
        self.automap_required_columns()

    def get_mapped_columns(self):
        return self.column_mappings

    def automap_required_columns(self):
        for column in self.required_columns:
            if column in self.df.columns:
                self.column_mappings[column] = column
                if column in self.combo_boxes:
                    self.combo_boxes[column].setCurrentText(column)

    def accept(self):
        try:
            for required_column, combo_box in self.combo_boxes.items():
                selected_column = combo_box.currentText()
                if selected_column != "Skip" and selected_column in self.df.columns:
                    # Ensure we don't rename multiple columns to the same new name
                    if selected_column in self.column_mappings.values() and self.column_mappings.get(required_column) != selected_column:
                        QMessageBox.warning(self, "Duplicate Mapping", f"The column {selected_column} is already mapped.")
                        return
                    self.column_mappings[required_column] = selected_column
            super(ColumnMappingDialog, self).accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while saving mappings: {e}")

    def restore_column_mappings(self):
        for required_column, selected_column in self.column_mappings.items():
            if required_column in self.combo_boxes and selected_column in self.df.columns:
                combo_box = self.combo_boxes[required_column]
                combo_box.setCurrentText(selected_column)

    def update_preview_table(self):
        self.preview_table.clear()
        self.preview_table.setRowCount(self.df.shape[0])
        self.preview_table.setColumnCount(self.df.shape[1])

        # Update DataFrame with temporary column names for preview
        preview_df = self.df.copy()
        for required_column, combo_box in self.combo_boxes.items():
            selected_column = combo_box.currentText()
            if selected_column != "Skip":
                preview_df.rename(columns={required_column: selected_column}, inplace=True)

        self.preview_table.setHorizontalHeaderLabels(preview_df.columns)

        for i in range(preview_df.shape[0]):
            for j in range(preview_df.shape[1]):
                item = QTableWidgetItem(str(preview_df.iloc[i, j]))
                self.preview_table.setItem(i, j, item)

class DocumentGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Generator V3")
        self.setGeometry(100, 100, 700, 600)

        self.excel_path = None
        self.word_templates = []
        self.sheet_name = None
        self.include_scores = False
        self.selected_score_columns = []
        self.column_mappings = {}
        self.output_dir = pathlib.Path(__file__).resolve().parents[1] / "Вихід"
        self.example_dir = pathlib.Path(__file__).resolve().parents[0] / "Приклади"
        self.expected_sheets = []

        self.setup_gui()

    def setup_gui(self):
        central_widget = QWidget()
        central_widget.setAcceptDrops(True)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        select_excel_button = QPushButton("Виберіть Excel файл", self)
        select_excel_button.clicked.connect(self.select_excel_file)
        layout.addWidget(select_excel_button)

        sheet_label = QLabel("Виберіть лист", self)
        layout.addWidget(sheet_label)
        self.sheet_dropdown = QComboBox(self)
        self.sheet_dropdown.addItem("Виберіть лист")
        self.sheet_dropdown.currentIndexChanged.connect(self.update_preview_from_sheet)
        layout.addWidget(self.sheet_dropdown)

        preview_label = QLabel("Попередній перегляд вибраного файлу Excel:", self)
        layout.addWidget(preview_label)

        self.preview_table = QTableWidget(self)
        layout.addWidget(self.preview_table)

        map_columns_button = QPushButton("Співставлення колонок", self)
        map_columns_button.clicked.connect(self.map_columns)
        layout.addWidget(map_columns_button)

        template_label = QLabel("Виберіть Шаблони документів", self)
        layout.addWidget(template_label)

        self.template_listbox = QListWidget(self)
        self.template_listbox.setSelectionMode(QListWidget.MultiSelection)
        self.template_listbox.setMaximumHeight(150)  # Set maximum height

        # Dynamically populate template_listbox from example_dir
        self.populate_template_list()

        layout.addWidget(self.template_listbox)

        buttons_layout = QHBoxLayout()
        self.choose_custom_templates_button = QPushButton("Choose Custom Templates", self)
        self.choose_custom_templates_button.clicked.connect(self.choose_custom_templates)
        buttons_layout.addWidget(self.choose_custom_templates_button)

        self.unselect_all_button = QPushButton("Unselect All Templates", self)
        self.unselect_all_button.clicked.connect(self.unselect_all_templates)
        buttons_layout.addWidget(self.unselect_all_button)

        layout.addLayout(buttons_layout)

        self.include_scores_checkbox = QCheckBox("Додати бали", self)
        self.include_scores_checkbox.stateChanged.connect(self.toggle_include_scores)
        layout.addWidget(self.include_scores_checkbox)

        self.select_score_columns_button = QPushButton("Вибір стовпців балів", self)
        self.select_score_columns_button.setEnabled(False)
        self.select_score_columns_button.clicked.connect(self.select_score_columns)
        layout.addWidget(self.select_score_columns_button)

        generate_button = QPushButton("Створення документів", self)
        generate_button.clicked.connect(self.generate_documents)
        layout.addWidget(generate_button)

        self.log_window = QTextEdit()
        self.log_window.setReadOnly(True)
        layout.addWidget(self.log_window)

    def populate_template_list(self):
        # Clear existing items
        self.template_listbox.clear()

        # Fetch templates from example_dir and add them to template_listbox
        try:
            example_dir = pathlib.Path(__file__).resolve().parents[0] / "Приклади"
            standard_templates = [file.stem for file in example_dir.glob("*.docx") if file.is_file()]

            # Add standard templates to the list
            self.template_listbox.addItems(standard_templates)

            # Add custom templates to the list
            for template in self.word_templates:
                template_name = pathlib.Path(template).stem
                if template_name not in standard_templates:
                    self.template_listbox.addItem(template_name)

        except Exception as e:
            self.log_message(f"Error fetching templates: {e}")

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
                self.sheet_dropdown.addItems(self.expected_sheets)
                self.sheet_dropdown.setCurrentIndex(0)  # Set to the first available sheet
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
            required_columns = [
                "Назва групи", "Реєстраційни номер", "Прізвище", "Ім'я", "По батькові", "Адреса", "Контактний номер",
                "Бютжет чи контракт", "Номер групи", "Освітній ступінь", "Спеціальність", "ДПО.Номер", "ДПО.Серія",
                "ДПО.Ким виданий", "Наказ про зарахування", "Серія документа", "Номер документа", "Ким видано",
                "Номер зно", "Рік зно", "Форма навчання", "ДПО", "Тип документа", "Додаток до типу документу",
                "Номер протоколу", "Дата видачі документа", "Дата протоколу", "Дата подачі заяви", "ДПО.Дата видачі",
                "Дата вступу", "Дата наказу"
            ]

            dialog = ColumnMappingDialog(df, required_columns, self.column_mappings, self)
            if dialog.exec_():
                self.column_mappings = dialog.get_mapped_columns()
                self.log_message("Column mappings updated.")

                self.update_preview_from_sheet()  # Refresh the preview to show new column names

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Під час зіставлення стовпців сталася помилка: {e}")
            self.log_message(f"Error in map_columns: {e}")

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
                transliterated_columns = dialog.get_transliterated_columns()

                formatted_columns = []
                for col in self.selected_score_columns:
                    # Strip trailing dots, convert to lower case, replace spaces with underscores, strip again for safety
                    formatted_col = col.replace('.',' ').lower().replace(' ', '_').rstrip('_')
                    formatted_columns.append(formatted_col)

                self.selected_score_columns = formatted_columns

                selected_columns_message = f"Selected score columns: {', '.join(self.selected_score_columns)}"
                self.log_message(selected_columns_message)

                suffix_columns_message = f"Score columns with '_slova' suffix: {', '.join([col + '_slova' for col in self.selected_score_columns])}"
                self.log_message(suffix_columns_message)

        except Exception as e:
            self.log_message(f"Спочатку виберіть файл Excel і аркуш: {e}")
            self.log_message(f"Error in select_score_columns: {e}")
            return

    def choose_custom_templates(self):
        try:
            # Create a QFileDialog instance
            file_dialog = QFileDialog()

            # Set properties of the file dialog
            file_dialog.setFileMode(QFileDialog.ExistingFiles)
            file_dialog.setNameFilter("Word Files (*.docx)")
            file_dialog.setWindowTitle("Choose Custom Templates")

            # Use native file dialog if available for better integration
            file_dialog.setOption(QFileDialog.DontUseNativeDialog, False)

            # Execute the dialog and process the result
            if file_dialog.exec_():
                custom_templates = file_dialog.selectedFiles()
                if not custom_templates:
                    return  # Exit if no files are selected

                # Log selected templates
                self.log_message(f"Selected custom templates: {', '.join(custom_templates)}")

                # Define example_dir where standard templates are stored
                example_dir = pathlib.Path(__file__).resolve().parents[0] / "Приклади"

                # Create example_dir if it does not exist
                example_dir.mkdir(parents=True, exist_ok=True)

                # Copy custom templates to example_dir
                for template in custom_templates:
                    template_path = pathlib.Path(template)
                    if template_path.exists():
                        template_name = template_path.name
                        destination = example_dir / template_name
                        shutil.copyfile(template, destination)
                        self.log_message(f"Copied template {template_name} to standard templates folder.")
                    else:
                        self.log_message(f"File {template} does not exist or is inaccessible.")

                # Clear existing word_templates before adding custom templates
                self.word_templates.clear()

                # Add copied templates to word_templates
                self.word_templates.extend(
                    str(example_dir / pathlib.Path(template).name) for template in custom_templates)

                self.log_message("Custom templates imported successfully.")

                # Update the template listbox view
                self.populate_template_list()

            else:
                # User canceled the file dialog
                self.log_message("Custom template selection canceled by user.")

        except Exception as e:
            # Handle exceptions and log error messages
            self.log_message(f"Error selecting custom templates: {e}")
            QMessageBox.critical(self, "Error", f"An error occurred while choosing custom templates: {e}")

    def unselect_all_templates(self):
        self.template_listbox.clearSelection()
        self.word_templates.clear()  # Clear the list of selected word templates
        self.word_templates = []
        self.log_message("All templates unselected.")

        # Repopulate template_listbox with standard templates
        self.populate_template_list()

    def generate_documents(self):
        try:
            if not all([self.excel_path, self.sheet_dropdown.currentText() != "Виберіть лист"]):
                self.log_message("Відсутня інформація: переконайтеся, що ви вибрали файл Excel і аркуш.")
                return

            selected_items = self.template_listbox.selectedItems()
            if not selected_items:
                self.log_message("Шаблон не вибрано: виберіть принаймні один шаблон документа.")
                self.prompt_choose_templates()  # Prompt user to choose templates
                return

            selected_templates = [item.text() for item in selected_items]
            self.word_templates.extend([str(self.example_dir / f"{template}.docx") for template in selected_templates if
                                        template != "Свій приклад документу"])

            if not self.word_templates:
                self.log_message("Шаблон не вибрано: виберіть дійсний шаблон документа.")
                self.prompt_choose_templates()  # Prompt user to choose templates
                return

            selected_sheet = self.sheet_dropdown.currentText()
            df = pd.read_excel(self.excel_path, sheet_name=selected_sheet).fillna(' ')

            if self.include_scores and not self.selected_score_columns:
                self.log_message("Не вибрано стовпців оцінок: виберіть стовпці оцінок для включення.")
                return

            # Additional data processing and document generation
            self.process_data(df)
            self.log_message("Генерація документів...")
            output_dir = QFileDialog.getExistingDirectory(self, "Виберіть папку для збереження документів")
            if output_dir:
                self.create_documents(df, output_dir)
                self.log_message(f"Документи успішно створено в {output_dir}")
        except Exception as e:
            self.log_message(f"Помилка під час створення документів: {e}")
            QMessageBox.critical(self, "Помилка", f"Під час створення документів сталася помилка: {e}")

    def check_standard_templates(self):
        try:
            # Fetch standard templates from the example directory
            example_dir = pathlib.Path(__file__).resolve().parents[0] / "Приклади"
            standard_templates_path = example_dir.glob("*.docx")
            available_templates = [str(template_path) for template_path in standard_templates_path if
                                   template_path.is_file()]

            if available_templates:
                # Clear existing selections
                self.template_listbox.clear()
                self.word_templates.clear()

                # Add standard templates to template_listbox and word_templates
                for template_path in available_templates:
                    template_name = os.path.splitext(os.path.basename(template_path))[0]
                    self.template_listbox.addItem(template_name)
                    self.word_templates.append(template_path)

                self.log_message("Використовуються доступні стандартні шаблони.")
                return True
            else:
                self.log_message("Стандартні шаблони не знайдені.")
                return False
        except Exception as e:
            self.log_message(f"Під час перевірки стандартних шаблонів сталася помилка: {e}")
            return False

    def prompt_choose_templates(self):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle("Шаблон не вибрано")
        msg_box.setText("Шаблон не знайдено. Будь ласка, виберіть користувацькі шаблони.")
        msg_box.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        result = msg_box.exec_()
        if result == QMessageBox.Ok:
            self.choose_custom_templates()

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

        # Process each required column
        try:
            # Define required columns based on include_scores
            required_columns = {
                "kod1": "Назва групи",
                "nomer": "Реєстраційни номер",
                "name1": "Прізвище",
                "name2": "Ім'я",
                "name3": "По батькові",
                "adresa": "Адреса",
                "mob_number": "Контактний номер",
                "form_b": "Бютжет чи контракт",
                "gr_num": "Номер групи",
                "stupen": "Освітній ступінь",
                "spc": "Спеціальність",
                "num_pass": "ДПО.Номер",
                "seria_pass": "ДПО.Серія",
                "vydan": "ДПО.Ким виданий",
                "nakaz": "Наказ про зарахування",
                "ser_sv": "Серія документа",
                "num_sv": "Номер документа",
                "kym_vydany": "Ким видано",
                "zno_num": "Номер зно",
                "zno_rik": "Рік зно",
                "forma_nav": "Форма навчання",
                "doc_of": "ДПО",
                "typ_doc": "Тип документа",
                "typ_doc_dod": "Додаток до типу документу",
                "prot_num": "Номер протоколу",
            }

            if self.include_scores:
                required_columns.update({
                    f"{column.lower().rstrip('.').replace(' ', '_')}": column for column in self.selected_score_columns
                })

            # Process each required column
            for new_column, old_column in required_columns.items():
                if old_column in df.columns:
                    df[new_column] = df[old_column]
                else:
                    # Ask user whether to map or skip the missing column
                    reply = QMessageBox.question(self, 'Missing Column',
                                                 f"Column '{old_column}' not found in DataFrame. Do you want to map it to another column or skip it?",
                                                 QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
                    if reply == QMessageBox.Yes:
                        # Provide options to the user to select from available columns
                        available_columns = list(df.columns)
                        mapped_column, ok = QInputDialog.getItem(self, "Map Column", f"Map '{old_column}' to:",
                                                                 available_columns, 0, False)
                        if ok and mapped_column:
                            df[new_column] = df[mapped_column]
                        else:
                            self.log_message(f"Column '{old_column}' mapping skipped by the user.")
                    elif reply == QMessageBox.No:
                        self.log_message(f"Column '{old_column}' skipped.")
                    else:
                        self.log_message(f"Process interrupted by the user.")
                        return

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
            df["m"] = datetime.today().strftime("%B")
            df["Y"] = datetime.today().strftime("%Y")

            # Handling score columns if included
            if self.include_scores:
                for column in self.selected_score_columns:
                    if column in df.columns:
                        column_transliterated = column.lower().replace(' ', '_')
                        df[f"{column_transliterated}_slova"] = df[column].apply(
                            lambda x: num2words(x, lang='uk') if pd.notnull(x) and isinstance(x, (int, float)) else ''
                        )
                    else:
                        QMessageBox.warning(self, "Warning", f"Column '{column}' not found in DataFrame.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while processing data: {e}")

    def create_documents(self, df, output_dir):
        for idx, row in df.iterrows():
            context = row.to_dict()
            context = {key.lower().replace(' ', '_').replace('.', '').replace(',', ''): value for key, value in
                       context.items()}

            # Extract name parts for the file name
            surname = context.get("прізвище", "")
            name = context.get("ім'я", "")
            patronymic = context.get("по_батькові", "")

            # Construct the file name
            file_name = f"{surname} {name} {patronymic}.docx"

            for template_path in self.word_templates:
                try:
                    template_name = pathlib.Path(template_path).stem
                    doc = DocxTemplate(template_path)
                    doc.render(context)
                    doc.save(f"{output_dir}/{template_name}_{file_name}")
                except Exception as e:
                    self.log_message(f"Error rendering document for row {idx} with template {template_path}: {e}")
                    self.log_message(f"Error in create_documents: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DocumentGeneratorApp()
    window.show()
    sys.exit(app.exec_())