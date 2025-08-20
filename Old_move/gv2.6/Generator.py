import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton, QListWidget, \
    QComboBox, QTableWidget, QTableWidgetItem, QCheckBox, QMessageBox, QFileDialog, QDialog, QDialogButtonBox, \
    QListWidgetItem
from PyQt5.QtCore import Qt
from PyQt5.Qt import *
from docxtpl import DocxTemplate
from num2words import num2words
from babel.dates import format_datetime
from datetime import datetime
import pandas as pd
import pathlib

class ConsoleWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Console")
        self.setGeometry(800, 100, 500, 300)  # Adjust window position and size as needed

        self.console = QTextEdit(self)
        self.console.setReadOnly(True)

        self.setCentralWidget(self.console)

    def log_to_console(self, message, level='info'):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"{timestamp} [{level.upper()}]: {message}"
        self.console.append(log_entry)
class ScoreColumnSelectorDialog(QDialog):
    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Score Columns")
        self.columns = columns
        self.selected_columns = []

        layout = QVBoxLayout(self)

        self.list_widget = QListWidget(self)
        self.list_widget.setSelectionMode(QListWidget.MultiSelection)

        # Filter columns to include only float or integer columns
        numeric_columns = [col for col in columns if pd.api.types.is_numeric_dtype(col)]

        for column in numeric_columns:
            item = QListWidgetItem(column)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def accept(self):
        self.selected_columns = [item.text() for item in self.list_widget.selectedItems()]
        super().accept()
    def accept(self):
        self.selected_columns = [item.text() for item in self.list_widget.selectedItems()]
        super().accept()

    def accept(self):
        self.selected_columns = [item.text() for item in self.list_widget.selectedItems()]
        super().accept()

class DocumentGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Document Generator")
        self.setGeometry(100, 100, 700, 600)  # Set initial window size

        self.excel_path = None
        self.word_template = None
        self.sheet_name = None
        self.include_scores = False
        self.selected_score_columns = []
        self.output_dir = pathlib.Path(__file__).resolve().parents[1] / "Вихід"
        self.example_dir = pathlib.Path(__file__).resolve().parents[0] / "Приклади"
        self.expected_sheets = []  # Initially empty

        self.setup_gui()

        # Initialize console window
        self.console_window = ConsoleWindow()

    def setup_gui(self):
        central_widget = QWidget()
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Excel file selection
        select_excel_button = QPushButton("Select Excel File", self)
        select_excel_button.clicked.connect(self.select_excel_file)
        layout.addWidget(select_excel_button)

        # Dropdown for sheet selection
        sheet_label = QLabel("Select Sheet", self)
        layout.addWidget(sheet_label)
        self.sheet_dropdown = QComboBox(self)
        self.sheet_dropdown.addItem("Select Sheet")
        self.sheet_dropdown.currentIndexChanged.connect(self.update_preview_from_sheet)
        layout.addWidget(self.sheet_dropdown)

        # Preview label
        preview_label = QLabel("Preview of Selected Excel File:", self)
        layout.addWidget(preview_label)

        # Table widget for preview
        self.preview_table = QTableWidget(self)
        layout.addWidget(self.preview_table)

        # Template selection
        template_label = QLabel("Select Document Template", self)
        layout.addWidget(template_label)

        self.template_listbox = QListWidget(self)
        self.template_listbox.setSelectionMode(QListWidget.SingleSelection)
        self.template_listbox.setMaximumHeight(150)
        templates = [
            "Анкета для банку",
            "Аркуш випробувань",
            "Витяг з наказу",
            "Опис справи",
            "Повідомлення",
            "Свій приклад документу"
        ]
        self.template_listbox.addItems(templates)
        layout.addWidget(self.template_listbox)

        # Include scores checkbox and button for selecting columns
        self.include_scores_checkbox = QCheckBox("Include Scores", self)
        self.include_scores_checkbox.stateChanged.connect(self.toggle_include_scores)
        layout.addWidget(self.include_scores_checkbox)

        self.select_score_columns_button = QPushButton("Select Score Columns", self)
        self.select_score_columns_button.setEnabled(False)
        self.select_score_columns_button.clicked.connect(self.select_score_columns)
        layout.addWidget(self.select_score_columns_button)

        # Generate documents
        generate_button = QPushButton("Generate Documents", self)
        generate_button.clicked.connect(self.generate_documents)
        layout.addWidget(generate_button)

        # Console button
        console_button = QPushButton("Open Console", self)
        console_button.clicked.connect(self.show_console)
        layout.addWidget(console_button)

    def select_excel_file(self):
        try:
            file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", str(pathlib.Path.home() / 'Desktop'),
                                                  "Excel Files (*.xlsx)")
            if file:
                self.excel_path = file
                self.console_window.log_to_console(f"Selected Excel file: {self.excel_path}")
                if not self.check_expected_sheets():
                    self.excel_path = None  # Reset excel_path if expected sheets are not found
                    return
                self.show_excel_preview()
                self.update_sheet_dropdown()
        except Exception as e:
            if 'Worksheet named "' in str(e) and ' not found' in str(e):
                pass
            else:
                self.console_window.log_to_console(f"An error occurred while reading Excel file: {e}", 'error')

    def check_expected_sheets(self):
        try:
            # Fetch sheet names from the selected Excel file
            self.expected_sheets = pd.ExcelFile(self.excel_path).sheet_names
            return True
        except Exception as e:
            self.console_window.log_to_console(f"An error occurred while reading Excel file: {e}", 'error')
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
                if 'Worksheet named "' in str(e) and ' not found' in str(e):
                    pass
                else:
                    self.log_to_console(f"An error occurred while reading Excel file: {e}", level='error')

    def show_excel_preview(self):
        # This method is no longer directly needed for preview since update_preview_from_sheet handles it
        pass

    def update_preview_from_sheet(self):
        selected_sheet = self.sheet_dropdown.currentText()
        if selected_sheet != "Select Sheet":
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
                    print(str(e))
                else:
                    print(str(e))

    def toggle_include_scores(self, state):
        self.include_scores = state == Qt.Checked
        self.select_score_columns_button.setEnabled(self.include_scores)
        if self.include_scores:
            self.select_score_columns()

    def select_score_columns(self):
        if not self.excel_path or self.sheet_dropdown.currentText() == "Select Sheet":
            QMessageBox.warning(self, "No Sheet Selected", "Please select an Excel file and a sheet first.")
            return

        try:
            df = pd.read_excel(self.excel_path, sheet_name=self.sheet_dropdown.currentText())

            # Filter columns to include only float or integer columns
            numeric_columns = df.select_dtypes(include=['number']).columns.tolist()

            dialog = ScoreColumnSelectorDialog(numeric_columns, self)
            if dialog.exec_() == QDialog.Accepted:
                self.selected_score_columns = dialog.selected_columns
                self.log_to_console(f"Selected score columns: {', '.join(self.selected_score_columns)}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while reading Excel file: {e}")

    def generate_documents(self):
        try:
            if not all([self.excel_path, self.sheet_dropdown.currentText() != "Select Sheet"]):
                QMessageBox.warning(self, "Missing Information",
                                    "Please make sure to select an Excel file and a sheet.")
                return

            selected_template = self.template_listbox.currentItem()
            if not selected_template:
                QMessageBox.warning(self, "No Template", "Please select a document template.")
                return

            template_name = selected_template.text()
            if template_name == "Свій приклад документу":
                file, _ = QFileDialog.getOpenFileName(self, "Select Document Template",
                                                      filter="Word Documents (*.docx)")
                if file:
                    self.word_template = file
            else:
                self.word_template = str(self.example_dir / f"{template_name}.docx")

            if not self.word_template:
                QMessageBox.warning(self, "No Template", "Please select a valid document template.")
                return

            what = template_name
            selected_sheet = self.sheet_dropdown.currentText()
            df = pd.read_excel(self.excel_path, sheet_name=selected_sheet).fillna(' ')

            if self.include_scores and not self.selected_score_columns:
                QMessageBox.warning(self, "No Score Columns Selected", "Please select score columns to include.")
                return

            # Additional data processing and document generation
            self.process_data(df)
            self.create_documents(df, what)
            self.log_to_console(f"Documents generated successfully in {self.output_dir}", level='info')
        except Exception as e:
            self.log_to_console(f"An error occurred: {e}", level='error')

    def process_data(self, df):
        df["d"] = datetime.today().strftime("%d")
        df["m"] = format_datetime(datetime.today(), "MMMM", locale='uk_UA')
        df["Y"] = datetime.today().strftime("%Y")
        if self.include_scores:
            try:
                for column in self.selected_score_columns:
                    df[f"{column}_slova"] = df[column].apply(
                        lambda x: num2words(x, lang='uk') if pd.notnull(x) and isinstance(x, (int, float)) else '')
            except Exception as e:
                self.log_to_console(f"An error occurred during score conversion: {e}", level='error')

    def create_documents(self, df, what):
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            for record in df.to_dict(orient="records"):
                doc = DocxTemplate(self.word_template)
                doc.render(record)
                output_path = self.output_dir / f"{record['Прізвище']}_{record['Імя']}_{what}.docx"
                doc.save(output_path)
                self.log_to_console(f"Document saved to: {output_path}", level='info')
        except Exception as e:
            self.log_to_console(f"An error occurred while creating documents: {e}", level='error')

    def log_to_console(self, message, level='info'):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"{timestamp} [{level.upper()}]: {message}"
        self.console.append(log_entry)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DocumentGeneratorApp()
    window.show()
    sys.exit(app.exec_())