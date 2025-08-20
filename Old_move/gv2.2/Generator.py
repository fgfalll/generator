import os
import pathlib
import pandas as pd
from pathlib import Path
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QListWidget, QComboBox, QTextEdit, QMessageBox, QFileDialog, QTableWidget, QTableWidgetItem
from docxtpl import DocxTemplate
from num2words import num2words
from babel.dates import format_datetime
from datetime import datetime


class DocumentGeneratorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Document Generator")
        self.setGeometry(100, 100, 700, 600)  # Set initial window size

        self.excel_path = None
        self.word_template = None
        self.sheet_name = None
        self.include_scores = False
        self.output_dir = Path(__file__).resolve().parents[1] / "Вихід"
        self.example_dir = Path(__file__).resolve().parents[0] / "Приклади"

        self.setup_gui()

    def setup_gui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

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
        preview_label.setFont(self.font())  # Set font directly from QMainWindow
        layout.addWidget(preview_label)

        # Table widget for preview
        self.preview_table = QTableWidget(self)
        layout.addWidget(self.preview_table)

        # Template selection
        template_label = QLabel("Select Document Template", self)
        template_label.setFont(self.font())  # Set font directly from QMainWindow
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

        # Include scores
        include_scores_button = QPushButton("Include Scores", self)
        include_scores_button.clicked.connect(self.toggle_include_scores)
        layout.addWidget(include_scores_button)

        # Generate documents
        generate_button = QPushButton("Generate Documents", self)
        generate_button.clicked.connect(self.generate_documents)
        layout.addWidget(generate_button)

    def select_excel_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", str(pathlib.Path.home() / 'Desktop'), "Excel Files (*.xlsx)")
        if file:
            self.excel_path = file
            QMessageBox.information(self, "File Selected", f"Selected Excel file: {self.excel_path}")
            self.show_excel_preview()
            self.update_sheet_dropdown()

    def update_sheet_dropdown(self):
        if self.excel_path:
            try:
                sheet_names = pd.ExcelFile(self.excel_path).sheet_names
                self.sheet_dropdown.clear()
                self.sheet_dropdown.addItem("Select Sheet")
                self.sheet_dropdown.addItems(sheet_names)
                self.sheet_dropdown.setCurrentIndex(1)  # Set to the first available sheet
                self.update_preview_from_sheet()  # Update preview for the initial sheet
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred while reading Excel file: {e}")

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
                QMessageBox.critical(self, "Error", f"An error occurred while reading Excel file: {e}")

    def show_excel_preview(self):
        # This method is no longer directly needed for preview since update_preview_from_sheet handles it
        pass

    def toggle_include_scores(self):
        self.include_scores = not self.include_scores
        status = "included" if self.include_scores else "not included"
        QMessageBox.information(self, "Include Scores", f"Scores will be {status}.")

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

            # Check if scores columns are required and available
            if self.include_scores:
                required_columns = ["ЗНО.Українська мова та література", "ЗНО.Математика", "Регіональний коефіцієнт",
                                    "Галузевий коефіцієнт", "Конкурсний бал"]
                available_columns = df.columns.tolist()

                if not all(col in available_columns for col in required_columns):
                    missing_columns = [col for col in required_columns if col not in available_columns]
                    QMessageBox.warning(self, "Scores Not Available",
                                        f"Selected sheet does not contain all required score columns: {', '.join(missing_columns)}")
                    return

            # Additional data processing and document generation
            self.process_data(df)
            self.create_documents(df, what)
            QMessageBox.information(self, "Success", f"Documents generated successfully in {self.output_dir}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")

    def process_data(self, df):
        df["d"] = datetime.today().strftime("%d")
        df["m"] = format_datetime(datetime.today(), "MMMM", locale='uk_UA')
        df["Y"] = datetime.today().strftime("%Y")

        if self.include_scores:
            try:
                # Convert numeric scores to words, handling exceptions for invalid data
                df["baly_ukr_slova"] = df["ЗНО.Українська мова та література"].apply(
                    lambda x: num2words(x, lang='uk') if pd.notnull(x) and isinstance(x, (int, float)) else '')
                df["baly_mat_slova"] = df["ЗНО.Математика"].apply(
                    lambda x: num2words(x, lang='uk') if pd.notnull(x) and isinstance(x, (int, float)) else '')
                df["reg_kof_slov"] = df["Регіональний коефіцієнт"].apply(
                    lambda x: num2words(x, lang='uk') if pd.notnull(x) and isinstance(x, (int, float)) else '')
                df["gal_kof_slov"] = df["Галузевий коефіцієнт"].apply(
                    lambda x: num2words(x, lang='uk') if pd.notnull(x) and isinstance(x, (int, float)) else '')
                df["sr_bal_slova"] = df["Конкурсний бал"].apply(
                    lambda x: num2words(x, lang='uk') if pd.notnull(x) and isinstance(x, (int, float)) else '')
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred during score conversion: {e}")

    def create_documents(self, df, what):
        self.output_dir.mkdir(parents=True, exist_ok=True)
        for index, record in df.iterrows():
            doc = DocxTemplate(self.word_template)
            doc.render(record.to_dict())
            output_path = self.output_dir / f"{what}.docx"
            doc.save(str(output_path))
            print(f"Document saved to: {output_path}")


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = DocumentGeneratorApp()
    window.show()
    sys.exit(app.exec_())
