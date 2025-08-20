import datetime
import os
import pathlib
from pathlib import Path
import pandas as pd
from babel.dates import format_datetime
from docxtpl import DocxTemplate
from num2words import num2words
from tkinter import Tk, filedialog, Label, Button, Listbox, SINGLE, messagebox


class DocumentGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Generator")

        self.excel_path = None
        self.word_template = None
        self.sheet_name = None
        self.include_scores = None
        self.output_dir = Path(__file__).resolve().parents[1] / "Вихід"
        self.example_dir = Path(__file__).resolve().parents[0] / "Приклади"

        self.setup_gui()

    def setup_gui(self):
        # Excel file selection
        self.select_excel_button = Button(self.root, text="Select Excel File", command=self.select_excel_file)
        self.select_excel_button.pack(pady=10)

        # Template selection
        self.template_label = Label(self.root, text="Select Document Template")
        self.template_label.pack(pady=10)

        self.template_listbox = Listbox(self.root, selectmode=SINGLE)
        self.template_listbox.pack(pady=10)
        templates = [
            "Анкета для банку",
            "Аркуш випробувань",
            "Витяг з наказу",
            "Опис справи",
            "Повідомлення",
            "Свій приклад документу"
        ]
        for template in templates:
            self.template_listbox.insert('end', template)

        # Sheet selection
        self.select_sheet_button = Button(self.root, text="Select Sheet", command=self.select_sheet)
        self.select_sheet_button.pack(pady=10)

        # Include scores
        self.include_scores_button = Button(self.root, text="Include Scores", command=self.toggle_include_scores)
        self.include_scores_button.pack(pady=10)

        # Generate documents
        self.generate_button = Button(self.root, text="Generate Documents", command=self.generate_documents)
        self.generate_button.pack(pady=10)

    def select_excel_file(self):
        file = filedialog.askopenfile(mode='r', initialdir=pathlib.Path.home() / 'Desktop', title="Select Excel File",
                                      filetypes=[('Excel Files', '*.xlsx')])
        if file:
            self.excel_path = os.path.abspath(file.name)
            messagebox.showinfo("File Selected", f"Selected Excel file: {self.excel_path}")

    def select_sheet(self):
        if not self.excel_path:
            messagebox.showwarning("No Excel File", "Please select an Excel file first.")
            return
        sheet_names = pd.ExcelFile(self.excel_path).sheet_names
        self.sheet_listbox = Listbox(self.root, selectmode=SINGLE)
        for sheet in sheet_names:
            self.sheet_listbox.insert('end', sheet)
        self.sheet_listbox.pack(pady=10)
        self.sheet_listbox.bind('<<ListboxSelect>>', self.on_sheet_select)

    def on_sheet_select(self, event):
        selection = event.widget.curselection()
        if selection:
            self.sheet_name = event.widget.get(selection[0])
            messagebox.showinfo("Sheet Selected", f"Selected Sheet: {self.sheet_name}")

    def toggle_include_scores(self):
        self.include_scores = not self.include_scores
        status = "included" if self.include_scores else "not included"
        messagebox.showinfo("Include Scores", f"Scores will be {status}.")

    def generate_documents(self):
        if not all([self.excel_path, self.sheet_name]):
            messagebox.showwarning("Missing Information", "Please make sure to select an Excel file and a sheet.")
            return

        selected_template = self.template_listbox.curselection()
        if not selected_template:
            messagebox.showwarning("No Template", "Please select a document template.")
            return

        template_name = self.template_listbox.get(selected_template[0])
        if template_name == "Свій приклад документу":
            file = filedialog.askopenfile(mode='r', title="Select Document Template",
                                          filetypes=[('Word Documents', '*.docx')])
            if file:
                self.word_template = os.path.abspath(file.name)
        else:
            self.word_template = self.example_dir / f"{template_name}.docx"

        if not self.word_template:
            messagebox.showwarning("No Template", "Please select a valid document template.")
            return

        what = template_name
        df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name).fillna(value=' ')

        # Additional data processing and document generation
        try:
            self.process_data(df)
            self.create_documents(df, what)
            messagebox.showinfo("Success", f"Documents generated successfully in {self.output_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def process_data(self, df):
        df["d"] = datetime.datetime.today().strftime("%d")
        df["m"] = format_datetime(datetime.datetime.today(), "MMMM", locale='uk_UA')
        df["Y"] = datetime.datetime.today().strftime("%Y")
        if self.include_scores:
            df["baly_ukr_slova"] = df["ЗНО.Українська мова та література"].apply(num2words, lang='uk')
            df["baly_mat_slova"] = df["ЗНО.Математика"].apply(num2words, lang='uk')
            df["reg_kof_slov"] = df["Регіональний коефіцієнт"].apply(num2words, lang='uk')
            df["gal_kof_slov"] = df["Галузевий коефіцієнт"].apply(num2words, lang='uk')
            df["sr_bal_slova"] = df["Конкурсний бал"].apply(num2words, lang='uk')

    def create_documents(self, df, what):
        self.output_dir.mkdir(exist_ok=True)
        for record in df.to_dict(orient="records"):
            doc = DocxTemplate(self.word_template)
            doc.render(record)
            output_path = self.output_dir / f"{what}.docx"
            doc.save(output_path)
            print(f"Document saved to: {output_path}")


if __name__ == "__main__":
    root = Tk()
    app = DocumentGeneratorApp(root)
    root.mainloop()
