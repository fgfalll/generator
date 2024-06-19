import os
import pathlib
import pandas as pd
from pathlib import Path
from tkinter import Tk, filedialog, Label, Button, Listbox, OptionMenu, StringVar, messagebox, Text, Scrollbar
from docxtpl import DocxTemplate
from num2words import num2words
from babel.dates import format_datetime
from IPython.display import HTML


class DocumentGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Document Generator")
        self.root.geometry("700x600")  # Set initial window size

        self.excel_path = None
        self.word_template = None
        self.sheet_name = None
        self.include_scores = None
        self.output_dir = Path(__file__).resolve().parents[1] / "Вихід"
        self.example_dir = Path(__file__).resolve().parents[0] / "Приклади"

        self.setup_gui()

    def setup_gui(self):
        # Excel file selection
        self.select_excel_button = Button(self.root, text="Select Excel File", command=self.select_excel_file, height=2, width=20)
        self.select_excel_button.grid(row=0, column=0, padx=10, pady=10)

        # Dropdown for sheet selection
        self.sheet_var = StringVar(self.root)
        self.sheet_var.set("Select Sheet")
        self.sheet_dropdown = OptionMenu(self.root, self.sheet_var, "Select Sheet")
        self.sheet_dropdown.grid(row=0, column=1, padx=10, pady=10)

        # Preview label
        self.preview_label = Label(self.root, text="Preview of Selected Excel File:", font=("Helvetica", 12, "bold"))
        self.preview_label.grid(row=1, column=0, columnspan=2, pady=10)

        # Text widget for preview
        self.preview_text = Text(self.root, height=10, width=70, wrap='word', state='disabled')
        self.preview_text.grid(row=2, column=0, columnspan=2, padx=10, pady=10)
        scrollbar = Scrollbar(self.root, command=self.preview_text.yview)
        scrollbar.grid(row=2, column=2, sticky='ns')
        self.preview_text.config(yscrollcommand=scrollbar.set)

        # Template selection
        self.template_label = Label(self.root, text="Select Document Template", font=("Helvetica", 12, "bold"))
        self.template_label.grid(row=3, column=0, columnspan=2, pady=10)

        self.template_listbox = Listbox(self.root, selectmode="single", height=6, width=40)
        self.template_listbox.grid(row=4, column=0, columnspan=2, padx=10, pady=10)
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

        # Include scores
        self.include_scores_button = Button(self.root, text="Include Scores", command=self.toggle_include_scores, height=2, width=20)
        self.include_scores_button.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

        # Generate documents
        self.generate_button = Button(self.root, text="Generate Documents", command=self.generate_documents, height=2, width=20)
        self.generate_button.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

    def select_excel_file(self):
        file = filedialog.askopenfile(mode='r', initialdir=pathlib.Path.home() / 'Desktop', title="Select Excel File",
                                      filetypes=[('Excel Files', '*.xlsx')])
        if file:
            self.excel_path = os.path.abspath(file.name)
            messagebox.showinfo("File Selected", f"Selected Excel file: {self.excel_path}")
            self.show_excel_preview()
            self.update_sheet_dropdown()

    def update_sheet_dropdown(self):
        if self.excel_path:
            try:
                sheet_names = pd.ExcelFile(self.excel_path).sheet_names
                self.sheet_var.set("Select Sheet")
                menu = self.sheet_dropdown["menu"]
                menu.delete(0, "end")
                for sheet in sheet_names:
                    menu.add_command(label=sheet, command=lambda value=sheet: self.sheet_var.set(value))
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while reading Excel file: {e}")

    def show_excel_preview(self):
        if self.excel_path:
            try:
                df = pd.read_excel(self.excel_path)
                # Convert DataFrame to string representation for table format
                preview_text = df.to_string(index=False)
                self.preview_text.config(state='normal')
                self.preview_text.delete('1.0', 'end')
                self.preview_text.insert('1.0', preview_text)
                self.preview_text.config(state='disabled')
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while reading Excel file: {e}")

    def toggle_include_scores(self):
        self.include_scores = not self.include_scores
        status = "included" if self.include_scores else "not included"
        messagebox.showinfo("Include Scores", f"Scores will be {status}.")

    def generate_documents(self):
        if not all([self.excel_path, self.sheet_var.get() != "Select Sheet"]):
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
        df = pd.read_excel(self.excel_path, sheet_name=self.sheet_var.get()).fillna(value=' ')

        # Additional data processing and document generation
        try:
            self.process_data(df)
            self.create_documents(df, what)
            messagebox.showinfo("Success", f"Documents generated successfully in {self.output_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def process_data(self, df):
        df["d"] = pd.Timestamp('today').strftime("%d")
        df["m"] = format_datetime(pd.Timestamp('today'), "MMMM", locale='uk_UA')
        df["Y"] = pd.Timestamp('today').strftime("%Y")
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
