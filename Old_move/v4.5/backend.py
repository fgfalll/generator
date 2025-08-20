import shutil
import pathlib
import pandas as pd
from PyQt6.QtCore import QObject
from PyQt6.QtWidgets import QFileDialog, QMessageBox, QInputDialog
from docxtpl import DocxTemplate
from num2words import num2words
from babel.dates import format_datetime
from datetime import datetime
from transliterate import translit
import io
import mammoth

from ui import MainWindow
from column_mapper import ColumnMappingDialog
from score_mapper_dialog import ScoreMappingDialog
from settings_dialog import SettingsDialog
import theme_manager

class Backend(QObject):
    def __init__(self, ui, app):
        super().__init__()
        self.ui = ui
        self.app = app
        
        self.excel_path = None
        self.df = None
        self.column_mappings = {}
        self.score_mappings = []
        
        self.categorized_templates = {}
        self.template_paths = {}

        self.base_dir = pathlib.Path(__file__).resolve().parent
        self.example_dir = self.base_dir / "Приклади"
        self.example_dir.mkdir(parents=True, exist_ok=True)

        self.CATEGORY_KEYWORDS = {
            'Аркуш випробувань': ['випробувань', 'тестовий', 'test'],
            'Витяг з наказу': ['витяг', 'наказ', 'order'],
            'Опис справи': ['опис', 'справ', 'case'],
            'Повідомлення': ['повідомлення', 'лист', 'notification']
        }
        self.DEFAULT_CATEGORY = 'Різне'
        
        self.required_columns = [
            "Назва групи", "Реєстраційний номер", "Прізвище", "Ім'я", "По батькові", "Адреса", "Контактний номер",
            "Бютжет чи контракт", "Номер групи", "ОКР", "Спеціальність", "ДПО.Номер", "ДПО.Серія",
            "ДПО.Ким виданий", "Наказ про зарахування", "Серія документа", "Номер документа", "Ким видано",
            "Номер зно", "Рік зно", "Форма навчання", "ДПО", "Тип документа", "Додаток до типу документу",
            "Номер протоколу", "Дата видачі документа", "Дата протоколу", "Дата подачі заяви", "ДПО.Дата видачі",
            "Дата вступу", "Дата наказу"
        ]

        self.TEMPLATE_KEYWORDS = {
            "Назва групи": "kod1", "Реєстраційний номер": "nomer", "Прізвище": "name1", "Ім'я": "name2",
            "По батькові": "name3", "Адреса": "adresa", "Контактний номер": "mob_number",
            "Бютжет чи контракт": "form_b", "Номер групи": "gr_num", "ОКР": "stupen", "Спеціальність": "spc",
            "ДПО.Номер": "num_pass", "ДПО.Серія": "seria_pass", "ДПО.Ким виданий": "vydan",
            "Наказ про зарахування": "nakaz", "Серія документа": "ser_sv", "Номер документа": "num_sv",
            "Ким видано": "kym_vydany", "Номер зно": "zno_num", "Рік зно": "zno_rik",
            "Форма навчання": "forma_nav", "ДПО": "doc_of", "Тип документа": "typ_doc",
            "Додаток до типу документу": "typ_doc_dod", "Номер протоколу": "prot_num",
            "Дата видачі документа": "data_sv", "Дата протоколу": "data_prot", "Дата подачі заяви": "zayava_vid",
            "ДПО.Дата видачі": "data", "Дата вступу": "data_vstup", "Дата наказу": "data_nakaz",
        }

    def get_template_keywords(self):
        """Returns the dictionary mapping human-readable names to template keys."""
        return self.TEMPLATE_KEYWORDS

    def _log(self, message, level='INFO'):
        self.ui.log_message(message, level)

    def connect_signals(self):
        self.ui.open_settings_signal.connect(self.open_settings)
        self.ui.select_excel_file_signal.connect(self.select_excel_file)
        self.ui.sheet_changed_signal.connect(self.load_sheet_data)
        self.ui.map_columns_signal.connect(self.map_columns)
        self.ui.refresh_templates_signal.connect(self.scan_and_categorize_templates)
        self.ui.template_category_changed_signal.connect(self.on_template_category_changed)
        self.ui.choose_custom_templates_signal.connect(self.choose_custom_templates)
        self.ui.unselect_all_templates_signal.connect(self.ui.template_listbox.clearSelection)
        self.ui.configure_scores_signal.connect(self.open_score_mapping_dialog)
        self.ui.generate_documents_signal.connect(self.generate_documents)
        
        self.scan_and_categorize_templates()

    def open_settings(self):
        current_settings = theme_manager.load_settings()
        dialog = SettingsDialog(current_settings["theme_name"], current_settings["custom_colors"], self.ui)
        if dialog.exec():
            new_settings = dialog.get_settings()
            theme_manager.save_settings(new_settings["theme_name"], new_settings["custom_colors"])
            new_qss = theme_manager.generate_qss(new_settings["theme_name"], new_settings["custom_colors"])
            self.app.setStyleSheet(new_qss)
            self._log("Налаштування теми оновлено.", 'INFO')

    def scan_and_categorize_templates(self):
        self._log("Сканування та категоризація шаблонів...", 'PROCESS')
        self.template_paths.clear()
        self.categorized_templates = {cat: [] for cat in self.CATEGORY_KEYWORDS}
        self.categorized_templates[self.DEFAULT_CATEGORY] = []
        all_docx_files = list(self.example_dir.rglob("*.docx"))

        for path in all_docx_files:
            if not path.is_file(): continue
            stem = path.stem
            self.template_paths[stem] = path
            assigned = False
            for category, keywords in self.CATEGORY_KEYWORDS.items():
                if any(keyword in stem.lower() for keyword in keywords):
                    self.categorized_templates[category].append(stem)
                    assigned = True; break
            if not assigned:
                self.categorized_templates[self.DEFAULT_CATEGORY].append(stem)

        active_categories = {cat: tpls for cat, tpls in self.categorized_templates.items() if tpls}
        self.categorized_templates = active_categories
        self.ui.update_template_categories_dropdown(list(self.categorized_templates.keys()))
        self._log(f"Знайдено {len(all_docx_files)} шаблонів.", 'INFO')

    def on_template_category_changed(self, category_name):
        if not category_name: self.ui.populate_template_list([]); return
        template_stems = self.categorized_templates.get(category_name, [])
        self.ui.populate_template_list(sorted(template_stems))
        self._log(f"Відображено шаблони з категорії '{category_name}'.", 'INFO')

    def choose_custom_templates(self):
        choice = self.ui.show_add_templates_dialog()
        if choice == 'files':
            files, _ = QFileDialog.getOpenFileNames(self.ui, "Виберіть файли шаблонів", "", "Word Files (*.docx)")
            if files:
                for file_path in files: self._copy_template(pathlib.Path(file_path))
        elif choice == 'folder':
            folder = QFileDialog.getExistingDirectory(self.ui, "Виберіть папку для сканування")
            if folder:
                folder_path = pathlib.Path(folder)
                self._log(f"Сканування папки: {folder_path}", 'PROCESS')
                for file_path in folder_path.rglob("*.docx"): self._copy_template(file_path)
        else:
            self._log("Операцію додавання скасовано.", 'WARNING'); return
        self.scan_and_categorize_templates()

    def _copy_template(self, source_path):
        try:
            destination = self.example_dir / source_path.name
            shutil.copyfile(source_path, destination)
            self._log(f"Імпортовано шаблон: {source_path.name}", 'INFO')
        except Exception as e:
            self._log(f"Не вдалося скопіювати {source_path.name}: {e}", 'ERROR')
    
    def select_excel_file(self):
        file, _ = QFileDialog.getOpenFileName(self.ui, "Вибір Excel файлу", str(pathlib.Path.home() / 'Desktop'), "Excel Files (*.xlsx)")
        if not file: self._log("Вибір файлу скасовано.", "WARNING"); return
        
        try:
            self.excel_path = file
            self._log(f"Вибраний Excel файл: {self.excel_path}", 'INFO')
            xls = pd.ExcelFile(self.excel_path)
            sheet_names = xls.sheet_names
            self.ui.update_sheet_dropdown(sheet_names)
            self.ui.set_file_loaded_state(True)
            self.ui.set_sheet_loaded_state(False)
            if sheet_names: self.ui.sheet_dropdown.setCurrentText(sheet_names[0])
        except Exception as e:
            self.excel_path = None
            self._log(f"Помилка читання Excel файлу: {e}", 'ERROR')
            QMessageBox.critical(self.ui, "Error", f"Під час читання файлу Excel сталася помилка: {e}")
            self.ui.set_file_loaded_state(False)
            self.ui.set_sheet_loaded_state(False)

    def load_sheet_data(self, sheet_name):
        if not self.excel_path or not sheet_name or sheet_name == "...":
            self.df = None
            self.ui.update_preview_table(None)
            self.ui.set_sheet_loaded_state(False); return
        
        try:
            self.df = pd.read_excel(self.excel_path, sheet_name=sheet_name, dtype=str)
            self.df.fillna('', inplace=True)
            self.ui.update_preview_table(self.df)
            self._log(f"Завантажено аркуш '{sheet_name}'.", 'INFO')
            self.ui.set_sheet_loaded_state(True)
            # Reset score mappings when a new sheet is loaded
            self.score_mappings = []
            self.ui.update_configured_scores_display(self.score_mappings)
        except Exception as e:
            self.df = None
            self.ui.update_preview_table(None)
            self._log(f"Помилка завантаження аркуша '{sheet_name}': {e}", 'ERROR')
            QMessageBox.critical(self.ui, "Error", f"Не вдалося завантажити аркуш '{sheet_name}': {e}")
            self.ui.set_sheet_loaded_state(False)

    def map_columns(self):
        if self.df is None or self.df.empty:
            QMessageBox.warning(self.ui, "Аркуш не вибрано", "Спочатку виберіть файл Excel та аркуш."); return
        
        dialog = ColumnMappingDialog(self.df, self.required_columns, self.column_mappings, self, self.ui)
        if dialog.exec():
            self.column_mappings = dialog.get_mapped_columns()
            self.df = dialog.get_modified_df()
            self.ui.update_preview_table(self.df)
            self._log("Співставлення колонок та дані оновлено.", 'INFO')

    def open_score_mapping_dialog(self):
        if self.df is None or self.df.empty:
            self._log("Спочатку завантажте дані з Excel.", 'WARNING'); return

        numeric_df = self.df.apply(pd.to_numeric, errors='coerce')
        numeric_columns = numeric_df.dropna(axis=1, how='all').columns.tolist()

        dialog = ScoreMappingDialog(numeric_columns, self.score_mappings, self.ui)
        if dialog.exec():
            self.score_mappings = dialog.get_score_mappings()
            self.ui.update_configured_scores_display(self.score_mappings)
            self._log("Налаштування колонок з балами оновлено.", 'INFO')

    def generate_documents(self):
        try:
            if self.df is None or self.df.empty: self._log("Не вибрано файл Excel або аркуш.", 'ERROR'); return
            selected_templates_names = [item.text() for item in self.ui.template_listbox.selectedItems()]
            if not selected_templates_names: self._log("Не вибрано жодного шаблону.", 'ERROR'); return
            if self.ui.include_scores_checkbox.isChecked() and not self.score_mappings:
                self._log("Включено оцінки, але не налаштовано відповідні стовпці.", 'ERROR'); return
            
            output_dir = QFileDialog.getExistingDirectory(self.ui, "Виберіть папку для збереження документів")
            if not output_dir: self._log("Генерацію скасовано.", 'WARNING'); return

            self._log("Обробка даних...", 'PROCESS')
            processed_df = self._process_data(self.df, self.column_mappings)
            if processed_df is None: self._log("Генерацію зупинено.", 'ERROR'); return

            self._log(f"Створення {len(processed_df) * len(selected_templates_names)} документів...", 'PROCESS')
            self._create_documents(processed_df, selected_templates_names, output_dir)
            self._log(f"Документи успішно створено в {output_dir}", 'INFO')
        except Exception as e:
            self._log(f"Критична помилка під час створення документів: {e}", 'ERROR')
            QMessageBox.critical(self.ui, "Помилка", f"Сталася помилка: {e}")

    def _process_data(self, source_df, mappings):
        df = pd.DataFrame()
        for required, original in mappings.items():
            if original in source_df.columns:
                template_key = self.TEMPLATE_KEYWORDS.get(required, required)
                df[template_key] = source_df[original]
        
        for required in self.required_columns:
            template_key = self.TEMPLATE_KEYWORDS.get(required, required)
            if template_key not in df.columns: df[template_key] = ''
        
        def format_date_col(column_key):
            if column_key not in df.columns: return pd.Series([''] * len(df))
            return pd.to_datetime(df[column_key], errors='coerce').dt.strftime('%d.%m.%Y').fillna('')
        
        df["data_sv"] = format_date_col("data_sv")
        df["data_prot"] = format_date_col("data_prot")
        df["zayava_vid"] = format_date_col("zayava_vid")
        df["data"] = format_date_col("data")
        df["data_vstup"] = format_date_col("data_vstup")
        df["data_nakaz"] = format_date_col("data_nakaz")
        
        today = datetime.today()
        df["d"] = today.strftime("%d")
        df["m"] = format_datetime(today, "LLLL", locale='uk_UA')
        df["y"] = today.strftime("%Y")

        if self.ui.include_scores_checkbox.isChecked():
            for score_map in self.score_mappings:
                col = score_map['source']
                if col in source_df.columns:
                    numeric_scores = pd.to_numeric(source_df[col], errors='coerce')
                    # Add numeric value
                    df[score_map['key']] = numeric_scores.fillna('')
                    # Add written value if requested
                    if score_map['add_written']:
                        df[score_map['written_key']] = numeric_scores.apply(
                            lambda x: num2words(x, lang='uk') if pd.notnull(x) else ''
                        )
                else:
                    self._log(f"Увага: колонка з оцінками '{col}' не знайдена в даних.", 'WARNING')
        return df

    def _create_documents(self, df, selected_templates, output_dir):
        df = df.fillna('')
        for idx, row in df.iterrows():
            context = row.to_dict()
            name1 = str(context.get("name1", f"студент{idx}"))
            name2 = str(context.get("name2", ""))
            name3 = str(context.get("name3", ""))
            file_name_base = f"{name1} {name2} {name3}".strip().replace("  ", " ")
            for template_name in selected_templates:
                template_path = self.template_paths.get(template_name)
                if not template_path or not template_path.exists():
                    self._log(f"Шаблон '{template_name}' не знайдено: {template_path}", 'ERROR'); continue
                try:
                    doc = DocxTemplate(template_path)
                    doc.render(context)
                    output_path = pathlib.Path(output_dir) / f"{template_name}_{file_name_base}.docx"
                    doc.save(output_path)
                except Exception as e:
                    self._log(f"Помилка генерації для {file_name_base} з шаблоном {template_name}: {e}", 'ERROR')

    def handle_test_file_generation(self, row_data, current_mappings, template_name):
        try:
            self._log(f"Запуск генерації тестового файлу з шаблоном '{template_name}'...", 'PROCESS')
            
            test_df = pd.DataFrame([row_data])
            processed_df = self._process_data(test_df, current_mappings)

            if processed_df is None or processed_df.empty:
                self._log("Не вдалося обробити дані для тестового файлу.", 'ERROR'); return

            output_dir = pathlib.Path.home() / 'Desktop'
            context = processed_df.iloc[0].to_dict()
            template_path = self.template_paths.get(template_name)

            if not template_path or not template_path.exists():
                self._log(f"Шаблон '{template_name}' не знайдено.", 'ERROR'); return

            doc = DocxTemplate(template_path)
            doc.render(context)
            output_path = output_dir / f"Test_{template_name}.docx"
            doc.save(output_path)

            self._log(f"Тестовий файл '{output_path.name}' успішно створено на робочому столі.", 'INFO')
            QMessageBox.information(self.ui, "Успіх", f"Тестовий файл '{output_path.name}' було збережено на вашому робочому столі.")

        except Exception as e:
            self._log(f"Помилка під час створення тестового файлу: {e}", 'ERROR')
            QMessageBox.critical(self.ui, "Помилка", f"Під час створення тестового файлу сталася помилка: {e}")

    def handle_document_preview(self, row_data, current_mappings, template_name):
        """
        Generates a single document in-memory, converts it to HTML, and returns the HTML.
        Returns None on failure.
        """
        try:
            self._log(f"Створення попереднього перегляду для шаблону '{template_name}'...", 'PROCESS')
            
            test_df = pd.DataFrame([row_data])
            processed_df = self._process_data(test_df, current_mappings)

            if processed_df is None or processed_df.empty:
                self._log("Не вдалося обробити дані для попереднього перегляду.", 'ERROR')
                return None

            context = processed_df.iloc[0].to_dict()
            template_path = self.template_paths.get(template_name)

            if not template_path or not template_path.exists():
                self._log(f"Шаблон '{template_name}' не знайдено.", 'ERROR')
                return None

            doc = DocxTemplate(template_path)
            doc.render(context)
            
            # Save the document to an in-memory stream
            doc_stream = io.BytesIO()
            doc.save(doc_stream)
            doc_stream.seek(0) # Rewind the stream to the beginning

            # Convert the in-memory document to HTML
            # This preserves as much formatting as mammoth can convert to inline CSS
            result = mammoth.convert_to_html(doc_stream)
            html = result.value # The generated HTML

            # --- KEY CHANGE: Make existing placeholders interactive ---
            # Wrap each placeholder in a styled <a> tag to make it clickable
            for key in self.TEMPLATE_KEYWORDS.values():
                placeholder = f"{{{key}}}"
                # Style the link to make it visible but not distracting
                # The href attribute will be used to identify the clicked keyword
                interactive_placeholder = (
                    f'<a href="placeholder:{key}" style="text-decoration: none; color: inherit; '
                    f'font-weight: bold; background-color: rgba(52, 152, 219, 0.2); border-radius: 2px; padding: 0 2px;">{placeholder}</a>'
                )
                html = html.replace(placeholder, interactive_placeholder)

            return html

        except Exception as e:
            self._log(f"Помилка під час створення попереднього перегляду: {e}", 'ERROR')
            # Do not show a QMessageBox here, as the UI will handle it.
            return None