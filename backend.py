import shutil
import pathlib
import pandas as pd
import zipfile
import re
import json
import xml.etree.ElementTree as ET
from PyQt6.QtCore import QObject
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from docxtpl import DocxTemplate
from num2words import num2words
from babel.dates import format_datetime
from datetime import datetime

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
        self.score_mappings = []
        
        self.categorized_templates = {}
        self.template_paths = {}

        self.base_dir = pathlib.Path(__file__).resolve().parent
        self.example_dir = self.base_dir / "Приклади"
        self.example_dir.mkdir(parents=True, exist_ok=True)
        self.keywords_file = self.base_dir / "keywords.json"

        self.TEMPLATE_KEYWORDS = {}
        self.required_columns = []
        self._load_keywords()

        self.CATEGORY_KEYWORDS = {
            'Аркуш випробувань': ['випробувань', 'тестовий', 'test'],
            'Витяг з наказу': ['витяг', 'наказ', 'order'],
            'Опис справи': ['опис', 'справ', 'case'],
            'Повідомлення': ['повідомлення', 'лист', 'notification']
        }
        self.DEFAULT_CATEGORY = 'Різне'

    def _log(self, message, level='INFO'):
        self.ui.log_message(message, level)

    def _load_keywords(self):
        default_keywords = {
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
        if not self.keywords_file.exists():
            try:
                with open(self.keywords_file, 'w', encoding='utf-8') as f:
                    json.dump(default_keywords, f, indent=4, ensure_ascii=False)
                self.TEMPLATE_KEYWORDS = default_keywords
            except Exception as e:
                self._log(f"Не вдалося створити файл ключових слів: {e}", 'ERROR')
                self.TEMPLATE_KEYWORDS = default_keywords
        else:
            try:
                with open(self.keywords_file, 'r', encoding='utf-8') as f:
                    self.TEMPLATE_KEYWORDS = json.load(f)
            except Exception as e:
                self._log(f"Помилка читання файлу ключових слів: {e}", 'ERROR')
                self.TEMPLATE_KEYWORDS = default_keywords
        self.required_columns = list(self.TEMPLATE_KEYWORDS.keys())
        self._log(f"Завантажено {len(self.required_columns)} стандартних ключових слів.", 'INFO')

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
            self.app.setStyleSheet(theme_manager.generate_qss(new_settings["theme_name"], new_settings["custom_colors"]))

    def scan_and_categorize_templates(self):
        self.template_paths.clear()
        self.categorized_templates = {cat: [] for cat in self.CATEGORY_KEYWORDS.keys()}
        self.categorized_templates[self.DEFAULT_CATEGORY] = []
        for path in self.example_dir.rglob("*.docx"):
            if not path.is_file(): continue
            stem, assigned = path.stem, False
            self.template_paths[stem] = path
            for category, keywords in self.CATEGORY_KEYWORDS.items():
                if any(keyword in stem.lower() for keyword in keywords):
                    self.categorized_templates[category].append(stem)
                    assigned = True; break
            if not assigned: self.categorized_templates[self.DEFAULT_CATEGORY].append(stem)
        self.categorized_templates = {cat: tpls for cat, tpls in self.categorized_templates.items() if tpls}
        self.ui.update_template_categories_dropdown(list(self.categorized_templates.keys()))

    def on_template_category_changed(self, category_name):
        self.ui.populate_template_list(sorted(self.categorized_templates.get(category_name, [])))

    def choose_custom_templates(self):
        choice = self.ui.show_add_templates_dialog()
        if choice == 'files':
            files, _ = QFileDialog.getOpenFileNames(self.ui, "Виберіть файли шаблонів", "", "Word Files (*.docx)")
            if files: [self._copy_template(pathlib.Path(f)) for f in files]
        elif choice == 'folder':
            folder = QFileDialog.getExistingDirectory(self.ui, "Виберіть папку для сканування")
            if folder: [self._copy_template(f) for f in pathlib.Path(folder).rglob("*.docx")]
        if choice: self.scan_and_categorize_templates()

    def _copy_template(self, source_path):
        try:
            shutil.copyfile(source_path, self.example_dir / source_path.name)
            self._log(f"Імпортовано шаблон: {source_path.name}", 'INFO')
        except Exception as e: self._log(f"Не вдалося скопіювати {source_path.name}: {e}", 'ERROR')
    
    def select_excel_file(self):
        file, _ = QFileDialog.getOpenFileName(self.ui, "Вибір Excel файлу", str(pathlib.Path.home()), "Excel Files (*.xlsx)")
        if not file: return
        try:
            self.excel_path = file
            sheet_names = pd.ExcelFile(self.excel_path).sheet_names
            self.ui.update_sheet_dropdown(sheet_names)
            self.ui.set_file_loaded_state(True)
            if sheet_names: self.ui.sheet_dropdown.setCurrentText(sheet_names[0])
            self._log(f"Вибраний Excel файл: {self.excel_path}", 'INFO')
        except Exception as e:
            self._log(f"Помилка читання Excel файлу: {e}", 'ERROR')

    def load_sheet_data(self, sheet_name):
        if not self.excel_path or not sheet_name or sheet_name == "...":
            self.df, self.score_mappings = None, []
            self.ui.update_preview_table(None); self.ui.set_sheet_loaded_state(False)
            self.ui.update_configured_scores_display([]); return
        try:
            self.df = pd.read_excel(self.excel_path, sheet_name=sheet_name, dtype=str).fillna('')
            self.ui.update_preview_table(self.df)
            self.ui.set_sheet_loaded_state(True)
            self.score_mappings = []
            self.ui.update_configured_scores_display([])
            self._log(f"Завантажено аркуш '{sheet_name}'.", 'INFO')
        except Exception as e:
            self._log(f"Помилка завантаження аркуша '{sheet_name}': {e}", 'ERROR')

    def map_columns(self, missing_columns=None):
        if self.df is None or self.df.empty: return False

        columns_to_show = missing_columns if missing_columns else self.required_columns
        current_mappings = {col: col for col in columns_to_show if col in self.df.columns}
        dialog = ColumnMappingDialog(self.df, columns_to_show, current_mappings, self, self.ui)
        
        if dialog.exec():
            mappings = dialog.get_mapped_columns()
            self.df = dialog.get_modified_df()
            rename_dict = {v: k for k, v in mappings.items() if k != v and v in self.df.columns}
            if rename_dict:
                self.df.rename(columns=rename_dict, inplace=True)
                self._log(f"Перейменовано {len(rename_dict)} колонок.", 'INFO')
            self.ui.update_preview_table(self.df)
            return True
        return False

    def open_score_mapping_dialog(self):
        if self.df is None or self.df.empty: return
        numeric_cols = self.df.apply(pd.to_numeric, errors='coerce').dropna(axis=1, how='all').columns.tolist()
        dialog = ScoreMappingDialog(numeric_cols, self.score_mappings, self.ui)
        if dialog.exec():
            self.score_mappings = dialog.get_score_mappings()
            self.ui.update_configured_scores_display(self.score_mappings)
            self._add_scores_to_dataframe()
            self.ui.update_preview_table(self.df)
            self._log("Таблиця даних оновлена новими колонками балів.", "INFO")

    def _add_scores_to_dataframe(self):
        if self.df is None: return
        def format_score(x):
            if pd.isna(x): return ''
            return str(int(x)) if x == int(x) else str(x)

        for score_map in self.score_mappings:
            source_col = score_map['source']
            if source_col in self.df.columns:
                numeric_scores = pd.to_numeric(self.df[source_col], errors='coerce')
                self.df[score_map['key']] = numeric_scores.apply(format_score)
                if score_map.get('add_written', False):
                    self.df[score_map['written_key']] = numeric_scores.apply(
                        lambda x: num2words(int(x), lang='uk') if pd.notnull(x) and x == int(x) else (num2words(x, lang='uk') if pd.notnull(x) else '')
                    )
            
    def _extract_template_keywords(self, template_path):
        keywords = set()
        try:
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            with zipfile.ZipFile(template_path) as docx:
                xml_content = docx.read('word/document.xml')
                root = ET.fromstring(xml_content)
                full_text = "".join(node.text for node in root.findall('.//w:t', ns) if node.text)
                keywords.update(item.strip() for item in re.findall(r'\{\{(.*?)\}\}', full_text))
        except Exception: pass
        return keywords

    def generate_documents(self):
        try:
            if self.df is None or self.df.empty:
                QMessageBox.warning(self.ui, "Немає даних", "Будь ласка, завантажте дані з файлу Excel."); return
            
            selected_templates = [item.text() for item in self.ui.template_listbox.selectedItems()]
            if not selected_templates:
                QMessageBox.warning(self.ui, "Не вибрано шаблон", "Будь ласка, виберіть хоча б один шаблон у Кроці 2."); return

            while True:
                all_used_keywords = {k for name in selected_templates for k in self._extract_template_keywords(self.template_paths[name])}
                if not all_used_keywords:
                    QMessageBox.warning(self.ui, "Порожні шаблони", "У шаблонах не знайдено полів."); return

                known_score_keys = {m['key'] for m in self.score_mappings} | {m['written_key'] for m in self.score_mappings if m.get('add_written')}
                auto_keys, unconfigured_scores = {'d', 'm', 'y'}, []
                columns_that_need_mapping = []

                for keyword in sorted(all_used_keywords):
                    if keyword in auto_keys or keyword in known_score_keys: continue
                    
                    is_standard = any(v == keyword for v in self.TEMPLATE_KEYWORDS.values())
                    
                    if is_standard:
                        desc = next((k for k, v in self.TEMPLATE_KEYWORDS.items() if v == keyword), None)
                        if desc and desc not in self.df.columns:
                            columns_that_need_mapping.append(desc)
                    elif 'zno_' in keyword or '_baly' in keyword:
                        unconfigured_scores.append(keyword)
                    else: # Custom keyword
                        if keyword not in self.df.columns:
                            columns_that_need_mapping.append(keyword)

                if unconfigured_scores:
                    msg = "У ваших шаблонах знайдено ключові слова, схожі на бали, але вони не налаштовані:\n\n" + "\n".join(f"- {kw}" for kw in unconfigured_scores) + "\n\nВикористайте **'Налаштувати колонки балів'**."
                    QMessageBox.warning(self.ui, "Неналаштовані бали", msg); return
                
                if not columns_that_need_mapping: break

                msg_box = QMessageBox(self.ui)
                msg_box.setWindowTitle("Потрібно налаштувати поля")
                msg_box.setText("Для генерації потрібно вказати відповідність для полів:\n\n" + "\n".join(f"- {col}" for col in columns_that_need_mapping[:15]))
                cfg_btn = msg_box.addButton("Налаштувати...", QMessageBox.ButtonRole.ActionRole)
                gen_btn = msg_box.addButton("Продовжити з порожніми", QMessageBox.ButtonRole.AcceptRole)
                msg_box.addButton("Скасувати", QMessageBox.ButtonRole.RejectRole)
                msg_box.exec()

                if msg_box.clickedButton() == gen_btn: break
                if msg_box.clickedButton() != cfg_btn: return

                if not self.map_columns(missing_columns=columns_that_need_mapping):
                    return

            output_dir = QFileDialog.getExistingDirectory(self.ui, "Виберіть папку для збереження документів")
            if not output_dir: return

            processed_df = self._process_data(self.df.copy())
            if processed_df is None: return

            self._create_documents(processed_df, selected_templates, output_dir)
            QMessageBox.information(self.ui, "Успіх", f"Генерацію завершено!\nДокументи збережено в папці:\n{output_dir}")
        except Exception as e:
            self._log(f"Критична помилка: {e}", 'ERROR')

    def _process_data(self, source_df):
        df = source_df
        self._add_scores_to_dataframe()
        
        date_keys = [k for k in self.required_columns if "Дата" in k and k in df.columns]
        for key in date_keys:
            df[key] = pd.to_datetime(df[key], errors='coerce').dt.strftime('%d.%m.%Y').fillna('')
        
        today = datetime.today()
        df["d"], df["m"], df["y"] = today.strftime("%d"), format_datetime(today, "LLLL", locale='uk_UA'), today.strftime("%Y")
        
        rename_dict = {k: v for k, v in self.TEMPLATE_KEYWORDS.items() if k in df.columns}
        df.rename(columns=rename_dict, inplace=True)
        return df

    def _create_documents(self, df, selected_templates, output_dir):
        df = df.astype(str).fillna('')
        for idx, row in df.iterrows():
            context = row.to_dict()
            name1_key = self.TEMPLATE_KEYWORDS.get("Прізвище", "name1")
            name2_key = self.TEMPLATE_KEYWORDS.get("Ім'я", "name2")
            name3_key = self.TEMPLATE_KEYWORDS.get("По батькові", "name3")
            
            name1 = str(context.get(name1_key, f"студент{idx}"))
            name2 = str(context.get(name2_key, ""))
            name3 = str(context.get(name3_key, ""))
            file_name_base = f"{name1} {name2} {name3}".strip().replace("  ", " ")
            
            for template_name in selected_templates:
                template_path = self.template_paths.get(template_name)
                if not template_path: continue
                
                template_keys = self._extract_template_keywords(template_path)
                final_context = {key: context.get(key, '') for key in template_keys}

                try:
                    doc = DocxTemplate(template_path)
                    doc.render(final_context)
                    doc.save(pathlib.Path(output_dir) / f"{template_name}_{file_name_base}.docx")
                except Exception as e:
                    self._log(f"Помилка генерації для {file_name_base} з {template_name}: {e}", 'ERROR')

    def handle_test_file_generation(self, row_data, current_mappings, template_name):
        try:
            test_df = pd.DataFrame([row_data])
            rename_dict = {v: k for k, v in current_mappings.items()}
            test_df.rename(columns=rename_dict, inplace=True)

            processed_df = self._process_data(test_df)
            if processed_df is None or processed_df.empty: return

            output_dir = pathlib.Path.home() / 'Desktop'
            context = processed_df.iloc[0].to_dict()
            template_path = self.template_paths.get(template_name)
            if not template_path: return

            doc = DocxTemplate(template_path)
            doc.render(context)
            doc.save(output_dir / f"Test_{template_name}.docx")
            QMessageBox.information(self.ui, "Успіх", f"Тестовий файл збережено на робочому столі.")
        except Exception as e:
            self._log(f"Помилка створення тестового файлу: {e}", 'ERROR')