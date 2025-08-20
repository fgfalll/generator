from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QComboBox,
                             QTableWidget, QTableWidgetItem, QMessageBox, QScrollArea, QFormLayout, QWidget,
                             QGroupBox, QGridLayout, QLineEdit, QInputDialog, QFileDialog, QCheckBox)
import pandas as pd

class PreviewWindow(QDialog):
    """
    A dedicated, non-modal window for previewing and editing the DataFrame.
    It communicates with the main ColumnMappingDialog to ensure data is synchronized.
    """
    def __init__(self, main_dialog):
        super().__init__(main_dialog)
        self.setWindowTitle("Попередній перегляд та редагування даних")
        self.setGeometry(150, 150, 1200, 600)
        
        self.main_dialog = main_dialog
        layout = QVBoxLayout(self)
        
        self.preview_table = QTableWidget(self)
        layout.addWidget(self.preview_table)
        
        buttons_layout = QHBoxLayout()
        add_col_button = QPushButton("Додати колонку")
        remove_col_button = QPushButton("Видалити колонку")
        save_as_button = QPushButton("Зберегти як...")
        
        buttons_layout.addWidget(add_col_button); buttons_layout.addWidget(remove_col_button)
        buttons_layout.addStretch(); buttons_layout.addWidget(save_as_button)
        layout.addLayout(buttons_layout)
        
        add_col_button.clicked.connect(self._add_column)
        remove_col_button.clicked.connect(self._remove_column)
        save_as_button.clicked.connect(self._save_as)
        self.preview_table.itemChanged.connect(self._handle_item_changed)
        
        self.refresh_table()

    def refresh_table(self):
        self.preview_table.blockSignals(True)
        self.preview_table.clear()
        df = self.main_dialog.df
        self.preview_table.setRowCount(df.shape[0])
        self.preview_table.setColumnCount(df.shape[1])
        self.preview_table.setHorizontalHeaderLabels(df.columns)
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item_data = df.iloc[i, j]
                self.preview_table.setItem(i, j, QTableWidgetItem(str(item_data) if pd.notna(item_data) else ""))
        self.preview_table.resizeColumnsToContents()
        self.preview_table.blockSignals(False)

    def _handle_item_changed(self, item):
        try: self.main_dialog.df.iloc[item.row(), item.column()] = item.text()
        except Exception: pass

    def _add_column(self):
        name, ok = QInputDialog.getText(self, "Додати колонку", "Введіть назву нової колонки:")
        if ok and name:
            if name in self.main_dialog.df.columns:
                QMessageBox.warning(self, "Помилка", f"Колонка '{name}' вже існує."); return
            self.main_dialog.df[name] = ''
            self.main_dialog._invalidate_undo_state() # Invalidate undo on data change
            self.refresh_table()
            self.main_dialog._update_all_combos()

    def _remove_column(self):
        current_col_index = self.preview_table.currentColumn()
        if current_col_index == -1:
            QMessageBox.warning(self, "Колонку не вибрано", "Клацніть на клітинку в колонці, яку хочете видалити."); return
        column_name = self.main_dialog.df.columns[current_col_index]
        reply = QMessageBox.question(self, "Видалити колонку", f"Видалити колонку '{column_name}'?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.main_dialog.df.drop(columns=[column_name], inplace=True)
            self.main_dialog._invalidate_undo_state() # Invalidate undo on data change
            self.refresh_table()
            self.main_dialog._update_all_combos()

    def _save_as(self):
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти файл як...", "", "Excel Files (*.xlsx)")
        if path:
            try:
                self.main_dialog.df.to_excel(path, index=False)
                QMessageBox.information(self, "Успіх", f"Файл успішно збережено:\n{path}")
            except Exception as e:
                QMessageBox.critical(self, "Помилка збереження", f"Не вдалося зберегти файл: {e}")

class ColumnMappingDialog(QDialog):
    IDEAL_COLUMN_ORDER = [
        "Тип пропозиції", "Назва КП", "Освітній ступінь", "Вступ на основі", "Спеціальність", "Форма навчання",
        "Курс", "Структурний підрозділ", "Чи скорочений термін навчання", "Прізвище", "Ім'я", "По батькові",
        "Контактний номер", "Електронна адреса", "Статус заявки", "Назва групи", "Реєстраційний номер",
        "Номер групи", "Конкурсний бал", "Бютжет чи контракт", "Чи подано оригінал",
        "Чи подано довідку про місцезнаходження оригіналів", "Особа претендує на застосування сільського коефіцієнту",
        "Номер протоколу", "Дата протоколу", "Наказ про зарахування", "Дата наказу", "Дата вступу", "Тип документа",
        "Додаток до типу документу", "Серія документа", "Номер документа", "Дата видачі документа", "Ким видано",
        "Відзнака", "Дата подачі заяви", "Номер зно", "Рік зно", "ЗНО.Українська мова",
        "ЗНО.Українська мова та література", "ЗНО.Історія України", "ЗНО.Математика", "ЗНО.Біологія",
        "ЗНО.Географія", "ЗНО.Англійська мова", "ДПО", "Адреса", "ДПО.Серія", "ДПО.Номер", "ДПО.Ким виданий",
        "ДПО.Дата видачі", "ДПО.Дійсний до", "РНОКПП"
    ]

    def __init__(self, df, required_columns, column_mappings=None, backend=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Перевірка, очищення та співставлення даних")
        self.setGeometry(100, 100, 1000, 850) # Increased height for new section
        
        self.original_df = df.copy()
        self.df = df.copy()
        self.required_columns = required_columns
        self.column_mappings = column_mappings if column_mappings else {}
        self.combo_boxes = {}
        self.backend = backend
        self.filter_widgets = []
        self.preview_window = None
        self.delimiter_map = {", (Кома)": ",", "; (Крапка з комою)": ";", "	 (Табуляція)": "\t", " (Пробіл)": " "}
        self.pre_split_df = None

        main_layout = QVBoxLayout(self)
        
        # --- Mapping Group ---
        mapping_group = QGroupBox("1. Співставлення колонок")
        mapping_layout = QVBoxLayout(mapping_group)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        scroll_content = QWidget(scroll); scroll_content.setObjectName("ScrollContent")
        form_layout = QFormLayout(scroll_content)
        for req_col in self.required_columns:
            combo = QComboBox(); combo.addItem("Пропустити")
            form_layout.addRow(QLabel(req_col), combo)
            self.combo_boxes[req_col] = combo
        scroll.setWidget(scroll_content)
        mapping_layout.addWidget(scroll)
        main_layout.addWidget(mapping_group)
        
        # --- Cleanup Group ---
        cleanup_group = QGroupBox("2. Очищення та трансформація даних")
        main_cleanup_layout = QVBoxLayout(cleanup_group)

        structure_box = QGroupBox("Структура таблиці")
        structure_layout = QHBoxLayout(structure_box)
        conform_button = QPushButton("Сформувати ідеальну структуру")
        structure_layout.addWidget(conform_button)
        main_cleanup_layout.addWidget(structure_box)

        filters_box = QGroupBox("Фільтрація рядків")
        filters_v_layout = QVBoxLayout(filters_box)
        self.filters_layout = QVBoxLayout()
        add_filter_button = QPushButton("Додати фільтр")
        apply_filters_button = QPushButton("Застосувати всі фільтри")
        filters_v_layout.addLayout(self.filters_layout)
        filters_v_layout.addWidget(add_filter_button)
        filters_v_layout.addWidget(apply_filters_button)
        main_cleanup_layout.addWidget(filters_box)

        split_box = QGroupBox("Розділення колонки")
        split_layout = QGridLayout(split_box)
        split_layout.addWidget(QLabel("Розділити:"), 0, 0); self.split_column_combo = QComboBox(); split_layout.addWidget(self.split_column_combo, 0, 1)
        split_layout.addWidget(QLabel("Роздільник:"), 0, 2)
        delimiter_layout = QHBoxLayout(); self.delimiter_combo = QComboBox(); self.delimiter_combo.addItems(list(self.delimiter_map.keys()) + ["Інший..."])
        self.custom_delimiter_edit = QLineEdit(); self.custom_delimiter_edit.setVisible(False)
        delimiter_layout.addWidget(self.delimiter_combo, 1); delimiter_layout.addWidget(self.custom_delimiter_edit, 1)
        split_layout.addLayout(delimiter_layout, 0, 3)
        split_layout.addWidget(QLabel("Нові назви:"), 1, 0); self.new_column_names_edit = QLineEdit(); self.new_column_names_edit.setPlaceholderText("Ім'я,По батькові")
        split_layout.addWidget(self.new_column_names_edit, 1, 1, 1, 3)
        split_buttons_layout = QHBoxLayout()
        apply_split_button = QPushButton("Розділити")
        self.revert_split_button = QPushButton("Скасувати розділення"); self.revert_split_button.setVisible(False)
        split_buttons_layout.addWidget(apply_split_button); split_buttons_layout.addWidget(self.revert_split_button)
        split_layout.addLayout(split_buttons_layout, 0, 4, 2, 1)
        main_cleanup_layout.addWidget(split_box)

        # --- KEY FEATURE: New group for merging columns ---
        merge_box = QGroupBox("Об'єднання колонок")
        merge_layout = QGridLayout(merge_box)
        self.merge_column_names_edit = QLineEdit(); self.merge_column_names_edit.setPlaceholderText("Прізвище, Ім'я, По батькові")
        self.new_merged_column_name_edit = QLineEdit(); self.new_merged_column_name_edit.setPlaceholderText("ПІБ")
        self.merge_delimiter_edit = QLineEdit(); self.merge_delimiter_edit.setPlaceholderText(" (пробіл)")
        self.delete_original_cols_check = QCheckBox("Видалити вихідні колонки"); self.delete_original_cols_check.setChecked(True)
        apply_merge_button = QPushButton("Об'єднати")
        merge_layout.addWidget(QLabel("Колонки для об'єднання:"), 0, 0); merge_layout.addWidget(self.merge_column_names_edit, 0, 1)
        merge_layout.addWidget(QLabel("Нова назва колонки:"), 1, 0); merge_layout.addWidget(self.new_merged_column_name_edit, 1, 1)
        merge_layout.addWidget(QLabel("Роздільник:"), 2, 0); merge_layout.addWidget(self.merge_delimiter_edit, 2, 1)
        merge_layout.addWidget(self.delete_original_cols_check, 3, 1)
        merge_layout.addWidget(apply_merge_button, 0, 2, 4, 1)
        main_cleanup_layout.addWidget(merge_box)
        
        main_layout.addWidget(cleanup_group)

        # --- Preview and Main Buttons ---
        preview_group = QGroupBox("3. Попередній перегляд")
        preview_layout = QHBoxLayout(preview_group)
        self.preview_button = QPushButton("Попередній перегляд та редагування")
        preview_layout.addWidget(self.preview_button)
        main_layout.addWidget(preview_group)

        buttons_layout = QHBoxLayout()
        test_button = QPushButton("Згенерувати тест"); reset_button = QPushButton("Скинути зміни")
        ok_button = QPushButton("Зберегти та закрити"); cancel_button = QPushButton("Скасувати")
        buttons_layout.addWidget(test_button); buttons_layout.addWidget(reset_button)
        buttons_layout.addStretch(); buttons_layout.addWidget(ok_button); buttons_layout.addWidget(cancel_button)
        main_layout.addLayout(buttons_layout)
        
        # --- Connections ---
        ok_button.clicked.connect(self.accept); cancel_button.clicked.connect(self.reject)
        reset_button.clicked.connect(self._reset_data); test_button.clicked.connect(self._generate_test_file)
        apply_split_button.clicked.connect(self._split_column)
        apply_merge_button.clicked.connect(self._merge_columns) # Connect new button
        self.revert_split_button.clicked.connect(self._revert_split)
        add_filter_button.clicked.connect(self._add_filter_row)
        apply_filters_button.clicked.connect(self._apply_all_filters)
        self.preview_button.clicked.connect(self._open_preview_window)
        self.delimiter_combo.currentTextChanged.connect(self._on_delimiter_changed)
        conform_button.clicked.connect(self._conform_to_ideal_structure)
        
        self._add_filter_row(is_deletable=False)
        self._update_all_widgets()
        self._automap_columns()

    def _invalidate_undo_state(self):
        self.pre_split_df = None
        self.revert_split_button.setVisible(False)
        
    def _conform_to_ideal_structure(self):
        reply = QMessageBox.question(self, "Сформувати ідеальну структуру?",
                                     "Ця дія додасть відсутні ідеальні колонки, видалить усі зайві та змінить порядок існуючих.\n\n"
                                     "Дані в зайвих колонках буде втрачено. Цю дію не можна скасувати.\n\n"
                                     "Продовжити?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.No: return
        self._invalidate_undo_state()
        new_df = pd.DataFrame(index=self.df.index)
        for col_name in self.IDEAL_COLUMN_ORDER:
            new_df[col_name] = self.df[col_name] if col_name in self.df.columns else ""
        self.df = new_df
        self._update_all_widgets()
        self._automap_columns()
        QMessageBox.information(self, "Успіх", "Структуру таблиці було приведено до ідеального формату.")

    def _on_delimiter_changed(self, text):
        self.custom_delimiter_edit.setVisible(text == "Інший...")

    def _open_preview_window(self):
        if not self.preview_window or not self.preview_window.isVisible():
            self.preview_window = PreviewWindow(self)
            self.preview_window.show()
        else:
            self.preview_window.raise_(); self.preview_window.activateWindow()

    def _add_filter_row(self, is_deletable=True):
        fw_widget = QWidget()
        row = QHBoxLayout(fw_widget); row.setContentsMargins(0,0,0,0)
        col = QComboBox(); cond = QComboBox(); cond.addItems(["містить", "не містить", "дорівнює", "не дорівнює"])
        val = QLineEdit(); val.setPlaceholderText("Значення")
        row.addWidget(col); row.addWidget(cond); row.addWidget(val, 1)
        fw = {"row": fw_widget, "col_combo": col, "cond_combo": cond, "val_edit": val}
        if is_deletable:
            rem_btn = QPushButton("Видалити"); rem_btn.clicked.connect(lambda: self._remove_filter_row(fw))
            row.addWidget(rem_btn)
        self.filters_layout.addWidget(fw_widget); self.filter_widgets.append(fw)
        self._update_filter_combos()

    def _remove_filter_row(self, fw):
        if fw in self.filter_widgets: self.filter_widgets.remove(fw); fw["row"].deleteLater()

    def _apply_all_filters(self):
        self._invalidate_undo_state()
        temp_df = self.df.copy()
        for fw in self.filter_widgets:
            col, val = fw["col_combo"].currentText(), fw["val_edit"].text()
            if not col or not val: continue
            cond = fw["cond_combo"].currentText(); series = temp_df[col].astype(str)
            if cond == "містить": mask = series.str.contains(val, case=False, na=False)
            elif cond == "не містить": mask = ~series.str.contains(val, case=False, na=False)
            elif cond == "дорівнює": mask = series.str.lower() == val.lower()
            else: mask = series.str.lower() != val.lower()
            temp_df = temp_df[mask]
        self.df = temp_df.reset_index(drop=True)
        self._notify_preview_window()
        QMessageBox.information(self, "Фільтри застосовано", f"Залишилось рядків: {len(self.df)}")

    def _notify_preview_window(self):
        if self.preview_window and self.preview_window.isVisible(): self.preview_window.refresh_table()

    def _update_all_widgets(self):
        self._update_all_combos()
        self._notify_preview_window()
        
    def _update_all_combos(self):
        cols = self.df.columns.tolist()
        for combo in list(self.combo_boxes.values()) + [self.split_column_combo]:
            curr = combo.currentText()
            combo.blockSignals(True); combo.clear()
            if combo in self.combo_boxes.values(): combo.addItem("Пропустити")
            combo.addItems(cols)
            if curr in cols: combo.setCurrentText(curr)
            elif combo in self.combo_boxes.values(): combo.setCurrentIndex(0)
            combo.blockSignals(False)
        self._update_filter_combos()

    def _update_filter_combos(self):
        cols = self.df.columns.tolist()
        for fw in self.filter_widgets:
            curr = fw["col_combo"].currentText()
            fw["col_combo"].blockSignals(True); fw["col_combo"].clear(); fw["col_combo"].addItems(cols)
            if curr in cols: fw["col_combo"].setCurrentText(curr)
            fw["col_combo"].blockSignals(False)
    
    def _automap_columns(self):
        for req in self.required_columns:
            if req in self.df.columns and req in self.combo_boxes: self.combo_boxes[req].setCurrentText(req)

    def _reset_data(self):
        reply = QMessageBox.question(self, "Скинути всі зміни", "Відновити дані до стану на момент відкриття вікна?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self._invalidate_undo_state()
            self.df = self.original_df.copy()
            for fw in self.filter_widgets: fw["row"].deleteLater()
            self.filter_widgets.clear()
            self._add_filter_row(is_deletable=False)
            self._update_all_widgets()
            self._automap_columns()

    def _split_column(self):
        col_split = self.split_column_combo.currentText(); delim_txt = self.delimiter_combo.currentText()
        names_str = self.new_column_names_edit.text()
        delim = self.custom_delimiter_edit.text() if delim_txt == "Інший..." else self.delimiter_map.get(delim_txt)
        if not all([col_split, delim is not None, names_str]):
            QMessageBox.warning(self, "Недостатньо даних", "Заповніть усі поля для розділення."); return
        names = [n.strip() for n in names_str.split(',') if n.strip()]
        if not names: QMessageBox.warning(self, "Недостатньо даних", "Вкажіть назви нових колонок."); return
        for name in names:
            if name in self.df.columns and name != col_split:
                QMessageBox.warning(self, "Конфлікт імен", f"Колонка з назвою '{name}' вже існує."); return
        try:
            self.pre_split_df = self.df.copy()
            original_col_index = self.df.columns.get_loc(col_split)
            split_data = self.df[col_split].astype(str).str.split(delim, n=len(names) - 1, expand=True)
            self.df.drop(columns=[col_split], inplace=True)
            for i, name in reversed(list(enumerate(names))):
                column_data = split_data[i].fillna('') if i < split_data.shape[1] else [''] * len(self.df)
                self.df.insert(original_col_index, name, column_data)
            self._update_all_widgets()
            self.revert_split_button.setVisible(True)
            QMessageBox.information(self, "Успіх", f"Колонку '{col_split}' було видалено та замінено новими.")
        except Exception as e:
            QMessageBox.critical(self, "Помилка", f"Сталася помилка під час розділення: {e}"); self.pre_split_df = None

    def _revert_split(self):
        if self.pre_split_df is not None:
            self.df = self.pre_split_df.copy()
            self._invalidate_undo_state()
            self._update_all_widgets()
            QMessageBox.information(self, "Скасовано", "Операцію розділення колонки було скасовано.")
        else:
            QMessageBox.warning(self, "Неможливо скасувати", "Немає операції розділення для скасування.")
            
    def _merge_columns(self):
        cols_str = self.merge_column_names_edit.text()
        new_col_name = self.new_merged_column_name_edit.text().strip()
        delimiter = self.merge_delimiter_edit.text()
        
        if not all([cols_str, new_col_name]):
            QMessageBox.warning(self, "Недостатньо даних", "Вкажіть колонки для об'єднання та назву нової колонки."); return
        
        cols_to_merge = [c.strip() for c in cols_str.split(',') if c.strip()]
        if len(cols_to_merge) < 2:
            QMessageBox.warning(self, "Помилка", "Для об'єднання потрібно щонайменше дві колонки."); return
        
        for col in cols_to_merge:
            if col not in self.df.columns:
                QMessageBox.warning(self, "Помилка", f"Колонку '{col}' не знайдено в таблиці."); return
        
        if new_col_name in self.df.columns and new_col_name not in cols_to_merge:
            QMessageBox.warning(self, "Конфлікт імен", f"Колонка з назвою '{new_col_name}' вже існує."); return
        
        try:
            self._invalidate_undo_state() # Any merge invalidates the split-undo
            
            # Combine columns
            self.df[new_col_name] = self.df[cols_to_merge].astype(str).agg(delimiter.join, axis=1)

            # Delete original columns if requested
            if self.delete_original_cols_check.isChecked():
                cols_to_drop = [c for c in cols_to_merge if c != new_col_name]
                if cols_to_drop:
                    self.df.drop(columns=cols_to_drop, inplace=True)

            self._update_all_widgets()
            QMessageBox.information(self, "Успіх", f"Колонки було успішно об'єднано в '{new_col_name}'.")

        except Exception as e:
            QMessageBox.critical(self, "Помилка", f"Сталася помилка під час об'єднання: {e}")

    def _generate_test_file(self):
        if not self.backend: QMessageBox.critical(self, "Помилка", "Зв'язок з додатком втрачено."); return
        templates = list(self.backend.template_paths.keys())
        if not templates: QMessageBox.warning(self, "Немає шаблонів", "Не знайдено шаблонів."); return
        sel_row = -1
        if self.preview_window and self.preview_window.isVisible(): sel_row = self.preview_window.preview_table.currentRow()
        if sel_row < 0: QMessageBox.warning(self, "Рядок не вибрано", "Відкрийте попередній перегляд та виберіть рядок."); return
        
        template, ok = QInputDialog.getItem(self, "Вибір шаблону", "Виберіть шаблон:", sorted(templates), 0, False)
        if ok and template:
            row_data = self.df.iloc[sel_row]
            mappings = self._get_current_mappings_from_ui()
            self.backend.handle_test_file_generation(row_data, mappings, template)

    def _get_current_mappings_from_ui(self):
        return {r: c.currentText() for r, c in self.combo_boxes.items() if c.currentText() != "Пропустити"}

    def accept(self):
        mappings = self._get_current_mappings_from_ui()
        if len(mappings.values()) != len(set(mappings.values())):
            QMessageBox.warning(self, "Дублювання колонок", "Одна колонка не може бути призначена для кількох полів."); return
        self.column_mappings = mappings
        super().accept()

    def closeEvent(self, event):
        if self.preview_window: self.preview_window.close()
        super().closeEvent(event)

    def get_mapped_columns(self): return self.column_mappings
    def get_modified_df(self): return self.df