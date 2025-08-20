from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
                             QDialogButtonBox, QWidget, QPushButton, QColorDialog,
                             QScrollArea, QFormLayout, QGroupBox)
from PyQt6.QtCore import pyqtSignal
from copy import deepcopy
import theme_manager

class ColorPickerButton(QPushButton):
    """A button that displays a color and opens a QColorDialog when clicked."""
    color_changed = pyqtSignal(str)

    def __init__(self, initial_color, parent=None):
        super().__init__(parent)
        self.setFixedSize(60, 28)
        self.clicked.connect(self.pick_color)
        self._color = None
        self.set_color(initial_color)

    def set_color(self, color_hex):
        from PyQt6.QtGui import QColor
        self._color = QColor(color_hex)
        self.update_style()

    def get_color(self):
        return self._color.name()

    def pick_color(self):
        dialog = QColorDialog(self._color, self)
        dialog.setOption(QColorDialog.ColorDialogOption.DontUseNativeDialog, True)
        
        if dialog.exec():
            self._color = dialog.currentColor()
            self.update_style()
            self.color_changed.emit(self.get_color())

    def update_style(self):
        self.setStyleSheet(f"background-color: {self.get_color()}; border: 1px solid grey; border-radius: 2px;")

class SettingsDialog(QDialog):
    def __init__(self, current_theme, custom_colors, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Налаштування")
        self.setMinimumSize(450, 400)
        
        self.base_themes = theme_manager.THEMES
        self.current_base_theme = current_theme
        self.working_colors = deepcopy(custom_colors)

        main_layout = QVBoxLayout(self)

        preset_layout = QHBoxLayout()
        preset_layout.addWidget(QLabel("Базова тема:"))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(self.base_themes.keys())
        self.theme_combo.setCurrentText(self.current_base_theme)
        preset_layout.addWidget(self.theme_combo)
        main_layout.addLayout(preset_layout)
        
        custom_colors_group = QGroupBox("Користувацькі кольори")
        custom_group_layout = QVBoxLayout(custom_colors_group)
        main_layout.addWidget(custom_colors_group)

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        form_widget = QWidget(); form_widget.setObjectName("ColorPickerForm")
        form_layout = QFormLayout(form_widget)
        self.color_pickers = {}
        
        for key, description in theme_manager.COLOR_DESCRIPTIONS.items():
            base_color = self.base_themes[self.current_base_theme][key]
            current_color = self.working_colors.get(key, base_color)
            picker = ColorPickerButton(current_color)
            picker.color_changed.connect(lambda color, k=key: self.on_color_changed(k, color))
            self.color_pickers[key] = picker
            form_layout.addRow(description, picker)
        
        scroll.setWidget(form_widget)
        custom_group_layout.addWidget(scroll)

        reset_button = QPushButton("Скинути налаштування")
        custom_group_layout.addWidget(reset_button)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        main_layout.addWidget(self.button_box)

        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.theme_combo.currentTextChanged.connect(self.on_base_theme_changed)
        reset_button.clicked.connect(self.reset_custom_colors)

    def on_base_theme_changed(self, theme_name):
        self.current_base_theme = theme_name
        self.working_colors = {}
        self.update_color_pickers()

    def reset_custom_colors(self):
        self.working_colors = {}
        self.update_color_pickers()

    def on_color_changed(self, key, new_color):
        self.working_colors[key] = new_color

    def update_color_pickers(self):
        for key, picker in self.color_pickers.items():
            base_color = self.base_themes[self.current_base_theme][key]
            picker.set_color(self.working_colors.get(key, base_color))

    def get_settings(self):
        return {
            "theme_name": self.current_base_theme,
            "custom_colors": self.working_colors
        }