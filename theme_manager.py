import json
import pathlib
from copy import deepcopy

THEMES = {
    "Strict Dark": {
        "main_bg": "#212529", "content_bg": "#343a40", "darker_bg": "#1c1f23",
        "border": "#495057", "primary_text": "#f8f9fa", "secondary_text": "#adb5bd",
        "disabled_text": "#6c757d", "primary_button": "#0d6efd", "primary_hover": "#0b5ed7",
        "primary_pressed": "#0a58ca", "secondary_button": "#6c757d",
        "generate_button": "#198754", "generate_hover": "#157347", "generate_pressed": "#146c43",
    },
    "Strict Light": {
        "main_bg": "#f8f9fa", "content_bg": "#ffffff", "darker_bg": "#e9ecef",
        "border": "#dee2e6", "primary_text": "#212529", "secondary_text": "#6c757d",
        "disabled_text": "#adb5bd", "primary_button": "#0d6efd", "primary_hover": "#0b5ed7",
        "primary_pressed": "#0a58ca", "secondary_button": "#6c757d",
        "generate_button": "#198754", "generate_hover": "#157347", "generate_pressed": "#146c43",
    },
    "Dark": {
        "main_bg": "#2c3e50", "content_bg": "#34495e", "darker_bg": "#273746",
        "border": "#4a6278", "primary_text": "#ecf0f1", "secondary_text": "#bdc3c7",
        "disabled_text": "#95a5a6", "primary_button": "#3498db", "primary_hover": "#4ea9e4",
        "primary_pressed": "#2980b9", "secondary_button": "#566573",
        "generate_button": "#2ecc71", "generate_hover": "#3fe381", "generate_pressed": "#27ae60",
    },
    "Light": {
        "main_bg": "#ecf0f1", "content_bg": "#ffffff", "darker_bg": "#e8ecef",
        "border": "#bdc3c7", "primary_text": "#2c3e50", "secondary_text": "#7f8c8d",
        "disabled_text": "#babdbe", "primary_button": "#3498db", "primary_hover": "#4ea9e4",
        "primary_pressed": "#2980b9", "secondary_button": "#bdc3c7",
        "generate_button": "#27ae60", "generate_hover": "#2ecc71", "generate_pressed": "#229954",
    }
}

COLOR_DESCRIPTIONS = {
    "main_bg": "Основний фон", "content_bg": "Фон контенту", "darker_bg": "Фон консолі/таблиць",
    "border": "Колір рамок", "primary_text": "Основний текст", "secondary_text": "Додатковий текст",
    "primary_button": "Основна кнопка", "generate_button": "Кнопка 'Створити'",
}

CONFIG_FILE = pathlib.Path(__file__).resolve().parent / "config.json"

def get_available_themes():
    return list(THEMES.keys())

def save_settings(theme_name, custom_colors=None):
    settings = {"theme_name": theme_name, "custom_colors": custom_colors or {}}
    with open(CONFIG_FILE, "w") as f:
        json.dump(settings, f, indent=4)

def load_settings():
    defaults = {"theme_name": "Strict Dark", "custom_colors": {}}
    if not CONFIG_FILE.exists():
        return defaults
    try:
        with open(CONFIG_FILE, "r") as f:
            config = json.load(f)
            if config.get("theme_name") not in THEMES:
                config["theme_name"] = defaults["theme_name"]
            return { "theme_name": config.get("theme_name"), "custom_colors": config.get("custom_colors", {}) }
    except (json.JSONDecodeError, IOError):
        return defaults

def generate_qss(theme_name="Strict Dark", custom_colors=None):
    colors = deepcopy(THEMES.get(theme_name, THEMES["Strict Dark"]))
    if custom_colors:
        colors.update(custom_colors)

    return f"""
        QMainWindow, QDialog, QWidget#MainContentWidget {{
            background-color: {colors["main_bg"]};
            font-family: 'Segoe UI', Arial, sans-serif; font-size: 10pt; color: {colors["primary_text"]};
        }}
        QWidget#ColorPickerForm {{ background-color: {colors["content_bg"]}; }}
        
        QMenuBar {{
            background-color: {colors["content_bg"]};
            color: {colors["primary_text"]};
        }}
        QMenuBar::item {{
            padding: 4px 8px;
            background-color: transparent;
        }}
        QMenuBar::item:selected {{
            background-color: {colors["darker_bg"]};
        }}
        QMenu {{
            background-color: {colors["content_bg"]};
            color: {colors["primary_text"]};
            border: 1px solid {colors["border"]};
        }}
        QMenu::item:selected {{
            background-color: {colors["primary_button"]};
            color: white;
        }}

        QGroupBox {{
            background-color: {colors["content_bg"]}; border: 1px solid {colors["border"]};
            border-radius: 2px; margin-top: 1ex; padding: 8px;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin; subcontrol-position: top left; padding: 0 4px;
            left: 10px; color: {colors["primary_text"]}; font-weight: bold;
        }}
        QLabel {{ color: {colors["secondary_text"]}; background-color: transparent; }}
        QLabel#DialogTitle {{ color: {colors["primary_text"]}; font-size: 13pt; font-weight: bold; padding: 8px; }}
        
        QLineEdit, QInputDialog QLineEdit {{
            background-color: {colors["darker_bg"]}; border: 1px solid {colors["border"]};
            border-radius: 2px; padding: 4px; color: {colors["primary_text"]};
        }}
        QLineEdit:hover, QInputDialog QLineEdit:hover {{ border-color: {colors["primary_button"]}; }}

        QPushButton {{
            background-color: {colors["primary_button"]}; color: white; border: none;
            padding: 6px 12px; border-radius: 2px; font-weight: bold;
        }}
        QPushButton:hover {{ background-color: {colors["primary_hover"]}; }}
        QPushButton:pressed {{ background-color: {colors["primary_pressed"]}; }}
        QPushButton:disabled {{ background-color: {colors["secondary_button"]}; color: {colors["disabled_text"]}; }}
        QPushButton#GenerateButton {{ background-color: {colors["generate_button"]}; padding: 10px; }}
        QPushButton#GenerateButton:hover {{ background-color: {colors["generate_hover"]}; }}
        QPushButton#GenerateButton:pressed {{ background-color: {colors["generate_pressed"]}; }}
        QPushButton#RefreshButton {{
            background-color: transparent; border: 1px solid {colors["secondary_button"]};
            padding: 4px; min-width: 28px;
        }}
        QPushButton#RefreshButton:hover {{
             background-color: {colors["content_bg"]}; border-color: {colors["primary_button"]};
        }}
        QPushButton#ChoiceButton {{
            background-color: {colors["content_bg"]}; border: 1px solid {colors["border"]};
            padding: 25px 15px; text-align: center; border-radius: 2px;
        }}
        QPushButton#ChoiceButton:hover {{ background-color: {colors["main_bg"]}; border-color: {colors["primary_button"]}; }}
        
        QComboBox, QInputDialog QComboBox {{
            background-color: {colors["main_bg"]}; border: 1px solid {colors["border"]};
            border-radius: 2px; padding: 4px; color: {colors["primary_text"]};
        }}
        QComboBox:hover, QInputDialog QComboBox:hover {{ border-color: {colors["primary_button"]}; }}
        QComboBox::drop-down, QInputDialog QComboBox::drop-down {{ border: none; width: 18px; }}
        QComboBox QAbstractItemView, QInputDialog QComboBox QAbstractItemView {{
            background-color: {colors["content_bg"]}; border: 1px solid {colors["border"]};
            selection-background-color: {colors["primary_button"]}; color: {colors["primary_text"]}; padding: 4px;
        }}

        QListWidget, QTableWidget, QTextEdit {{
            background-color: {colors["darker_bg"]}; border: 1px solid {colors["border"]};
            border-radius: 2px; color: {colors["primary_text"]};
        }}
        QListWidget::item:selected, QTableWidget::item:selected {{ background-color: {colors["primary_button"]}; color: white; }}
        QHeaderView::section {{
            background-color: {colors["content_bg"]}; color: {colors["primary_text"]}; padding: 4px;
            border: 1px solid {colors["border"]}; font-weight: bold;
        }}
        QCheckBox {{ spacing: 8px; color: {colors["secondary_text"]}; }}
        QCheckBox::indicator {{
            width: 14px; height: 14px; border: 1px solid {colors["border"]};
            border-radius: 2px; background-color: {colors["main_bg"]};
        }}
        QCheckBox::indicator:checked {{ background-color: {colors["generate_button"]}; }}
        QCheckBox::indicator:disabled {{ background-color: {colors["content_bg"]}; }}
        
        QWidget#ScrollContent {{ background: transparent; }}
        QScrollArea {{
            background: transparent;
            border: none;
        }}

        QTabWidget::pane {{ border: 1px solid {colors["border"]}; border-top: none; border-radius: 0 0 2px 2px; }}
        QTabBar::tab {{
            background-color: {colors["main_bg"]}; color: {colors["secondary_text"]};
            border: 1px solid {colors["border"]}; border-bottom: none; padding: 7px 15px;
            border-top-left-radius: 2px; border-top-right-radius: 2px;
        }}
        QTabBar::tab:selected {{ background-color: {colors["content_bg"]}; color: {colors["primary_text"]}; font-weight: bold; }}
        QTabBar::tab:!selected:hover {{ background-color: {colors["content_bg"]}; }}
        QScrollBar:vertical {{ border: none; background: {colors["main_bg"]}; width: 10px; margin: 0; }}
        QScrollBar::handle:vertical {{ background: {colors["border"]}; min-height: 20px; border-radius: 5px; }}
        QScrollBar::handle:vertical:hover {{ background: {colors["secondary_text"]}; }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
        
        QColorDialog {{ background-color: {colors["main_bg"]}; }}
        QColorDialog QGroupBox {{ background-color: {colors["main_bg"]}; color: {colors["primary_text"]}; }}
        QColorDialog QLabel {{ color: {colors["primary_text"]}; }}
        QColorDialog QLineEdit, QColorDialog QSpinBox {{
            background-color: {colors["darker_bg"]}; color: {colors["primary_text"]};
            border: 1px solid {colors["border"]}; border-radius: 2px;
        }}
        QColorDialog QPushButton {{
            background-color: {colors["primary_button"]}; color: white; border-radius: 2px;
        }}
        QColorDialog QPushButton:hover {{ background-color: {colors["primary_hover"]}; }}
        QColorDialog QPushButton:pressed {{ background-color: {colors["primary_pressed"]}; }}
    """