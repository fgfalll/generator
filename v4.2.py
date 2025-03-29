import sys
import os
import importlib.util
import shutil
import yaml
import platform
import psutil
import re
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTextEdit, QButtonGroup, QComboBox,
    QMenuBar, QAction, QDialog, QTabWidget, QFormLayout,
    QSpinBox, QCheckBox, QLineEdit, QListWidget, QListWidgetItem,
    QDialogButtonBox, QMessageBox, QRadioButton, QGroupBox, QFileDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QTime

try:
    from license_manager import LicenseManager
except ImportError:
    print("WARNING: Could not import LicenseManager. License functionality will be unavailable.")
    class LicenseManager(QMainWindow): # Dummy class
        log_signal = pyqtSignal(str, str)
        def __init__(self, parent=None):
            super().__init__(parent)
            self.setWindowTitle("License Manager (Unavailable)")
            layout = QVBoxLayout()
            label = QLabel("Error: License Manager module not found.")
            layout.addWidget(label)
            widget = QWidget()
            widget.setLayout(layout)
            self.setCentralWidget(widget)
            self.resize(400, 100)

# --- Configuration File Path ---
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".DesktopOrganizer")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.yaml")
os.makedirs(CONFIG_DIR, exist_ok=True)

# --- Default Settings ---
DEFAULT_SETTINGS = {
    'application': {
        'autostart_timer_enabled': True,
    },
    'timer': {
        'override_default_enabled': False,
        'default_minutes': 3,
    },
    'drives': {
        'main_drive_policy': 'D',
    },
    'file_manager': {
        'max_file_size_mb': 100,
        'allowed_extensions': ['.lnk'],
        'allowed_filenames': []
    }
}

# --- Module Loading Configuration ---
MODULE_DIR_NAME = "modules" # Subdirectory name for optional modules
EXPECTED_MODULES = {
    "license_manager": {
        "filename": "license_manager.py",
        "class_name": "LicenseManager", # Expected main class in the module
        "menu_text": "&Запустити модуль",
        "menu_object_name": "manageLicenseAction" # To find the action later
    },
    "program_install": {
        "filename": "program_install.py",
        "class_name": "ProgramInstallerUI", # Assuming this will be the class name
        "menu_text": "Запустити модуль",
        "menu_object_name": "installProgramAction"
    },
    # Add more modules here if needed
}

# --- Settings Dialog ---
class SettingsDialog(QDialog):
    settings_applied = pyqtSignal(dict)

    def __init__(self, current_settings, parent=None):
        super().__init__(parent)
        self.current_settings = current_settings.copy()
        self.setWindowTitle("Налаштування")
        self.setMinimumWidth(500)

        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self.create_general_tab()
        self.create_file_manager_tab()

        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel | QDialogButtonBox.Apply)
        self.button_box.button(QDialogButtonBox.Apply).clicked.connect(self.apply_changes)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

        self.load_settings_to_ui()

    def create_general_tab(self):
        tab_general = QWidget()
        layout = QVBoxLayout(tab_general)

        app_group = QGroupBox("Поведінка таймера")
        app_layout = QFormLayout(app_group)
        self.chk_enable_autostart = QCheckBox("Автоматичний старт таймера при запуску")
        app_layout.addRow(self.chk_enable_autostart)

        timer_group = QGroupBox("Налаштування таймера")
        timer_layout = QFormLayout(timer_group)
        self.chk_override_timer = QCheckBox("Встановити користувацьке значення")
        self.spin_default_timer = QSpinBox()
        self.spin_default_timer.setRange(1, 60)
        self.spin_default_timer.setSuffix(" минут")
        self.chk_override_timer.toggled.connect(self.spin_default_timer.setEnabled)
        timer_layout.addRow(self.chk_override_timer)
        timer_layout.addRow("Тривалість за замовчуванням:", self.spin_default_timer)

        drive_group = QGroupBox("Налаштування дисків")
        drive_layout = QVBoxLayout(drive_group)
        drive_layout.addWidget(QLabel("Резервний диск завжди C:"))

        self.rb_drive_d = QRadioButton("Встановити основний диск D:")
        self.rb_drive_auto = QRadioButton("Автоматичне визначення наступного доступного диска")
        drive_layout.addWidget(self.rb_drive_d)
        drive_layout.addWidget(self.rb_drive_auto)

        layout.addWidget(app_group)
        layout.addWidget(timer_group)
        layout.addWidget(drive_group)
        layout.addStretch()
        self.tabs.addTab(tab_general, "Загальні")

    def create_file_manager_tab(self):
        tab_fm = QWidget()
        layout = QFormLayout(tab_fm)

        self.spin_max_size = QSpinBox()
        self.spin_max_size.setRange(1, 10240)
        self.spin_max_size.setSuffix(" MB")
        layout.addRow("Максимальний розмір файлу:", self.spin_max_size)

        ext_layout = QHBoxLayout()
        self.list_extensions = QListWidget()
        self.list_extensions.setFixedHeight(80)
        ext_controls_layout = QVBoxLayout()
        self.edit_add_ext = QLineEdit()
        self.edit_add_ext.setPlaceholderText(".example")
        btn_add_ext = QPushButton("Додати")
        btn_rem_ext = QPushButton("Видалити вибране")
        btn_add_ext.clicked.connect(self.add_extension)
        btn_rem_ext.clicked.connect(self.remove_extension)
        ext_controls_layout.addWidget(self.edit_add_ext)
        ext_controls_layout.addWidget(btn_add_ext)
        ext_controls_layout.addWidget(btn_rem_ext)
        ext_layout.addWidget(self.list_extensions)
        ext_layout.addLayout(ext_controls_layout)
        layout.addRow("Пропустити розширення:", ext_layout)

        name_layout = QHBoxLayout()
        self.list_filenames = QListWidget()
        self.list_filenames.setFixedHeight(80)
        name_controls_layout = QVBoxLayout()
        self.edit_add_name = QLineEdit()
        self.edit_add_name.setPlaceholderText("імя_файлу_без_розширення")
        btn_add_name = QPushButton("Додати")
        btn_rem_name = QPushButton("Видалити вибране")
        btn_add_name.clicked.connect(self.add_filename)
        btn_rem_name.clicked.connect(self.remove_filename)
        name_controls_layout.addWidget(self.edit_add_name)
        name_controls_layout.addWidget(btn_add_name)
        name_controls_layout.addWidget(btn_rem_name)
        name_layout.addWidget(self.list_filenames)
        name_layout.addLayout(name_controls_layout)
        layout.addRow("Пропустити файли:", name_layout)

        self.tabs.addTab(tab_fm, "Фільтри для файлів")

    def add_extension(self):
        ext = self.edit_add_ext.text().strip().lower()
        if ext.startswith('.') and len(ext) > 1:
            if not self.list_extensions.findItems(ext, Qt.MatchExactly):
                self.list_extensions.addItem(ext)
                self.edit_add_ext.clear()
        else:
            QMessageBox.warning(self, "Неправильне розширення", "Розширення повинно починатися з “.” і не бути порожнім.")

    def remove_extension(self):
        for item in self.list_extensions.selectedItems():
            self.list_extensions.takeItem(self.list_extensions.row(item))

    def add_filename(self):
        name = self.edit_add_name.text().strip()
        if name and not any(c in name for c in '/\\:*?"<>|'):
             if not self.list_filenames.findItems(name, Qt.MatchExactly):
                self.list_filenames.addItem(name)
                self.edit_add_name.clear()
        else:
             QMessageBox.warning(self, "Неправильне ім'я файлу", "Ім'я файлу не може бути порожнім або містити недопустимі символи.")

    def remove_filename(self):
        for item in self.list_filenames.selectedItems():
            self.list_filenames.takeItem(self.list_filenames.row(item))

    def load_settings_to_ui(self):
        app_cfg = self.current_settings.get('application', DEFAULT_SETTINGS['application'])
        self.chk_enable_autostart.setChecked(app_cfg.get('autostart_timer_enabled', True))

        timer_cfg = self.current_settings.get('timer', DEFAULT_SETTINGS['timer'])
        self.chk_override_timer.setChecked(timer_cfg.get('override_default_enabled', False))
        self.spin_default_timer.setValue(timer_cfg.get('default_minutes', 3))
        self.spin_default_timer.setEnabled(self.chk_override_timer.isChecked())

        drive_cfg = self.current_settings.get('drives', DEFAULT_SETTINGS['drives'])
        policy = drive_cfg.get('main_drive_policy', 'D')
        if policy == 'auto':
            self.rb_drive_auto.setChecked(True)
        else:
            self.rb_drive_d.setChecked(True)

        fm_cfg = self.current_settings.get('file_manager', DEFAULT_SETTINGS['file_manager'])
        self.spin_max_size.setValue(fm_cfg.get('max_file_size_mb', 100))

        self.list_extensions.clear()
        self.list_extensions.addItems(fm_cfg.get('allowed_extensions', []))

        self.list_filenames.clear()
        self.list_filenames.addItems(fm_cfg.get('allowed_filenames', []))

    def get_settings_from_ui(self):
        updated_settings = {
            'application': {
                'autostart_timer_enabled': self.chk_enable_autostart.isChecked()
            },
            'timer': {
                'override_default_enabled': self.chk_override_timer.isChecked(),
                'default_minutes': self.spin_default_timer.value()
            },
            'drives': {
                'main_drive_policy': 'auto' if self.rb_drive_auto.isChecked() else 'D'
            },
            'file_manager': {
                'max_file_size_mb': self.spin_max_size.value(),
                'allowed_extensions': sorted([self.list_extensions.item(i).text() for i in range(self.list_extensions.count())]),
                'allowed_filenames': sorted([self.list_filenames.item(i).text() for i in range(self.list_filenames.count())])
            }
        }
        return updated_settings

    def apply_changes(self):
        new_settings = self.get_settings_from_ui()
        self.current_settings = new_settings
        self.settings_applied.emit(new_settings)

    def accept(self):
        self.apply_changes()
        super().accept()


# --- File Mover Thread ---
class FileMover(QThread):
    update_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(int, int, str)

    def __init__(self, target_drive, fallback_drive, settings):
        super().__init__()
        self.target_drive = target_drive
        self.fallback_drive = fallback_drive
        self.settings = settings
        self.base_folder_name = "Робочі столи"

    def run(self):
        try:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            now = datetime.now()

            fm_settings = self.settings.get('file_manager', DEFAULT_SETTINGS['file_manager'])
            skip_extensions = {ext.lower() for ext in fm_settings.get('allowed_extensions', [])}
            skip_filenames = {name for name in fm_settings.get('allowed_filenames', [])}
            max_size_bytes = fm_settings.get('max_file_size_mb', 100) * 1024 * 1024

            target_base_path = os.path.join(f"{self.target_drive}:\\", self.base_folder_name)
            fallback_base_path = os.path.join(f"{self.fallback_drive}:\\", self.base_folder_name)

            effective_base_path = ""
            if self.check_drive_exists(self.target_drive):
                 effective_base_path = target_base_path
            elif self.check_drive_exists(self.fallback_drive):
                 self.update_signal.emit(f"⚠️ Диск {self.target_drive}: недоступний. Використовуємо {self.fallback_drive}:")
                 effective_base_path = fallback_base_path
            else:
                self.update_signal.emit(f"❌ Критична помилка: Цільовий диск {self.target_drive}: та резервний {self.fallback_drive}: недоступні.")
                self.finished_signal.emit(0, 0, "Error: No accessible drives")
                return

            year = now.strftime("%Y")
            timestamp = now.strftime("%d-%m-%Y %H-%M")
            dest_path = os.path.join(effective_base_path, f"Робочий стіл {year}", f"Робочий стіл {timestamp}")

            os.makedirs(dest_path, exist_ok=True)
            self.update_signal.emit(f"📁 Цільова папка: {dest_path}")

            success = errors = 0
            if not os.path.isdir(desktop):
                self.update_signal.emit(f"❌ Помилка: Папка робочого столу не знайдена за шляхом {desktop}")
                self.finished_signal.emit(0, 0, dest_path)
                return

            items_to_move = os.listdir(desktop)
            if not items_to_move:
                 self.update_signal.emit("ℹ️ Робочий стіл порожній. Немає чого переміщувати.")

            for item in items_to_move:
                src = os.path.join(desktop, item)
                item_name_no_ext, item_ext = os.path.splitext(item)
                item_ext_lower = item_ext.lower()

                if item_ext_lower in skip_extensions:
                    self.update_signal.emit(f"⏭️ Пропущено за розширенням: {item}")
                    continue

                if item_name_no_ext in skip_filenames:
                    self.update_signal.emit(f"⏭️ Пропущено за ім'ям файлу: {item}")
                    continue

                if os.path.isfile(src):
                    try:
                        file_size = os.path.getsize(src)
                        if file_size > max_size_bytes:
                            self.update_signal.emit(f"⏭️ Пропущено за розміром ({file_size / (1024*1024):.1f}MB): {item}")
                            continue
                    except OSError as e:
                         self.update_signal.emit(f"⚠️ Не вдалося отримати розмір {item}: {e}")
                         continue

                try:
                    final_dest = os.path.join(dest_path, item)
                    shutil.move(src, final_dest)
                    success += 1
                    self.update_signal.emit(f"✅ Переміщено: {item}")
                except Exception as e:
                    errors += 1
                    self.update_signal.emit(f"❌ Помилка переміщення '{item}': {str(e)}")

            self.finished_signal.emit(success, errors, dest_path)

        except Exception as e:
            self.update_signal.emit(f"❌ Критична помилка потоку: {str(e)}")
            self.finished_signal.emit(0, 0, "Error in thread")

    def check_drive_exists(self, drive_letter):
        drive = f"{drive_letter}:\\"
        return os.path.exists(drive)

# --- Main Window ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = self.load_settings()
        self.mover_thread = None
        # self.license_manager_window = None # Remove this - we'll handle dynamically
        self.loaded_modules = {}  # Stores loaded module classes/functions
        self.module_windows = {}  # Stores instances of opened module windows
        self.module_actions = {}  # Stores menu actions related to modules

        self.auto_start_timer = QTimer(self)
        self.auto_start_timer.timeout.connect(self.update_timer)
        self.remaining_time = 0
        self.selected_drive = 'C'
        self.d_exists = False
        self.e_exists = False

        self.initUI()  # Create UI elements first
        self.load_optional_modules()  # Attempt to load modules
        self.update_ui_for_modules()  # Enable/disable menus based on loaded modules

        self.apply_settings_to_ui()  # Apply loaded settings to UI

        QTimer.singleShot(500, self.auto_configure_start)  # Existing delayed config

    def load_settings(self):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                loaded_settings = yaml.safe_load(f)
                if loaded_settings:
                    merged = self._merge_dicts(DEFAULT_SETTINGS.copy(), loaded_settings)
                    return merged
                else:
                    return DEFAULT_SETTINGS.copy()
        except FileNotFoundError:
            print(f"Config file not found at {CONFIG_FILE}. Using defaults.")
            return DEFAULT_SETTINGS.copy()
        except yaml.YAMLError as e:
            print(f"Error parsing config file {CONFIG_FILE}: {e}. Using defaults.")
            return DEFAULT_SETTINGS.copy()
        except Exception as e:
            print(f"Unexpected error loading config {CONFIG_FILE}: {e}. Using defaults.")
            return DEFAULT_SETTINGS.copy()

    def get_module_dir(self):
        """Determines the path to the 'modules' directory relative to the script or executable."""
        if getattr(sys, 'frozen', False):
            # Running as a bundled executable (PyInstaller)
            base_path = os.path.dirname(sys.executable)
        else:
            # Running as a normal Python script
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, MODULE_DIR_NAME)

    def load_optional_modules(self):
        """Scans the module directory and loads any found optional modules."""
        module_dir = self.get_module_dir()
        self.log_message(f"ℹ️ Перевірка наявності додаткових модулів у: {module_dir}")

        if not os.path.isdir(module_dir):
            self.log_message(f"ℹ️ Каталог модулів не знайдено. Пропущено необов'язкові модулі.")
            return

        for key, config in EXPECTED_MODULES.items():
            module_path = os.path.join(module_dir, config["filename"])

            if os.path.exists(module_path):
                self.log_message(f"ℹ️ Знайдено потенційний файл модуля: {config['filename']}")
                try:
                    # Create a unique module name to avoid conflicts
                    module_name = f"dynamic_modules.{key}"

                    # Load the module using importlib
                    spec = importlib.util.spec_from_file_location(module_name, module_path)
                    if spec is None:
                        raise ImportError(f"Не вдалося отримати специфікацію для модуля за адресою {module_path}")

                    module = importlib.util.module_from_spec(spec)
                    if module is None:
                        raise ImportError(f"Не вдалося створити модуль зі специфікації {module_name}")

                    # Add to sys.modules BEFORE executing, crucial for relative imports within the module
                    sys.modules[module_name] = module
                    spec.loader.exec_module(module)

                    # --- Integration Point ---
                    # Find the expected class within the loaded module
                    if hasattr(module, config["class_name"]):
                        self.loaded_modules[key] = getattr(module, config["class_name"])
                        self.log_message(f"✅ Успішно завантажено модуль '{key}'.")
                    else:
                        self.log_message(
                            f"⚠️ Модуль '{key}' завантажений, але необхідний клас '{config['class_name']}' не знайдено.")
                        # Optional: Clean up sys.modules if class not found?
                        # del sys.modules[module_name]

                except Exception as e:
                    self.log_message(f"❌ Помилка завантаження модуля '{config['filename']}': {e}")
                    # Ensure partially loaded module is removed from sys.modules
                    if module_name in sys.modules:
                        del sys.modules[module_name]
            else:
                self.log_message(f"ℹ️ Додатковий модуль '{config['filename']}' не знайдено.")

    def _merge_dicts(self, base, updates):
        for key, value in updates.items():
            if isinstance(value, dict) and key in base and isinstance(base[key], dict):
                self._merge_dicts(base[key], value)
            else:
                base[key] = value
        return base

    def save_settings(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                yaml.dump(self.settings, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
        except Exception as e:
            print(f"Помилка збереження налаштувань до {CONFIG_FILE}: {e}")
            QMessageBox.critical(self, "Помилка збереження", f"Could not save settings to {CONFIG_FILE}:\n{e}")

    def find_next_available_drive(self):
        available_drives = []
        try:
            partitions = psutil.disk_partitions(all=False)
            for p in partitions:
                if platform.system() == "Windows" and re.match("^[A-Z]:\\\\?$", p.mountpoint) and p.mountpoint[0] != 'C':
                    if p.fstype and 'cdrom' not in p.opts.lower():
                         if 'removable' not in p.opts.lower():
                             if os.path.exists(p.mountpoint):
                                  available_drives.append(p.mountpoint[0])
            available_drives.sort()
            return available_drives[0] if available_drives else None
        except Exception as e:
            self.log_message(f"⚠️ Error detecting drives: {e}. Falling back.")
            return None

    def auto_configure_start(self):
        policy = self.settings.get('drives', {}).get('main_drive_policy', 'D')
        initial_drive = None

        self.check_drive_availability()

        if policy == 'D' and self.d_exists:
            initial_drive = 'D'
        elif policy == 'auto':
            detected_drive = self.find_next_available_drive()
            if detected_drive:
                initial_drive = detected_drive
            elif self.d_exists:
                self.log_message("ℹ️ Auto-detect failed or no suitable drive, falling back to D:")
                initial_drive = 'D'
        elif policy == 'D' and not self.d_exists and self.e_exists:
             self.log_message(f"ℹ️ Main drive policy 'D' specified, but D: not found. Falling back to E:")
             initial_drive = 'E'
        elif self.e_exists and not initial_drive:
             self.log_message(f"ℹ️ Main drive policy '{policy}' failed or not applicable, falling back to E:")
             initial_drive = 'E'

        if initial_drive:
            self.selected_drive = initial_drive
        else:
            self.selected_drive = 'C'
            if policy != 'C':
                self.log_message("⚠️ No suitable main drive found (D:, E:, or auto-detected). Using C:")

        self.log_message(f"⚙️ Initial main drive set to: {self.selected_drive}:")
        self.update_drive_buttons_visuals()

        self.apply_settings_to_ui()

        app_settings = self.settings.get('application', DEFAULT_SETTINGS['application'])
        if app_settings.get('autostart_timer_enabled', True):
            self.start_auto_timer()
        else:
             self.log_message("ℹ️ Autostart timer disabled in settings.")
             self.stop_auto_timer(log_disabled=True)


    def apply_settings_to_ui(self):
        timer_cfg = self.settings.get('timer', DEFAULT_SETTINGS['timer'])
        minutes_to_set = DEFAULT_SETTINGS['timer']['default_minutes']

        if timer_cfg.get('override_default_enabled', False):
            minutes_to_set = timer_cfg.get('default_minutes', minutes_to_set)

        index_to_set = -1
        for i in range(self.time_combo.count()):
            try:
                if int(self.time_combo.itemText(i).split()[0]) == minutes_to_set:
                    index_to_set = i
                    break
            except: pass
        if index_to_set != -1:
             self.time_combo.blockSignals(True)
             self.time_combo.setCurrentIndex(index_to_set)
             self.time_combo.blockSignals(False)
        else:
             self.time_combo.blockSignals(True)
             self.time_combo.setCurrentIndex(1)
             self.time_combo.blockSignals(False)

        if not self.auto_start_timer.isActive():
            self.update_timer_label_when_stopped()

        self.update_drive_buttons_visuals()

    def update_timer_label_when_stopped(self):
        minutes_text = self.time_combo.currentText()
        try:
            minutes_val = int(minutes_text.split()[0])
            self.remaining_time = minutes_val * 60
            self.timer_label.setText(f"Автозапуск вимкнено ({self.format_time()})")
        except:
            self.timer_label.setText("Автозапуск вимкнено (Помилка)")


    def initUI(self):
        self.setWindowTitle("Авто-організатор робочого столу v4.2")
        self.setGeometry(300, 300, 650, 500)

        menubar = self.menuBar()

        # --- File Menu ---
        file_menu = menubar.addMenu('&Файл')
        settings_action = QAction('&Налаштування...', self)
        settings_action.triggered.connect(self.open_settings_dialog)
        file_menu.addAction(settings_action)
        # --- Add Import Module Action ---
        import_module_action = QAction('&Імпортувати додатковий модуль', self)
        import_module_action.triggered.connect(self.import_modules_to_standard_dir)
        file_menu.addAction(import_module_action)
        file_menu.addSeparator()
        exit_action = QAction('&Вихід', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # --- Install Programs Menu ---
        install_menu = menubar.addMenu('&Встановлення програм')
        install_key = "program_install"
        if install_key in EXPECTED_MODULES:
            config = EXPECTED_MODULES[install_key]
            install_program_action = QAction(config["menu_text"], self)
            install_program_action.setObjectName(config["menu_object_name"])
            # Connect to a generic opener function, passing the module key
            install_program_action.triggered.connect(
                lambda checked=False, key=install_key: self.open_module_window(key))
            install_program_action.setEnabled(False)  # Initially disabled
            install_menu.addAction(install_program_action)
            self.module_actions[install_key] = install_program_action
        else:
            # Optional: Add a placeholder if the config is missing entirely
            placeholder_action = QAction("Install (Not Configured)", self)
            placeholder_action.setEnabled(False)
            install_menu.addAction(placeholder_action)

        # --- License Menu ---
        license_menu = menubar.addMenu('&Встановлення ліцензій')
        license_key = "license_manager"
        if license_key in EXPECTED_MODULES:
            config = EXPECTED_MODULES[license_key]
            manage_license_action = QAction(config["menu_text"], self)
            manage_license_action.setObjectName(config["menu_object_name"])
            # Connect to a generic opener function, passing the module key
            manage_license_action.triggered.connect(
                lambda checked=False, key=license_key: self.open_module_window(key))
            manage_license_action.setEnabled(False)  # Initially disabled
            license_menu.addAction(manage_license_action)
            self.module_actions[license_key] = manage_license_action
        else:
            placeholder_action = QAction("License (Not Configured)", self)
            placeholder_action.setEnabled(False)
            license_menu.addAction(placeholder_action)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        self.timer_label = QLabel("Завантаження...")
        main_layout.addWidget(self.timer_label)

        control_layout = QHBoxLayout()
        self.time_combo = QComboBox()
        self.time_combo.addItems(["1 хвилина", "3 хвилини", "5 хвилин", "10 хвилин"])
        self.time_combo.currentIndexChanged.connect(self.time_selection_changed)
        control_layout.addWidget(self.time_combo)
        self.start_now_btn = QPushButton("🚀 Старт зараз")
        self.start_now_btn.clicked.connect(self.start_now)
        control_layout.addWidget(self.start_now_btn)
        self.timer_control_btn = QPushButton("⏱️ Стоп таймер")
        self.timer_control_btn.clicked.connect(self.toggle_timer)
        control_layout.addWidget(self.timer_control_btn)
        main_layout.addLayout(control_layout)

        drive_group = QGroupBox("Вибір основного диска")
        drive_layout = QHBoxLayout(drive_group)
        self.btn_group = QButtonGroup(self)
        self.btn_drive_d = QPushButton("Диск D:")
        self.btn_drive_e = QPushButton("Диск E:")
        self.btn_group.addButton(self.btn_drive_d, ord('D'))
        self.btn_group.addButton(self.btn_drive_e, ord('E'))
        drive_layout.addWidget(self.btn_drive_d)
        drive_layout.addWidget(self.btn_drive_e)
        self.btn_group.buttonClicked[int].connect(self.set_selected_drive)
        main_layout.addWidget(drive_group)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        main_layout.addWidget(self.log)

    def import_modules_to_standard_dir(self):
        """Opens a dialog to select .py files and copies them to the standard module directory."""
        source_files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Module Files to Import",
            os.path.expanduser("~"),  # Start in user's home directory or last path
            "Python files (*.py);;All files (*.*)"
        )

        if not source_files:
            self.log_message("ℹ️ Module import cancelled by user.")
            return

        target_dir = self.get_module_dir()  # Get ./modules path
        try:
            os.makedirs(target_dir, exist_ok=True)  # Ensure the directory exists
        except OSError as e:
            self.log_message(f"❌ Critical Error: Could not create module directory '{target_dir}': {e}")
            QMessageBox.critical(self, "Import Error",
                                 f"Failed to create the target module directory:\n{target_dir}\n\n{e}")
            return

        copied_count = 0
        skipped_count = 0
        error_count = 0
        modules_changed = False

        for src_path in source_files:
            filename = os.path.basename(src_path)
            dest_path = os.path.join(target_dir, filename)

            # Check for overwrite
            if os.path.exists(dest_path):
                reply = QMessageBox.question(
                    self,
                    "Confirm Overwrite",
                    f"The module '{filename}' already exists in the standard folder.\nDo you want to overwrite it?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No  # Default to No
                )
                if reply == QMessageBox.No:
                    self.log_message(f"⏭️ Skipped overwrite for: {filename}")
                    skipped_count += 1
                    continue

            # Attempt to copy
            try:
                shutil.copy2(src_path, dest_path)  # copy2 preserves metadata
                self.log_message(f"✅ Imported: {filename}")
                copied_count += 1
                modules_changed = True
            except Exception as e:
                self.log_message(f"❌ Error importing '{filename}': {e}")
                error_count += 1

        summary = f"🏁 Import process finished. Copied: {copied_count}, Skipped: {skipped_count}, Errors: {error_count}."
        self.log_message(summary)
        QMessageBox.information(self, "Import Complete", summary)

        # Reload modules if any file was successfully copied
        if modules_changed:
            self.log_message("🔄 Reloading modules after import...")
            self.reload_modules_and_update_ui()

    def update_ui_for_modules(self):
        """Enables or disables menu actions based on successfully loaded modules."""
        for key, action in self.module_actions.items():
            is_loaded = key in self.loaded_modules
            action.setEnabled(is_loaded)
            #self.log_message(f"ℹ️ UI Action for '{key}' {'Enabled' if is_loaded else 'Disabled'}.")

    def open_module_window(self, module_key):
        """Opens the window associated with the given module key, if loaded."""
        if module_key in self.loaded_modules:
            try:
                ModuleClass = self.loaded_modules[module_key]

                # Check if an instance already exists and is visible
                if module_key in self.module_windows and self.module_windows[module_key].isVisible():
                    self.module_windows[module_key].activateWindow()
                    self.module_windows[module_key].raise_()
                else:
                    # --- Instantiation Point ---
                    # Pass 'self' (the main window) as parent, useful if the module needs
                    # to interact with the main app (e.g., logging, accessing settings).
                    # Ensure the module's __init__ accepts a parent argument.
                    module_window_instance = ModuleClass(parent=self)
                    self.module_windows[module_key] = module_window_instance  # Store the instance

                    # Optional: Connect signals from the module window back to the main window if needed
                    # if hasattr(module_window_instance, 'log_signal'):
                    #     module_window_instance.log_signal.connect(self.handle_module_log)

                    module_window_instance.show()

            except Exception as e:
                self.log_message(f"❌ Error instantiating or showing window for module '{module_key}': {e}")
                QMessageBox.critical(self, "Module Error", f"Could not launch the '{module_key}' module window.\n\n{e}")
        else:
            self.log_message(f"⚠️ Attempted to open module '{module_key}', but it is not loaded.")
            QMessageBox.warning(self, "Module Not Available",
                                f"The required module '{module_key}' was not found or failed to load.")

    def reload_modules_and_update_ui(self):
        """Clears, reloads modules, and updates the UI accordingly."""
        # 1. Close existing module windows (Optional but safer)
        keys_to_close = list(self.module_windows.keys())
        for key in keys_to_close:
            window = self.module_windows.pop(key, None)
            if window and window.isVisible():
                self.log_message(f"Attempting to close window for module '{key}'...")
                try:
                    # Disconnect signals maybe? Depends on module design
                    window.close()
                    # window.deleteLater() # More aggressive cleanup?
                except Exception as e:
                    self.log_message(f"Error closing window for '{key}': {e}")
            elif window:
                # window.deleteLater() # Cleanup non-visible too?
                pass

        QApplication.processEvents()  # Allow windows to close

        # 2. Clear internal references
        # Make copies of keys before iterating if deleting during iteration
        loaded_keys = list(self.loaded_modules.keys())
        self.loaded_modules.clear()
        self.log_message(f"Cleared internal module references: {loaded_keys}")

        # 3. Attempt to clean up sys.modules (Use the specific names we created)
        # This might STILL not be enough for C extensions!
        external_prefix = "dynamic_modules.external."
        default_prefix = "dynamic_modules.default."
        modules_to_delete = []
        for mod_name in sys.modules:
            if mod_name.startswith(external_prefix) or mod_name.startswith(default_prefix):
                modules_to_delete.append(mod_name)

        for mod_name in modules_to_delete:
            self.log_message(f"Removing '{mod_name}' from sys.modules...")
            try:
                del sys.modules[mod_name]
            except KeyError:
                pass  # Already gone

        self.update_ui_for_modules()
        self.log_message("--- Module Reload Finished ---")


    def handle_settings_applied(self, new_settings):
        self.settings = new_settings
        self.save_settings()
        self.apply_settings_to_ui()
        self.log_message("⚙️ Налаштування застосовано.")
        if not self.auto_start_timer.isActive():
            app_settings = self.settings.get('application', DEFAULT_SETTINGS['application'])
            if not app_settings.get('autostart_timer_enabled', True):
                 self.stop_auto_timer(log_disabled=True)
            else:
                 self.update_timer_label_when_stopped()


    def open_settings_dialog(self):
        dialog = SettingsDialog(self.settings, self)
        dialog.settings_applied.connect(self.handle_settings_applied)
        dialog.exec_()


    def show_install_placeholder(self):
        QMessageBox.information(self, "Install Programs", "This feature is not yet implemented.")


    def open_license_manager(self):
        if self.license_manager_window and self.license_manager_window.isVisible():
            self.license_manager_window.activateWindow()
            self.license_manager_window.raise_()
        else:
            try:
                # Re-check if the import worked initially
                from license_manager import LicenseManager as ActualLicenseManager
                self.license_manager_window = ActualLicenseManager(self)
                # Optional: Connect log signal
                # self.license_manager_window.log_signal.connect(self.handle_license_log)
                self.license_manager_window.show()
            except ImportError:
                 QMessageBox.critical(self, "Error", "License Manager module could not be loaded.")
            except Exception as e:
                 QMessageBox.critical(self, "Error", f"Failed to open License Manager:\n{e}")

    # Optional: Slot to handle logs from License Manager
    # def handle_license_log(self, level, message):
    #     prefix = "[LicenseMgr]"
    #     self.log_message(f"{prefix} [{level.upper()}] {message}")


    def time_selection_changed(self):
        if not self.auto_start_timer.isActive():
            self.update_timer_label_when_stopped()


    def check_drive_availability(self):
        self.d_exists = os.path.exists("D:\\")
        self.e_exists = os.path.exists("E:\\")

        self.btn_drive_d.setEnabled(self.d_exists)
        self.btn_drive_e.setEnabled(self.e_exists)

        self.update_drive_buttons_visuals()


    def update_drive_buttons_visuals(self):
        self.btn_drive_d.setText(f"Диск D: {'🟢' if self.d_exists else '🔴'}")
        self.btn_drive_e.setText(f"Диск E: {'🟢' if self.e_exists else '🔴'}")

        for button in self.btn_group.buttons():
            drive_letter = chr(self.btn_group.id(button))
            is_selected = (drive_letter == self.selected_drive)
            if is_selected:
                button.setStyleSheet("font-weight: bold; border: 2px solid blue;")
            else:
                # Reset style, consider inheriting default look or using a less specific style
                button.setStyleSheet("") # Reset to default or use "font-weight: normal; border: none;"


    def set_selected_drive(self, drive_id):
        drive_letter = chr(drive_id)
        if (drive_letter == 'D' and self.d_exists) or \
           (drive_letter == 'E' and self.e_exists):
            if drive_letter != self.selected_drive:
                self.selected_drive = drive_letter
                self.log_message(f"Обрано основний диск: {self.selected_drive}:")
                self.update_drive_buttons_visuals()
                self.stop_auto_timer()
        else:
             self.log_message(f"⚠️ Диск {drive_letter}: недоступний.")
             self.update_drive_buttons_visuals()


    def toggle_timer(self):
        if self.auto_start_timer.isActive():
            self.stop_auto_timer()
        else:
            # Check if autostart is globally disabled by settings before starting
            app_settings = self.settings.get('application', DEFAULT_SETTINGS['application'])
            if not app_settings.get('autostart_timer_enabled', True):
                 self.log_message("ℹ️ Timer cannot be started manually when autostart is disabled in settings.")
                 # Optionally show a QMessageBox here too
                 return
            self.start_auto_timer()


    def start_now(self):
        if self.mover_thread and self.mover_thread.isRunning():
             QMessageBox.warning(self, "Зайнято", "Процес переміщення вже запущено.")
             return
        self.stop_auto_timer()
        self.start_process()


    def start_auto_timer(self):
        if self.mover_thread and self.mover_thread.isRunning():
             self.log_message("ℹ️ Неможливо запустити таймер під час переміщення.")
             return

        # Explicitly check the setting again before starting
        app_settings = self.settings.get('application', DEFAULT_SETTINGS['application'])
        if not app_settings.get('autostart_timer_enabled', True):
             self.log_message("ℹ️ Timer start prevented by application settings (Autostart disabled).")
             self.stop_auto_timer(log_disabled=True) # Ensure UI reflects disabled state
             return

        minutes_text = self.time_combo.currentText()
        try:
             minutes = int(minutes_text.split()[0])
        except:
             minutes = 3
             self.log_message("⚠️ Не вдалося розпізнати час таймера, встановлено 3 хв.")

        self.remaining_time = minutes * 60
        if self.remaining_time <= 0:
             self.remaining_time = 180
        self.timer_label.setText(f"До автоматичного старту: {self.format_time()}")
        self.timer_control_btn.setText("⏱️ Стоп таймер")
        self.time_combo.setEnabled(False)
        self.btn_drive_d.setEnabled(False)
        self.btn_drive_e.setEnabled(False)
        self.auto_start_timer.start(1000)


    def stop_auto_timer(self, log_disabled=False):
        self.auto_start_timer.stop()
        if log_disabled:
             self.timer_label.setText("Автозапуск вимкнено (Налаштування)")
        else:
             self.update_timer_label_when_stopped()

        self.timer_control_btn.setText("▶️ Старт таймер")
        self.time_combo.setEnabled(True)
        self.check_drive_availability()


    def update_timer(self):
        self.remaining_time -= 1
        if self.remaining_time <= 0:
            self.auto_start_timer.stop()
            self.start_process()
            return
        self.timer_label.setText(f"До автоматичного старту: {self.format_time()}")


    def format_time(self):
        mins, secs = divmod(self.remaining_time, 60)
        return f"{mins:02}:{secs:02}"


    def start_process(self):
        if not self.selected_drive:
            self.log_message("❌ Помилка: Не вдалося визначити цільовий диск.")
            self.check_drive_availability()
            return

        if self.mover_thread and self.mover_thread.isRunning():
            self.log_message("⚠️ Процес вже запущено.")
            return

        self.log_message(f"\n🚀 Початок переміщення на диск {self.selected_drive}:...")
        self.start_now_btn.setEnabled(False)
        self.timer_control_btn.setEnabled(False)
        self.time_combo.setEnabled(False)
        self.btn_drive_d.setEnabled(False)
        self.btn_drive_e.setEnabled(False)

        self.mover_thread = FileMover(target_drive=self.selected_drive, fallback_drive='C', settings=self.settings.copy())
        self.mover_thread.update_signal.connect(self.log_message)
        self.mover_thread.finished_signal.connect(self.process_finished)
        self.mover_thread.start()


    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log.append(f"[{timestamp}] {message}")
        QApplication.processEvents()


    def process_finished(self, success, errors, path):
        self.log_message("\n🏁 Результат:")
        self.log_message(f"✅ Успішно: {success}")
        if errors > 0:
            self.log_message(f"❌ Помилок: {errors}")
        if not path.startswith("Error"):
             self.log_message(f"📁 Збережено до: {path}")
        else:
             self.log_message(f"❌ {path}")

        self.start_now_btn.setEnabled(True)
        self.timer_control_btn.setEnabled(True)
        self.time_combo.setEnabled(True)
        self.check_drive_availability()

        app_settings = self.settings.get('application', DEFAULT_SETTINGS['application'])
        if app_settings.get('autostart_timer_enabled', True):
             self.log_message("⚙️ Перезапуск таймера...")
             QTimer.singleShot(1000, self.start_auto_timer)
        else:
             self.stop_auto_timer(log_disabled=True)


    def closeEvent(self, event):
        # Optional: Add confirmation before closing if needed
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())