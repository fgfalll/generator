import sys
import os
import re
import platform
import time # For timestamps
import shutil # For potential backup

try:
    import winreg
except ImportError:
    # Handle systems where winreg is not available (non-Windows)
    winreg = None
    print("INFO: winreg module not found (non-Windows system). Windows installation disabled.")

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QFileDialog, QMessageBox,
    QToolBar, QSpacerItem, QSizePolicy, QCheckBox, QStatusBar,
    QLabel, QPlainTextEdit, QAction # Added/modified imports
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QRegExp, QSize, QRect, QMetaObject
from PyQt5.QtGui import (
    QTextCharFormat, QColor, QSyntaxHighlighter, QFont, QIcon, QPainter,
    QTextFormat # Added/modified imports
)

# --- Line Number Widget ---
class LineNumberArea(QWidget):
    def __init__(self, editor):
        super().__init__(editor)
        self.codeEditor = editor

    def sizeHint(self):
        return QSize(self.codeEditor.lineNumberAreaWidth(), 0)

    def paintEvent(self, event):
        self.codeEditor.lineNumberAreaPaintEvent(event)


class CodeEditor(QPlainTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.lineNumberArea = LineNumberArea(self)

        self.blockCountChanged.connect(self.updateLineNumberAreaWidth)
        self.updateRequest.connect(self.updateLineNumberArea)
        self.cursorPositionChanged.connect(self.highlightCurrentLine)

        self.updateLineNumberAreaWidth(0)
        self.highlightCurrentLine()
        self.setLineWrapMode(QPlainTextEdit.NoWrap)

    def lineNumberAreaWidth(self):
        digits = 1
        max_count = max(1, self.blockCount())
        while max_count >= 10:
            max_count //= 10
            digits += 1
        space = 3 + self.fontMetrics().horizontalAdvance('9') * digits + 3
        return space

    def updateLineNumberAreaWidth(self, _):
        self.setViewportMargins(self.lineNumberAreaWidth(), 0, 0, 0)

    def updateLineNumberArea(self, rect, dy):
        if dy:
            self.lineNumberArea.scroll(0, dy)
        else:
            self.lineNumberArea.update(0, rect.y(), self.lineNumberArea.width(), rect.height())

        if rect.contains(self.viewport().rect()):
            self.updateLineNumberAreaWidth(0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        cr = self.contentsRect()
        self.lineNumberArea.setGeometry(QRect(cr.left(), cr.top(), self.lineNumberAreaWidth(), cr.height()))

    def lineNumberAreaPaintEvent(self, event):
        painter = QPainter(self.lineNumberArea)
        painter.fillRect(event.rect(), QColor("#e0e0e0"))

        block = self.firstVisibleBlock()
        blockNumber = block.blockNumber()
        top = int(self.blockBoundingGeometry(block).translated(self.contentOffset()).top())
        bottom = top + int(self.blockBoundingRect(block).height())

        height = self.fontMetrics().height()
        while block.isValid() and top <= event.rect().bottom():
            if block.isVisible() and bottom >= event.rect().top():
                number = str(blockNumber + 1)
                painter.setPen(Qt.black)
                painter.drawText(0, top, self.lineNumberArea.width() - 3, height,
                                 Qt.AlignRight, number)

            block = block.next()
            top = bottom
            bottom = top + int(self.blockBoundingRect(block).height())
            blockNumber += 1

    def highlightCurrentLine(self):
        extraSelections = []
        if not self.isReadOnly():
            selection = QTextEdit.ExtraSelection()
            lineColor = QColor(Qt.yellow).lighter(160)
            selection.format.setBackground(lineColor)
            selection.format.setProperty(QTextFormat.FullWidthSelection, True)
            selection.cursor = self.textCursor()
            selection.cursor.clearSelection()
            extraSelections.append(selection)
        self.setExtraSelections(extraSelections)


# --- Syntax Highlighter ---
class NGITHighlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.rules = []
        self.error_lines = set()

        comment_format = QTextCharFormat()
        comment_format.setForeground(QColor('#008000'))
        comment_format.setFontItalic(True)
        self.rules.append((QRegExp("^#[^\n]*"), comment_format))

        key_format = QTextCharFormat()
        key_format.setForeground(QColor('#000080'))
        key_format.setFontWeight(QFont.Bold)
        self.rules.append((QRegExp("^\\s*([A-Z_][A-Z0-9_]*)\\s*(?==)"), key_format))

        value_format = QTextCharFormat()
        value_format.setForeground(QColor('#A31515'))
        self.rules.append((QRegExp("(?<==\\s*).*(?<!\\s)(?=\\s*(?:#|$))"), value_format))

        equals_format = QTextCharFormat()
        equals_format.setForeground(QColor("#505050"))
        self.rules.append((QRegExp("\\s*=\\s*"), equals_format))

        self.error_format = QTextCharFormat()
        self.error_format.setUnderlineColor(Qt.red)
        self.error_format.setUnderlineStyle(QTextCharFormat.SpellCheckUnderline)


    def highlightBlock(self, text):
        block_nr = self.currentBlock().blockNumber() + 1
        is_error = block_nr in self.error_lines
        if is_error:
             self.setFormat(0, len(text), self.error_format)

        specific_rules = self.rules[1:]
        for pattern, fmt in specific_rules:
            expression = QRegExp(pattern)
            index = expression.indexIn(text)
            while index >= 0:
                length = expression.matchedLength()
                self.setFormat(index, length, fmt)
                index = expression.indexIn(text, index + length)

        comment_rule = self.rules[0]
        pattern, fmt = comment_rule
        expression = QRegExp(pattern)
        index = expression.indexIn(text)
        if index >= 0:
             length = expression.matchedLength()
             self.setFormat(index, length, fmt)


    def set_error_lines(self, lines):
        self.error_lines = set(lines)
        self.rehighlight()


# --- License Manager Main Window ---
class LicenseManager(QMainWindow):
    log_signal = pyqtSignal(str, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_path = ""
        self.current_file_content_hash = None
        self.license_data = {}
        self.system_licenses = {}
        self.conflicts = {}
        self.log_visible = True
        self.install_system_wide_cb = None

        self.setup_ui()
        self.setWindowTitle("License Manager")
        self.setMinimumSize(800, 600)
        self.update_ui_state()

    def setup_ui(self):
        toolbar = QToolBar("Main Toolbar")
        toolbar.setIconSize(QSize(16, 16))
        self.addToolBar(toolbar)

        self.action_open = QAction(QIcon.fromTheme("document-open", QIcon()), "&Open...", self)
        self.action_open.setShortcut("Ctrl+O")
        self.action_open.setToolTip("Open an .ngit license file (Ctrl+O)")
        self.action_open.triggered.connect(self.open_file)
        toolbar.addAction(self.action_open)

        self.action_generate = QAction(QIcon.fromTheme("document-new", QIcon()), "&Generate Template...", self)
        self.action_generate.setShortcut("Ctrl+N")
        self.action_generate.setToolTip("Generate a sample .ngit file (Ctrl+N)")
        self.action_generate.triggered.connect(self.generate_sample)
        toolbar.addAction(self.action_generate)

        toolbar.addSeparator()

        self.action_validate_file = QAction(QIcon.fromTheme("document-properties", QIcon()), "&Validate File", self)
        self.action_validate_file.setShortcut("F5")
        self.action_validate_file.setToolTip("Validate the syntax and format of the current file (F5)")
        self.action_validate_file.triggered.connect(self.validate_file_action)
        toolbar.addAction(self.action_validate_file)

        self.action_check_system = QAction(QIcon.fromTheme("system-search", QIcon()), "Check &System", self)
        self.action_check_system.setShortcut("F6")
        self.action_check_system.setToolTip("Check currently active system environment licenses (F6)")
        self.action_check_system.triggered.connect(self.check_system_action)
        toolbar.addAction(self.action_check_system)

        toolbar.addSeparator()

        self.action_install = QAction(QIcon.fromTheme("system-software-install", QIcon()), "&Install Licenses", self)
        self.action_install.setShortcut("Ctrl+I")
        self.action_install.setToolTip("Install the validated licenses from the file (Ctrl+I)")
        self.action_install.triggered.connect(self.install_licenses)
        toolbar.addAction(self.action_install)

        self.install_system_wide_cb = QCheckBox("System-Wide (Admin)")
        self.install_system_wide_cb.setToolTip("Install for all users (requires administrator privileges)")
        toolbar.addWidget(self.install_system_wide_cb)

        self.conflict_indicator_label = QLabel("⚠️")
        self.conflict_indicator_label.setToolTip("Conflicts detected between file and system licenses!")
        self.conflict_indicator_label.setStyleSheet("color: orange; font-weight: bold;")
        self.conflict_indicator_label.setVisible(False)
        toolbar.addWidget(self.conflict_indicator_label)

        spacer = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        toolbar.addWidget(spacer)

        self.action_toggle_log = QAction(QIcon.fromTheme("view-list-text", QIcon()), "&Toggle Logs", self)
        self.action_toggle_log.setShortcut("Ctrl+L")
        self.action_toggle_log.setToolTip("Show/Hide the log panel (Ctrl+L)")
        self.action_toggle_log.triggered.connect(self.toggle_log)
        toolbar.addAction(self.action_toggle_log)

        self.action_clear_log = QAction(QIcon.fromTheme("edit-clear", QIcon()), "C&lear Log", self)
        self.action_clear_log.setShortcut("Ctrl+Shift+L")
        self.action_clear_log.setToolTip("Clear all messages from the log panel (Ctrl+Shift+L)")
        self.action_clear_log.triggered.connect(self.clear_log)
        toolbar.addAction(self.action_clear_log)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(0, 0, 0, 0)

        self.txt_content = CodeEditor()
        self.txt_content.setStyleSheet("QPlainTextEdit { background-color: #ffffff; font-family: Consolas, monospace; font-size: 10pt; }")
        self.highlighter = NGITHighlighter(self.txt_content.document())
        self.txt_content.textChanged.connect(self.mark_dirty)

        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet("background-color: #f0f0f0; font-size: 9pt;")
        self.txt_log.setMaximumHeight(150)

        layout.addWidget(self.txt_content)
        layout.addWidget(self.txt_log)

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_label = QLabel("Ready")
        self.status_bar.addWidget(self.status_label, 1)

        self.setup_icons()

    def setup_icons(self):
        self.action_open.setIcon(QIcon.fromTheme("document-open", QIcon()))
        self.action_generate.setIcon(QIcon.fromTheme("document-new", QIcon()))
        self.action_validate_file.setIcon(QIcon.fromTheme("document-properties", QIcon()))
        self.action_check_system.setIcon(QIcon.fromTheme("system-search", QIcon()))
        self.action_install.setIcon(QIcon.fromTheme("system-software-install", QIcon()))
        self.action_toggle_log.setIcon(QIcon.fromTheme("view-list-text", QIcon()))
        self.action_clear_log.setIcon(QIcon.fromTheme("edit-clear", QIcon()))


    def update_ui_state(self):
        has_file = bool(self.file_path)
        is_valid = has_file and bool(self.license_data) and not self.highlighter.error_lines
        has_conflicts = bool(self.conflicts)

        self.action_validate_file.setEnabled(has_file)
        # Only enable install on supported platforms (Windows or macOS for user install)
        can_install = is_valid and (platform.system() in ['Windows', 'Darwin'])
        self.action_install.setEnabled(can_install)
        self.install_system_wide_cb.setEnabled(can_install and platform.system() == 'Windows') # System only for Windows currently

        self.conflict_indicator_label.setVisible(has_conflicts and is_valid)

        if not has_file:
            self.status_label.setText("Ready. Open an .ngit file or generate a template.")
            self.setWindowTitle("License Manager")
        else:
            file_status = "Validated" if is_valid else ("Validation Failed" if self.highlighter.error_lines else "Not Validated")
            conflict_status = " Conflicts Detected!" if has_conflicts else ""
            dirty_status = " (Modified)" if self.is_dirty() else ""
            self.status_label.setText(f"File: {os.path.basename(self.file_path)}{dirty_status} [{file_status}{conflict_status}]")
            self.setWindowTitle(f"License Manager - {os.path.basename(self.file_path)}{'*' if self.is_dirty() else ''}")


    def toggle_log(self):
        self.log_visible = not self.log_visible
        self.txt_log.setVisible(self.log_visible)


    def clear_log(self):
        self.txt_log.clear()


    def log(self, level, message):
        timestamp = time.strftime("%H:%M:%S")
        icon_map = {'error': '❌ ERROR','warning': '⚠️ WARN','info': 'ℹ️ INFO','success': '✅ OK'}
        color_map = {'error': '#c62828','warning': '#e65100','info': '#37474f','success': '#2e7d32','system': '#1a237e','detail': '#455a64'}
        prefix = icon_map.get(level, '➡️')
        color = color_map.get(level, '#000000')
        is_conflict_detail = level == 'warning' and "System=" in message
        font_weight = "bold" if level in ['error', 'warning'] and not is_conflict_detail else "normal"
        if is_conflict_detail: color = color_map.get('detail')
        formatted_message = (f"<span style='color: #777;'>[{timestamp}]</span> <span style='color: {color}; font-weight: {font_weight};'>{prefix}: {message}</span>")
        self.txt_log.append(formatted_message)
        self.log_signal.emit(level, message)
        QApplication.processEvents()


    def generate_sample(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Sample Template", "template.ngit", "NGIT Files (*.ngit)")
        if not path: return
        try:
            with open(path, 'w', encoding='utf-8') as f: f.write(generate_ngit_template())
            self.log("success", f"Sample template created: {os.path.basename(path)}")
            self._load_file_content(path)
        except Exception as e:
            self.log("error", f"Failed to create sample: {str(e)}")


    def open_file(self):
        if self.is_dirty():
             reply = QMessageBox.question(self, 'Unsaved Changes',
                                          "You have unsaved changes. Open new file and discard changes?",
                                          QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
             if reply == QMessageBox.No: return
        path, _ = QFileDialog.getOpenFileName(self, "Open License File", "", "NGIT Files (*.ngit)")
        if not path: return
        self._load_file_content(path)


    def _load_file_content(self, path):
        try:
            with open(path, 'r', encoding='utf-8') as f: content = f.read()
            self.file_path = path
            self.txt_content.textChanged.disconnect(self.mark_dirty)
            self.txt_content.setPlainText(content)
            self.txt_content.textChanged.connect(self.mark_dirty)
            self.current_file_content_hash = hash(content)
            self.log("info", f"File loaded: {os.path.basename(path)}")
            self.validate_file_action(triggered_by_load=True)
            self.txt_content.document().clearUndoRedoStacks()
            self.txt_content.moveCursor(self.txt_content.textCursor().Start)
        except Exception as e:
            self.log("error", f"Error loading file {os.path.basename(path)}: {str(e)}")
            self.file_path = ""
            self.current_file_content_hash = None
            self.txt_content.clear()
        finally:
            self.update_ui_state()


    def mark_dirty(self):
        self.update_ui_state()


    def is_dirty(self):
        if not self.file_path or self.current_file_content_hash is None: return False
        current_hash = hash(self.txt_content.toPlainText())
        return current_hash != self.current_file_content_hash


    def validate_file_action(self, triggered_by_load=False):
        if not self.file_path:
            self.log("warning", "No file open to validate.")
            return
        if not triggered_by_load:
             self.log("info", "Manual validation requested...")
        self.log("info", f"Starting validation for {os.path.basename(self.file_path)}...")
        self.status_label.setText(f"Validating {os.path.basename(self.file_path)}...")
        QApplication.processEvents()
        current_content = self.txt_content.toPlainText()
        self.worker = FileValidator(current_content)
        self.worker.validation_complete.connect(self.handle_validation_result)
        self.worker.start()


    def check_system_action(self):
        self.log("info", "Checking system environment licenses...")
        self.status_label.setText("Checking system licenses...")
        QApplication.processEvents()
        self.system_licenses = self._get_system_licenses()
        if not self.system_licenses:
            self.log("warning", "No relevant licenses found in system environment.")
        else:
            self.log("info", "Active system licenses found:")
            for key, value in self.system_licenses.items():
                self.log("system", f"  {key} = {value}")
        self.status_label.setText("System check complete.")
        if self.license_data: self.check_license_conflicts()
        self.update_ui_state()


    def _get_system_licenses(self):
        found_licenses = {}
        for key in MANDATORY_LICENSES.keys():
            value = os.environ.get(key)
            if value is not None: found_licenses[key] = value
        return found_licenses


    def handle_validation_result(self, errors, data):
        error_lines = set()
        if errors:
            self.log("error", f"Validation failed for {os.path.basename(self.file_path)}:")
            max_errors_to_show = 10
            for i, (line_num, err_msg) in enumerate(errors):
                if i < max_errors_to_show: self.log("error", f"  Line {line_num}: {err_msg}")
                error_lines.add(line_num)
            if len(errors) > max_errors_to_show: self.log("error", f"  ... and {len(errors) - max_errors_to_show} more errors.")
            self.license_data = {}
        else:
            self.log("success", f"Validation passed for {os.path.basename(self.file_path)}.")
            self.license_data = data
            self.log("info", "Validated licenses in file:")
            for k, v in data.items(): self.log("detail", f"  {k} = {v}")
        self.highlighter.set_error_lines(error_lines)
        self.check_license_conflicts()
        self.update_ui_state()


    def check_license_conflicts(self):
        self.conflicts = {}
        if not self.license_data: return
        self.system_licenses = self._get_system_licenses()
        found_conflicts = {}
        for key, file_value in self.license_data.items():
            system_value = self.system_licenses.get(key)
            if system_value is not None and system_value != file_value:
                found_conflicts[key] = {'system': system_value, 'file': file_value}
        if found_conflicts:
            self.log("warning", "License conflicts DETECTED:")
            for key, values in found_conflicts.items():
                self.log("warning", f"  {key}: System='{values['system']}' vs File='{values['file']}'")
        else:
             if self.system_licenses: self.log("info", "No conflicts found between file and system licenses.")
        self.conflicts = found_conflicts


    def install_licenses(self):
        if not self.license_data or self.highlighter.error_lines:
            self.log("error", "Cannot install: Validation failed or no valid data loaded.")
            QMessageBox.critical(self, "Installation Error", "Cannot install licenses because the file is invalid or hasn't been validated successfully.")
            return
        if self.conflicts:
            if not self._confirm_overwrite(self.conflicts):
                self.log("info", "Installation cancelled by user due to conflicts.")
                return

        system_wide = self.install_system_wide_cb.isChecked()
        target = "System-Wide (All Users)" if system_wide else "Current User Only"
        self.log("info", f"Starting installation for: {target}...")
        self.status_label.setText(f"Installing ({target})...")
        QApplication.processEvents()
        install_ok = False
        try:
            current_os = platform.system()
            if current_os == 'Windows':
                if not winreg: raise OSError("Windows Registry access module (winreg) not available.")
                if system_wide: install_ok = self._install_windows_registry(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment", True)
                else: install_ok = self._install_windows_registry(winreg.HKEY_CURRENT_USER, "Environment", False)
            # REMOVED Linux elif block
            elif current_os == 'Darwin': # macOS
                 # System-wide on macOS not implemented, force user install
                 if system_wide:
                     self.log("warning", "System-wide installation on macOS is complex and not implemented. Installing for current user only.")
                 install_ok = self._install_macos_user() # Only user install for macOS
            else:
                 self.log("error", f"Unsupported operating system for installation: {current_os}")
                 QMessageBox.warning(self, "Unsupported OS", f"License installation is not supported on this operating system ({current_os}).")

        except PermissionError as pe:
             priv = "administrator/root" if system_wide and platform.system() == 'Windows' else "write" # Adjust privilege name
             self.log("error", f"Installation failed: Permission denied. Ensure you have sufficient ({priv}) privileges.")
             QMessageBox.critical(self, "Permission Error", f"Installation failed due to insufficient privileges.\nTry running as {priv} or check permissions.")
             install_ok = False
        except Exception as e:
            self.log("error", f"Installation failed: {str(e)}")
            QMessageBox.critical(self, "Installation Error", f"An unexpected error occurred during installation:\n{e}")
            install_ok = False
        if install_ok:
            self.log("success", "License installation completed successfully.")
            self.status_label.setText("Installation successful.")
            self.check_system_action() # Re-check system state
        else:
            # Check if error was due to unsupported OS and avoid redundant message
            if not (platform.system() not in ['Windows', 'Darwin']):
                self.status_label.setText("Installation failed.")
            else:
                 self.status_label.setText("Installation unavailable on this OS.")
        self.update_ui_state()


    def _install_windows_registry(self, root_key, key_path, requires_admin):
        if not winreg: return False
        from ctypes import WinError, windll
        key = None
        root_hkey = None
        try:
            try:
                root_hkey = winreg.ConnectRegistry(None, root_key)
                access_mask = winreg.KEY_READ | winreg.KEY_WRITE | winreg.KEY_SET_VALUE
                key = winreg.OpenKey(root_hkey, key_path, 0, access_mask)
            except PermissionError:
                hk_name = 'HKLM' if root_key == winreg.HKEY_LOCAL_MACHINE else 'HKCU'
                self.log("error", f"Permission denied to access registry key: {hk_name}\\{key_path}")
                if requires_admin: self.log("error", "Try running the application as Administrator.")
                raise
            except FileNotFoundError:
                hk_name = 'HKLM' if root_key == winreg.HKEY_LOCAL_MACHINE else 'HKCU'
                self.log("error", f"Registry key not found: {hk_name}\\{key_path}")
                return False
            set_errors = 0
            for var_name, value in self.license_data.items():
                try:
                    winreg.SetValueEx(key, var_name, 0, winreg.REG_SZ, value)
                    self.log("detail", f"Set registry value: {var_name}")
                except OSError as oe:
                    set_errors += 1; win_err_code = getattr(oe, 'winerror', 'N/A')
                    self.log("error", f"Failed to set registry value {var_name}: WinError {win_err_code}")
                except Exception as e:
                    set_errors += 1; self.log("error", f"Unexpected error setting registry value {var_name}: {e}")
            if set_errors > 0: self.log("error", f"{set_errors} error(s) occurred while setting registry values."); return False
            try:
                HWND_BROADCAST = 0xFFFF; WM_SETTINGCHANGE = 0x001A; SMTO_ABORTIFHUNG = 0x0002
                result = windll.user32.SendMessageTimeoutW(HWND_BROADCAST, WM_SETTINGCHANGE, 0, "Environment", SMTO_ABORTIFHUNG, 5000, None)
                if result == 0: self.log("warning", "Failed to broadcast environment change message (result=0). Manual refresh may be needed.")
                else: self.log("info", "Environment change message broadcast.")
            except Exception as e: self.log("warning", f"Broadcasting environment change failed: {e}")
            return True
        finally:
            # Corrected indentation for try/except within finally
            if key:
                try:
                    winreg.CloseKey(key)
                except Exception as close_e:
                    print(f"Ignoring error closing registry key: {close_e}") # Non-critical usually
            if root_hkey:
                try:
                    winreg.CloseKey(root_hkey) # Close registry handle
                except Exception as close_e:
                    print(f"Ignoring error closing registry handle: {close_e}") # Non-critical usually

    # --- REMOVED Linux Install Methods ---
    # _install_linux_system
    # _install_linux_user

    def _install_macos_user(self):
         """Handles macOS user-specific install (~/.zprofile or ~/.bash_profile). Returns True/False."""
         if platform.system() != 'Darwin': return False
         zprofile = os.path.expanduser("~/.zprofile")
         bash_profile = os.path.expanduser("~/.bash_profile")
         target_file = zprofile if os.path.exists(zprofile) else bash_profile
         lines_to_add = ["\n# --- License Manager Settings (User) ---"]
         for k, v in self.license_data.items():
             v_escaped = v.replace("'", "'\\''")
             lines_to_add.append(f"export {k}='{v_escaped}'")
             self.log("detail", f"Prepared user env (macOS): export {k}='{v_escaped}'")
         try:
             existing_content = ""; start_marker = "# --- License Manager Settings (User) ---"
             if os.path.exists(target_file):
                  with open(target_file, "r", encoding='utf-8') as f: existing_content = f.read()
             if start_marker in existing_content: self.log("warning", f"License Manager block marker found in {target_file}. Appending settings again. Manual cleanup might be needed.")
             with open(target_file, "a", encoding='utf-8') as f: f.write("\n".join(lines_to_add) + "\n")
             self.log("info", f"User environment settings appended to {target_file}")
             self.log("info", f"Open a new terminal window or run 'source {os.path.basename(target_file)}' to apply changes.")
             return True
         except Exception as e:
             self.log("error", f"Failed to write to {target_file}: {str(e)}")
             return False


    def _confirm_overwrite(self, conflicts_dict):
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Confirm Overwrite")
        msg.setText(f"Found {len(conflicts_dict)} license(s) in the file that conflict with existing system settings.")
        conflict_list = []
        max_show = 5
        for i, (key, values) in enumerate(conflicts_dict.items()):
             if i >= max_show: conflict_list.append(f"... and {len(conflicts_dict) - max_show} more."); break
             conflict_list.append(f" • {key}:\n    System = {values['system']}\n    File     = {values['file']}")
        msg.setInformativeText("Do you want to overwrite the existing system settings with the values from the file?")
        msg.setDetailedText("\n\n".join(conflict_list))
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setDefaultButton(QMessageBox.No)
        return msg.exec_() == QMessageBox.Yes


    def closeEvent(self, event):
         if self.is_dirty():
              reply = QMessageBox.question(self, 'Unsaved Changes',
                                           "You have unsaved changes. Do you want to discard them and close?",
                                           QMessageBox.Discard | QMessageBox.Cancel, QMessageBox.Cancel)
              if reply == QMessageBox.Discard: event.accept()
              else: event.ignore()
         else: event.accept()


# --- File Validator Thread ---
class FileValidator(QThread):
    validation_complete = pyqtSignal(list, dict)
    def __init__(self, content): super().__init__(); self.content = content
    def run(self):
        errors = []; license_data = {}; mandatory_found = False; seen_keys = set()
        lines = self.content.splitlines()
        for line_num, line in enumerate(lines, 1):
            stripped_line = line.strip()
            if not stripped_line or stripped_line.startswith('#'): continue
            if '=' not in stripped_line:
                if re.match(r'^[A-Z_][A-Z0-9_]*\s*$', stripped_line): errors.append((line_num, f"Missing '=' and value for key '{stripped_line}'"))
                else: errors.append((line_num, f"Invalid format (missing '=') in '{stripped_line[:50]}'"))
                continue
            parts = stripped_line.split('=', 1); key = parts[0].strip(); value = parts[1].strip()
            if key in seen_keys: errors.append((line_num, f"Duplicate key '{key}' found")); continue
            seen_keys.add(key)
            if not re.fullmatch(r'[A-Z_][A-Z0-9_]*', key): errors.append((line_num, f"Invalid characters or format in key name '{key}'")); continue
            if key not in MANDATORY_LICENSES: errors.append((line_num, f"Unknown license key '{key}'")); continue
            regex_pattern = MANDATORY_LICENSES.get(key)
            if regex_pattern:
                 is_path_key = key.endswith('_FILE') or key.endswith('_LICENSE_FILE')
                 if platform.system() == "Windows" and is_path_key and not value.startswith('\\\\') and ':' not in value and not re.match(r'^\d+@', value):
                      pass # Heuristic check for bad Windows paths
                 try:
                      if not re.fullmatch(regex_pattern, value): errors.append((line_num, f"Invalid value format for {key}")); continue
                 except re.error as re_err: errors.append((line_num, f"Internal regex error for {key}: {re_err}")); continue
            license_data[key] = value
            if key in MANDATORY_LICENSES: mandatory_found = True
        if not any(k in license_data for k in MANDATORY_LICENSES):
            if not errors: errors.append((len(lines) + 1, "No valid mandatory licenses found in the file"))
        errors.sort(key=lambda x: x[0])
        self.validation_complete.emit(errors, license_data)


# --- Constants ---
MANDATORY_LICENSES = {
    "CMG_LIC_HOST": r"^\d+@[\w\.\-]+$",
    "SLBSLS_LICENSE_FILE": r"^(\d+@[\w\.\-]+|([a-zA-Z]:\\|[\\/]{1,2})[\w\.\\\/\s\-\(\)]+\.lic)$",
    "LM_LICENSE_FILE": r"^(\d+@[\w\.\-]+|([a-zA-Z]:\\|[\\/]{1,2})[\w\.\\\/\s\-\(\)]+\.lic)$",
    "SCPLMD_LICENSE_FILE": r"^([a-zA-Z]:\\|[\\/]{1,2})[\w\.\\\/\s\-\(\)]+\.lic$",
    "tNavigator_LICENSE_SERVER": r"^([\w\.\-]+(:\d+)?|https?://[\w\.\-:%/]+)(;\S+)*$"
}


def generate_ngit_template():
    return """# NGIT License Configuration File v1.5
# Lines starting with # are comments.
# Ensure values are correct for your environment.
# At least one of the following licenses is typically required.

# Example: Port@Hostname or Port@IP_Address
CMG_LIC_HOST = 2700@license-server.example.com

# Example: Port@Hostname or a valid file path to a .lic file
SLBSLS_LICENSE_FILE = C:\\ProgramData\\Schlumberger\\Schlumberger Licensing\\slbsls.lic
# SLBSLS_LICENSE_FILE = 27027@another-server.local

# Example: Port@Hostname or a valid file path to a .lic file
LM_LICENSE_FILE = D:\\licenses\\flexlm\\product.lic
# LM_LICENSE_FILE = 27000@primary-lic-server

# Example: Must be a valid file path to a .lic file
SCPLMD_LICENSE_FILE = E:\\Licenses\\Company\\scplmd.lic

# Example: Hostname or Hostname:Port or URL (Check required format)
tNavigator_LICENSE_SERVER = tnav-license.corp.net:5053
# tNavigator_LICENSE_SERVER = https://tnav-license.cloud.com/auth
"""

# --- Main Execution ---
if __name__ == "__main__":
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    window = LicenseManager()
    window.show()
    sys.exit(app.exec_())