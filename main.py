import sys
import traceback
from PyQt6.QtWidgets import QApplication, QMessageBox
from ui import MainWindow
from backend import Backend

def global_exception_handler(exc_type, exc_value, exc_traceback):
    """Catch unhandled exceptions and show them in a GUI dialog instead of crashing silently."""
    print("Неочікувана помилка!", file=sys.stderr)
    traceback.print_exception(exc_type, exc_value, exc_traceback)
    
    # Try to grab the active application to parent the dialog
    app = QApplication.instance()
    if app:
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Icon.Critical)
        msg_box.setWindowTitle("Критична помилка")
        msg_box.setText("Сталася неочікувана помилка в програмі.\n\n"
                        f"{exc_type.__name__}: {exc_value}")
        
        details = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        msg_box.setDetailedText(details)
        msg_box.exec()

sys.excepthook = global_exception_handler

import os

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    from PyQt6.QtGui import QIcon
    app = QApplication(sys.argv)
    
    # Set global application icon
    app_icon = QIcon(resource_path("icon.ico"))
    app.setWindowIcon(app_icon)

    main_window = MainWindow()
    backend_logic = Backend(ui=main_window, app=app)
    backend_logic.connect_signals()

    main_window.show()
    sys.exit(app.exec())
