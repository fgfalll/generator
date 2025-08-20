import sys
from PyQt6.QtWidgets import QApplication
from ui import MainWindow
from backend import Backend
import theme_manager

if __name__ == "__main__":
    app = QApplication(sys.argv)

    settings = theme_manager.load_settings()
    app.setStyleSheet(theme_manager.generate_qss(settings["theme_name"], settings["custom_colors"]))

    main_window = MainWindow()
    backend_logic = Backend(ui=main_window, app=app)
    backend_logic.connect_signals()

    main_window.show()
    sys.exit(app.exec())