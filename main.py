import sys
from PyQt6.QtWidgets import QApplication
from ui import MainWindow
from backend import Backend

if __name__ == "__main__":
    app = QApplication(sys.argv)

    main_window = MainWindow()
    backend_logic = Backend(ui=main_window, app=app)
    backend_logic.connect_signals()

    main_window.show()
    sys.exit(app.exec())
