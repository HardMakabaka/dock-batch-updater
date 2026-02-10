"""DOCX Batch Updater - Main Entry Point.

This is the main entry point for the DOCX Batch Updater application.
It initializes the GUI application and starts the main event loop.
"""

import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt

from gui.main_window import MainWindow


def main() -> int:
    """Main entry point for the application.

    Returns:
        Exit code (0 for success, non-zero for error)
    """
    # Enable high DPI scaling
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    # Create application instance
    app = QApplication(sys.argv)
    app.setApplicationName("DOCX Batch Updater")
    app.setOrganizationName("DOCX Batch Updater")

    # Create and show main window
    window = MainWindow()
    window.show()

    # Run application
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
