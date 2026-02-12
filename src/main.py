"""DOCX Batch Updater - Main Entry Point.

This is the main entry point for the DOCX Batch Updater application.
It initializes the GUI application and starts the main event loop.
"""

import sys
import os

# 确保 src 目录在 sys.path 中（打包后需要）
if getattr(sys, 'frozen', False):
    # 打包后的环境
    _src_path = os.path.join(sys._MEIPASS, 'src')
else:
    # 开发环境
    _src_path = os.path.dirname(os.path.abspath(__file__))

if _src_path not in sys.path:
    sys.path.insert(0, _src_path)

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
    app.setApplicationName("DOCX 批量更新器")
    app.setOrganizationName("DOCX 批量更新器")

    # Create and show main window
    window = MainWindow()
    window.show()

    # Run application
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())

# policy-guard test change

# ruleA2 test change
