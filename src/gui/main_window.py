"""Main application window for DOCX Batch Updater.

This module provides the main GUI window that integrates all components
of the application.
"""

import os
from typing import Optional
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QCheckBox, QSplitter, QMessageBox,
    QFileDialog, QApplication
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon

from .widgets import FileListWidget, ReplacementRulesWidget, ProgressWidget, LogWidget
from core.batch_processor import BatchProcessor, ProcessingResult


class MainWindow(QMainWindow):
    """Main application window."""

    logMessage = pyqtSignal(str, str)
    progressUpdated = pyqtSignal(int, int)
    statisticsUpdated = pyqtSignal(object)

    def __init__(self):
        """Initialize main window."""
        super().__init__()

        self.batch_processor = BatchProcessor()
        self.is_processing = False

        self.init_ui()
        self.setup_connections()

    def init_ui(self):
        """Initialize user interface."""
        self.setWindowTitle("DOCX 批量更新器")
        self.setGeometry(100, 100, 1000, 700)
        self.setMinimumSize(800, 600)

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Main layout
        main_layout = QVBoxLayout(central_widget)

        # Create splitter for resizable sections
        main_splitter = QSplitter(Qt.Horizontal)

        # Left panel - Files and Rules
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        # File list widget
        self.file_list_widget = FileListWidget()
        left_layout.addWidget(self.file_list_widget)

        # Replacement rules widget
        self.rules_widget = ReplacementRulesWidget()
        left_layout.addWidget(self.rules_widget)

        # Backup checkbox
        backup_layout = QHBoxLayout()
        self.backup_checkbox = QCheckBox("创建备份文件")
        self.backup_checkbox.setChecked(True)
        backup_layout.addWidget(self.backup_checkbox)

        self.backup_dir_btn = QPushButton("选择备份目录…")
        self.backup_dir_btn.setEnabled(False)
        self.backup_dir_btn.clicked.connect(self.select_backup_directory)
        backup_layout.addWidget(self.backup_dir_btn)

        left_layout.addLayout(backup_layout)

        # Right panel - Progress and Log
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        # Progress widget
        self.progress_widget = ProgressWidget()
        right_layout.addWidget(self.progress_widget)

        # Log widget
        self.log_widget = LogWidget()
        right_layout.addWidget(self.log_widget)

        # Add panels to splitter
        main_splitter.addWidget(left_widget)
        main_splitter.addWidget(right_widget)
        main_splitter.setSizes([500, 500])

        main_layout.addWidget(main_splitter)

        # Bottom control buttons
        control_layout = QHBoxLayout()

        control_layout.addStretch()

        self.process_btn = QPushButton("开始处理")
        self.process_btn.setMinimumHeight(40)
        self.process_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.process_btn.clicked.connect(self.start_processing)
        self.process_btn.setEnabled(False)
        control_layout.addWidget(self.process_btn)

        self.stop_btn = QPushButton("停止")
        self.stop_btn.setMinimumHeight(40)
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.setEnabled(False)
        control_layout.addWidget(self.stop_btn)

        main_layout.addLayout(control_layout)

        # Backup directory path
        self.backup_dir = ""

    def setup_connections(self):
        """Setup signal connections."""
        # Enable/disable process button based on file and rule counts
        self.file_list_widget.filesChanged.connect(self.update_process_button)
        self.rules_widget.rulesChanged.connect(self.update_process_button)

        # Backup checkbox state
        self.backup_checkbox.stateChanged.connect(self.on_backup_checkbox_changed)

        # Thread-safe UI update signals
        self.logMessage.connect(self.log_widget.log)
        self.progressUpdated.connect(self.progress_widget.set_progress)
        self.statisticsUpdated.connect(self.progress_widget.set_statistics)

    def on_backup_checkbox_changed(self, state: int) -> None:
        """Handle backup checkbox state change.

        Args:
            state: Checkbox state (Qt.Checked or Qt.Unchecked)
        """
        self.backup_dir_btn.setEnabled(state == Qt.Checked)
        if state == Qt.Unchecked:
            self.backup_dir = ""

    def select_backup_directory(self) -> None:
        """Open dialog to select backup directory."""
        directory = QFileDialog.getExistingDirectory(
            self,
            "选择备份目录"
        )

        if directory:
            self.backup_dir = directory
            self.log_widget.log(f"备份目录已设置为：{directory}", "INFO")

    def update_process_button(self) -> None:
        """Update process button state based on file and rule counts."""
        has_files = len(self.file_list_widget.get_files()) > 0
        has_rules = len(self.rules_widget.get_rules()) > 0

        self.process_btn.setEnabled(has_files and has_rules and not self.is_processing)

    def start_processing(self) -> None:
        """Start batch processing of documents."""
        files = self.file_list_widget.get_files()
        rules = self.rules_widget.get_rules()

        if not files:
            QMessageBox.warning(self, "未选择文件", "请添加需要处理的文件。")
            return

        if not rules:
            QMessageBox.warning(self, "未设置规则", "请添加替换规则。")
            return

        # Confirm before processing
        reply = QMessageBox.question(
            self,
            "确认处理",
            f"将处理 {len(files)} 个文件，并应用 {len(rules)} 条替换规则。是否继续？",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.No:
            return

        # Set processing state
        self.is_processing = True
        self.process_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.file_list_widget.setEnabled(False)
        self.rules_widget.setEnabled(False)
        self.backup_checkbox.setEnabled(False)
        self.backup_dir_btn.setEnabled(False)

        # Reset progress
        self.progress_widget.reset()
        self.progress_widget.set_status("开始处理…")

        # Log start
        self.log_widget.log("=" * 50, "INFO")
        self.log_widget.log("开始批量处理", "INFO")
        self.log_widget.log(f"文件数：{len(files)}", "INFO")
        self.log_widget.log(f"规则数：{len(rules)}", "INFO")
        self.log_widget.log(f"备份：{'是' if self.backup_checkbox.isChecked() else '否'}", "INFO")
        if self.backup_checkbox.isChecked() and self.backup_dir:
            self.log_widget.log(f"备份目录：{self.backup_dir}", "INFO")
        self.log_widget.log("=" * 50, "INFO")

        # Start processing in thread
        from PyQt5.QtCore import QThread

        class ProcessingThread(QThread):
            def __init__(self, processor, files, rules, create_backup, backup_dir,
                         progress_callback, result_callback):
                super().__init__()
                self.processor = processor
                self.files = files
                self.rules = rules
                self.create_backup = create_backup
                self.backup_dir = backup_dir
                self.progress_callback = progress_callback
                self.result_callback = result_callback
                self.error: Optional[str] = None

            def run(self):
                try:
                    self.processor.process_documents(
                        self.files,
                        self.rules,
                        self.create_backup,
                        self.backup_dir,
                        self.progress_callback,
                        self.result_callback
                    )
                except Exception as e:
                    self.error = str(e)

        self.processing_thread = ProcessingThread(
            self.batch_processor,
            files,
            rules,
            self.backup_checkbox.isChecked(),
            self.backup_dir,
            self.update_progress,
            self.handle_result
        )
        self.processing_thread.finished.connect(self.processing_finished)
        self.processing_thread.start()

    def stop_processing(self) -> None:
        """Stop current processing."""
        self.batch_processor.stop()
        self.log_widget.log("正在停止处理…", "WARNING")

    def update_progress(self, current: int, total: int) -> None:
        """Update progress display.

        Args:
            current: Current progress value
            total: Total value
        """
        self.progressUpdated.emit(current, total)

    def log_async(self, message: str, level: str = "INFO") -> None:
        """Thread-safe log helper."""
        self.logMessage.emit(message, level)

    def handle_result(self, result: ProcessingResult) -> None:
        """Handle individual document processing result.

        Args:
            result: Processing result
        """
        filename = os.path.basename(result.file_path)

        if result.success:
            self.log_async(
                f"{filename}：成功（替换 {result.replacements} 次）",
                "SUCCESS"
            )
        else:
            self.log_async(
                f"{filename}：失败 - {result.message}",
                "ERROR"
            )

        # Update statistics
        summary = self.batch_processor.get_summary()
        self.statisticsUpdated.emit(summary)

    def processing_finished(self) -> None:
        """Handle processing completion."""
        self.is_processing = False
        self.process_btn.setEnabled(
            len(self.file_list_widget.get_files()) > 0 and
            len(self.rules_widget.get_rules()) > 0
        )
        self.stop_btn.setEnabled(False)
        self.file_list_widget.setEnabled(True)
        self.rules_widget.setEnabled(True)
        self.backup_checkbox.setEnabled(True)
        self.backup_dir_btn.setEnabled(self.backup_checkbox.isChecked())

        thread_error = getattr(self.processing_thread, 'error', None)
        if thread_error:
            self.log_widget.log(f"处理线程异常：{thread_error}", "ERROR")
            QMessageBox.critical(self, "处理失败", f"处理过程中发生未捕获异常：\n{thread_error}")
            self.progress_widget.set_status("处理失败")
            return

        # Get final summary
        summary = self.batch_processor.get_summary()

        # Log completion
        self.log_widget.log("=" * 50, "INFO")
        self.log_widget.log("处理完成", "INFO")
        self.log_widget.log(f"总文件数：{summary['total_files']}", "INFO")
        self.log_widget.log(f"成功：{summary['successful']}", "INFO")
        self.log_widget.log(f"失败：{summary['failed']}", "INFO")
        self.log_widget.log(f"总替换次数：{summary['total_replacements']}", "INFO")
        self.log_widget.log(f"成功率：{summary['success_rate']:.1%}", "INFO")

        if summary['failed'] > 0:
            self.log_widget.log("\n失败文件：", "WARNING")
            for result in self.batch_processor.get_failed_results():
                self.log_widget.log(
                    f"  - {os.path.basename(result.file_path)}：{result.message}",
                    "WARNING"
                )

        self.log_widget.log("=" * 50, "INFO")

        # Show summary message
        self.progress_widget.set_status("处理完成")

        if summary['failed'] > 0:
            QMessageBox.warning(
                self,
                "处理完成（有错误）",
                f"已处理 {summary['total_files']} 个文件。\n"
                f"成功：{summary['successful']}\n"
                f"失败：{summary['failed']}\n"
                f"详情请查看日志。"
            )
        else:
            QMessageBox.information(
                self,
                "处理完成",
                f"成功处理 {summary['successful']} 个文件。\n"
                f"总替换次数：{summary['total_replacements']}"
            )
