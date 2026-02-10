"""Custom GUI widgets for DOCX Batch Updater.

This module provides custom widgets for the application including
file list management and replacement rule configuration.
"""

import os
from typing import List, Optional, Tuple, Callable
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget,
    QListWidgetItem, QLineEdit, QLabel, QCheckBox, QGroupBox,
    QFileDialog, QMessageBox, QProgressBar, QTextEdit, QSplitter
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor


class FileListWidget(QWidget):
    """Widget for managing the list of files to process."""

    filesChanged = pyqtSignal(int)  # Emits count of files

    def __init__(self, parent: Optional[QWidget] = None):
        """Initialize file list widget.

        Args:
            parent: Parent widget
        """
        super().__init__(parent)

        self.files: List[str] = []

        layout = QVBoxLayout(self)

        # File list
        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.SingleSelection)
        layout.addWidget(self.file_list)

        # Button row
        button_layout = QHBoxLayout()

        self.add_files_btn = QPushButton("添加文件")
        self.add_files_btn.clicked.connect(self.add_files)
        button_layout.addWidget(self.add_files_btn)

        self.add_folder_btn = QPushButton("添加文件夹")
        self.add_folder_btn.clicked.connect(self.add_folder)
        button_layout.addWidget(self.add_folder_btn)

        self.remove_file_btn = QPushButton("移除所选")
        self.remove_file_btn.clicked.connect(self.remove_selected)
        button_layout.addWidget(self.remove_file_btn)

        self.clear_all_btn = QPushButton("清空")
        self.clear_all_btn.clicked.connect(self.clear_all)
        button_layout.addWidget(self.clear_all_btn)

        layout.addLayout(button_layout)

        # File count label
        self.file_count_label = QLabel("已选择 0 个文件")
        layout.addWidget(self.file_count_label)

    def add_files(self) -> None:
        """Add files to the list via file dialog."""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择 DOCX 文件",
            "",
            "Word Documents (*.docx);;All Files (*.*)"
        )

        if files:
            for file in files:
                if file not in self.files:
                    self.files.append(file)
                    item = QListWidgetItem(os.path.basename(file))
                    item.setData(Qt.UserRole, file)
                    item.setToolTip(file)
                    self.file_list.addItem(item)

            self._update_count()

    def add_folder(self) -> None:
        """Add all DOCX files from a folder."""
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择文件夹"
        )

        if folder:
            docx_files = []
            for root, _, files in os.walk(folder):
                for file in files:
                    if file.lower().endswith('.docx'):
                        docx_files.append(os.path.join(root, file))

            for file in docx_files:
                if file not in self.files:
                    self.files.append(file)
                    item = QListWidgetItem(os.path.basename(file))
                    item.setData(Qt.UserRole, file)
                    item.setToolTip(file)
                    self.file_list.addItem(item)

            self._update_count()

    def remove_selected(self) -> None:
        """Remove the currently selected file from the list."""
        current_row = self.file_list.currentRow()
        if current_row >= 0:
            item = self.file_list.takeItem(current_row)
            file_path = item.data(Qt.UserRole)
            if file_path in self.files:
                self.files.remove(file_path)
            self._update_count()

    def clear_all(self) -> None:
        """Clear all files from the list."""
        self.files.clear()
        self.file_list.clear()
        self._update_count()

    def get_files(self) -> List[str]:
        """Get the list of selected files.

        Returns:
            List of file paths
        """
        return self.files.copy()

    def _update_count(self) -> None:
        """Update the file count label."""
        count = len(self.files)
        self.file_count_label.setText(f"已选择 {count} 个文件")
        self.filesChanged.emit(count)


class ReplacementRulesWidget(QWidget):
    """Widget for managing replacement rules."""

    rulesChanged = pyqtSignal(int)  # Emits count of rules

    def __init__(self, parent: Optional[QWidget] = None):
        """Initialize replacement rules widget.

        Args:
            parent: Parent widget
        """
        super().__init__(parent)

        self.rules: List[Tuple[str, str]] = []

        layout = QVBoxLayout(self)

        # Group box for rules
        rules_group = QGroupBox("替换规则")
        rules_layout = QVBoxLayout()

        # Rule list
        self.rule_list = QListWidget()
        self.rule_list.setSelectionMode(QListWidget.SingleSelection)
        rules_layout.addWidget(self.rule_list)

        # Add rule form
        form_layout = QHBoxLayout()

        form_layout.addWidget(QLabel("查找："))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("要查找的文本…")
        form_layout.addWidget(self.search_input)

        form_layout.addWidget(QLabel("替换为："))
        self.replace_input = QLineEdit()
        self.replace_input.setPlaceholderText("替换后的文本…")
        form_layout.addWidget(self.replace_input)

        rules_layout.addLayout(form_layout)

        # Buttons
        button_layout = QHBoxLayout()

        self.add_rule_btn = QPushButton("添加规则")
        self.add_rule_btn.clicked.connect(self.add_rule)
        button_layout.addWidget(self.add_rule_btn)

        self.remove_rule_btn = QPushButton("移除所选")
        self.remove_rule_btn.clicked.connect(self.remove_selected)
        button_layout.addWidget(self.remove_rule_btn)

        self.clear_rules_btn = QPushButton("清空规则")
        self.clear_rules_btn.clicked.connect(self.clear_all)
        button_layout.addWidget(self.clear_rules_btn)

        rules_layout.addLayout(button_layout)

        rules_group.setLayout(rules_layout)
        layout.addWidget(rules_group)

        # Rule count label
        self.rule_count_label = QLabel("已定义 0 条规则")
        layout.addWidget(self.rule_count_label)

    def add_rule(self) -> None:
        """Add a new replacement rule."""
        search_text = self.search_input.text().strip()
        replace_text = self.replace_input.text()

        if not search_text:
            QMessageBox.warning(self, "提示", "请输入要查找的文本。")
            return

        self.rules.append((search_text, replace_text))

        display_text = f"'{search_text}' → '{replace_text}'"
        item = QListWidgetItem(display_text)
        item.setData(Qt.UserRole, (search_text, replace_text))
        self.rule_list.addItem(item)

        # Clear inputs
        self.search_input.clear()
        self.replace_input.clear()
        self.search_input.setFocus()

        self._update_count()

    def remove_selected(self) -> None:
        """Remove the currently selected rule."""
        current_row = self.rule_list.currentRow()
        if current_row >= 0:
            item = self.rule_list.takeItem(current_row)
            rule_data = item.data(Qt.UserRole)
            if rule_data in self.rules:
                self.rules.remove(rule_data)
            self._update_count()

    def clear_all(self) -> None:
        """Clear all rules."""
        self.rules.clear()
        self.rule_list.clear()
        self._update_count()

    def get_rules(self) -> List[Tuple[str, str]]:
        """Get the list of replacement rules.

        Returns:
            List of (search_text, replace_text) tuples
        """
        return self.rules.copy()

    def _update_count(self) -> None:
        """Update the rule count label."""
        count = len(self.rules)
        self.rule_count_label.setText(f"已定义 {count} 条规则")
        self.rulesChanged.emit(count)


class ProgressWidget(QWidget):
    """Widget for displaying processing progress."""

    def __init__(self, parent: Optional[QWidget] = None):
        """Initialize progress widget.

        Args:
            parent: Parent widget
        """
        super().__init__(parent)

        layout = QVBoxLayout(self)

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # Progress label
        self.progress_label = QLabel("就绪")
        layout.addWidget(self.progress_label)

        # Statistics label
        self.stats_label = QLabel("")
        layout.addWidget(self.stats_label)

    def set_progress(self, current: int, total: int) -> None:
        """Update progress.

        Args:
            current: Current progress value
            total: Total value
        """
        if total > 0:
            percentage = int((current / total) * 100)
            self.progress_bar.setValue(percentage)
            self.progress_label.setText(f"处理中：{current}/{total} 个文件")

    def set_status(self, status: str) -> None:
        """Set status message.

        Args:
            status: Status message
        """
        self.progress_label.setText(status)

    def set_statistics(self, stats: dict) -> None:
        """Set statistics display.

        Args:
            stats: Dictionary of statistics
        """
        text = f"Total: {stats.get('total_files', 0)} | "
        text += f"成功: {stats.get('successful', 0)} | "
        text += f"失败: {stats.get('failed', 0)} | "
        text += f"替换次数: {stats.get('total_replacements', 0)}"
        self.stats_label.setText(text)

    def reset(self) -> None:
        """Reset progress widget."""
        self.progress_bar.setValue(0)
        self.progress_label.setText("就绪")
        self.stats_label.setText("")


class LogWidget(QWidget):
    """Widget for displaying processing log."""

    def __init__(self, parent: Optional[QWidget] = None):
        """Initialize log widget.

        Args:
            parent: Parent widget
        """
        super().__init__(parent)

        layout = QVBoxLayout(self)

        # Log text area
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Courier New", 9))
        layout.addWidget(self.log_text)

        # Buttons
        button_layout = QHBoxLayout()

        self.copy_btn = QPushButton("复制日志")
        self.copy_btn.clicked.connect(self.copy_log)
        button_layout.addWidget(self.copy_btn)

        self.clear_btn = QPushButton("清空日志")
        self.clear_btn.clicked.connect(self.clear_log)
        button_layout.addWidget(self.clear_btn)

        layout.addLayout(button_layout)

    def log(self, message: str, level: str = "INFO") -> None:
        """Add a message to the log.

        Args:
            message: Message to log
            level: Log level (INFO, WARNING, ERROR)
        """
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")

        color_map = {
            "INFO": "black",
            "WARNING": "#8B8000",  # Dark yellow
            "ERROR": "#B22222",  # Firebrick red
            "SUCCESS": "#228B22"  # Forest green
        }

        color = color_map.get(level, "black")

        html = f'<span style="color:gray;">[{timestamp}]</span> '
        html += f'<span style="color:{color}; font-weight:bold;">[{level}]</span> '
        html += f'<span>{message}</span><br>'

        self.log_text.append(html)

    def copy_log(self) -> None:
        """Copy log content to clipboard."""
        self.log_text.selectAll()
        self.log_text.copy()

    def clear_log(self) -> None:
        """Clear all log content."""
        self.log_text.clear()

    def get_text(self) -> str:
        """Get log content as plain text.

        Returns:
            Log content as plain text
        """
        return self.log_text.toPlainText()
