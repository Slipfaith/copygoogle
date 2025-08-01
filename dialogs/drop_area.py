import os
from typing import List, Tuple

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, QHeaderView, QComboBox,
    QTableWidgetItem, QSpinBox, QDialogButtonBox, QPushButton, QFrame,
    QFileDialog, QLineEdit, QWidget, QGroupBox, QGridLayout, QTextEdit,
    QScrollArea, QSizePolicy
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont, QPixmap, QPainter, QColor

class DropArea(QWidget):
    """Область для drag & drop файлов."""

    file_dropped = Signal(str)
    files_dropped = Signal(list)

    def __init__(self, accept_multiple=False):
        super().__init__()
        self.accept_multiple = accept_multiple
        self.setAcceptDrops(True)
        self.setMinimumHeight(120)
        self.setStyleSheet(
            """
            DropArea {
                border: 2px dashed #e0e0e0;
                border-radius: 12px;
                background-color: #fafafa;
            }
            """
        )

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)

        # Иконка
        icon_label = QLabel("📁")
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_label.setStyleSheet("font-size: 36px;")
        layout.addWidget(icon_label)

        # Основной текст
        if accept_multiple:
            self.label = QLabel("Перетащите Excel файлы сюда")
            self.file_label = QLabel("или нажмите для выбора файлов")
        else:
            self.label = QLabel("Перетащите Excel файл сюда")
            self.file_label = QLabel("или нажмите для выбора файла")

        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = QFont()
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setStyleSheet("color: #424242; font-weight: 500;")

        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_label.setStyleSheet("color: #757575; font-size: 10pt; margin-top: 5px;")

        layout.addWidget(self.label)
        layout.addWidget(self.file_label)

        self.setLayout(layout)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.open_file_dialog()
        super().mousePressEvent(event)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            valid_files = [u for u in urls if u.toLocalFile().lower().endswith(('.xlsx', '.xls'))]
            if valid_files:
                event.acceptProposedAction()
                self.setStyleSheet(
                    """
                    DropArea {
                        border: 2px solid #90caf9;
                        border-radius: 12px;
                        background-color: #e3f2fd;
                    }
                    """
                )

    def dragLeaveEvent(self, event):
        self.setStyleSheet(
            """
            DropArea {
                border: 2px dashed #e0e0e0;
                border-radius: 12px;
                background-color: #fafafa;
            }
            """
        )

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls() if
                 u.toLocalFile().lower().endswith(('.xlsx', '.xls'))]
        if files:
            if self.accept_multiple:
                self.files_dropped.emit(files)
                self.update_files_info(files)
            else:
                self.file_dropped.emit(files[0])
                self.update_file_info(files[0])

        self.setStyleSheet(
            """
            DropArea {
                border: 2px dashed #e0e0e0;
                border-radius: 12px;
                background-color: #fafafa;
            }
            """
        )

    def open_file_dialog(self):
        if self.accept_multiple:
            files, _ = QFileDialog.getOpenFileNames(
                self,
                "Выберите Excel файлы",
                "",
                "Excel Files (*.xlsx *.xls)"
            )
            if files:
                self.files_dropped.emit(files)
                self.update_files_info(files)
        else:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Выберите Excel файл",
                "",
                "Excel Files (*.xlsx *.xls)"
            )
            if file_path:
                self.file_dropped.emit(file_path)
                self.update_file_info(file_path)

    def update_file_info(self, file_path: str):
        file_name = os.path.basename(file_path)
        self.label.setText(f"✓ {file_name}")
        self.label.setStyleSheet("color: #1976d2; font-weight: 500;")
        self.file_label.setText(f"Размер: {self._get_file_size(file_path)}")

    def update_files_info(self, files: List[str]):
        self.label.setText(f"✓ Выбрано файлов: {len(files)}")
        self.label.setStyleSheet("color: #1976d2; font-weight: 500;")
        total_size = sum(os.path.getsize(f) for f in files)
        self.file_label.setText(f"Общий размер: {self._format_size(total_size)}")

    def _get_file_size(self, file_path: str) -> str:
        size = os.path.getsize(file_path)
        return self._format_size(size)

    def _format_size(self, size: int) -> str:
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"