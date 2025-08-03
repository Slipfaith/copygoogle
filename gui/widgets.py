from PySide6.QtWidgets import QTextBrowser, QGraphicsDropShadowEffect, QWidget, QVBoxLayout, QLabel, QFileDialog, QSizePolicy
from PySide6.QtCore import Signal, Qt
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QColor

import os
import subprocess
import platform
from typing import List
from . import styles


class ClickableTextEdit(QTextBrowser):
    """Text browser that opens file links on click."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setOpenLinks(False)
        self.anchorClicked.connect(self.handle_click)

    def handle_click(self, url):
        if hasattr(url, "isLocalFile") and url.isLocalFile():
            path = url.toLocalFile()
        else:
            url_str = str(url)
            if url_str.startswith("file://"):
                path = url_str.replace("file://", "")
            else:
                return

        if os.path.exists(path):
            if platform.system() == 'Windows':
                subprocess.run(['explorer', '/select,', path])
            elif platform.system() == 'Darwin':
                subprocess.run(['open', '-R', path])
            else:
                subprocess.run(['xdg-open', os.path.dirname(path)])


class ModernDropArea(QWidget):
    """Area that accepts drag & drop of files."""

    file_dropped = Signal(str)
    files_dropped = Signal(list)

    def __init__(self, accept_multiple=False):
        super().__init__()
        self.accept_multiple = accept_multiple
        self.setAcceptDrops(True)
        self.setFixedHeight(80)
        self.setMaximumWidth(400)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.setStyleSheet(f"ModernDropArea {{{styles.DROP_AREA_STYLE}}}")

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 15, 20, 15)

        text = (
            "Перетащите файлы или нажмите для выбора"
            if accept_multiple else
            "Перетащите файл или нажмите для выбора"
        )

        self.label = QLabel(text)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet(styles.DROP_AREA_LABEL)

        self.file_info = QLabel("")
        self.file_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_info.setStyleSheet(styles.DROP_AREA_INFO)
        self.file_info.hide()

        layout.addWidget(self.label)
        layout.addWidget(self.file_info)
        self.setLayout(layout)

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(10)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.open_file_dialog()
        super().mousePressEvent(event)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            valid_files = [u for u in urls if u.toLocalFile().lower().endswith((".xlsx", ".xls"))]
            if valid_files:
                event.acceptProposedAction()
                self.setStyleSheet(f"ModernDropArea {{{styles.DROP_AREA_ACTIVE_STYLE}}}")

    def dragLeaveEvent(self, event):
        self.setStyleSheet(f"ModernDropArea {{{styles.DROP_AREA_STYLE}}}")

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()
                 if u.toLocalFile().lower().endswith((".xlsx", ".xls"))]
        if files:
            if self.accept_multiple:
                self.files_dropped.emit(files)
                self.update_files_info(files)
            else:
                self.file_dropped.emit(files[0])
                self.update_file_info(files[0])
        self.setStyleSheet(f"ModernDropArea {{{styles.DROP_AREA_STYLE}}}")

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
        self.label.hide()
        self.file_info.show()
        self.file_info.setText(f"✓ {os.path.basename(file_path)}")

    def update_files_info(self, files: List[str]):
        self.label.hide()
        self.file_info.show()
        self.file_info.setText(f"✓ Выбрано файлов: {len(files)}")

    def reset(self):
        self.label.show()
        self.file_info.hide()
        self.file_info.setText("")
