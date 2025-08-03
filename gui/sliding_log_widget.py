from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame
)
from PySide6.QtCore import QPropertyAnimation, QEasingCurve

from .widgets import ClickableTextEdit


class SlidingLogWidget(QWidget):
    """Widget displaying a sliding log panel on the right."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedWidth(0)
        self.max_width = 300
        self.min_width = 0
        self.is_expanded = False
        self.has_been_shown = False

        self.init_ui()
        self.init_animation()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 0, 0, 0)
        layout.setSpacing(8)

        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(8, 8, 8, 8)

        self.status_dot = QLabel("🟢")
        self.status_dot.setFixedSize(16, 16)

        title_label = QLabel("Лог")
        title_label.setStyleSheet("font-weight: bold; color: #333; font-size: 12px;")

        self.toggle_btn = QPushButton("🔽")
        self.toggle_btn.setFixedSize(20, 20)
        self.toggle_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
                color: #666;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                color: #333;
                background: #f0f0f0;
                border-radius: 10px;
            }
        """)
        self.toggle_btn.clicked.connect(self.toggle_visibility)

        header_layout.addWidget(self.status_dot)
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(self.toggle_btn)

        header_frame = QFrame()
        header_frame.setLayout(header_layout)
        header_frame.setStyleSheet("""
            QFrame {
                background: #f5f5f5;
                border-bottom: 1px solid #ddd;
            }
        """)

        self.log_text = ClickableTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background: white;
                border: none;
                color: #333;
                font-family: Consolas, monospace;
                font-size: 11px;
                padding: 8px;
            }
            QScrollBar:vertical {
                background: #f0f0f0;
                width: 6px;
                border-radius: 3px;
            }
            QScrollBar::handle:vertical {
                background: #ccc;
                border-radius: 3px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #999;
            }
        """)

        layout.addWidget(header_frame)
        layout.addWidget(self.log_text, 1)

        self.setStyleSheet("""
            SlidingLogWidget {
                background: white;
                border: 1px solid #ddd;
                border-right: none;
            }
        """)

    def init_animation(self):
        self.animation = QPropertyAnimation(self, b"maximumWidth")
        self.animation.setDuration(300)
        self.animation.setEasingCurve(QEasingCurve.Type.OutCubic)

    def slide_right(self):
        if self.is_expanded:
            return
        self.is_expanded = True
        self.animation.setStartValue(self.min_width)
        self.animation.setEndValue(self.max_width)
        self.animation.start()

    def slide_left(self):
        if not self.is_expanded:
            return
        self.is_expanded = False
        self.animation.setStartValue(self.max_width)
        self.animation.setEndValue(self.min_width)
        self.animation.start()

    def toggle_visibility(self):
        if self.is_expanded:
            self.slide_left()
            self.toggle_btn.setText("🔼")
            self.toggle_btn.setToolTip("Показать лог")
        else:
            self.slide_right()
            self.toggle_btn.setText("🔽")
            self.toggle_btn.setToolTip("Скрыть лог")

        if hasattr(self.parent(), "sync_toggle_button"):
            self.parent().sync_toggle_button()

    def add_log_message(self, message: str, message_type: str = "info"):
        if message_type == "error":
            self.status_dot.setText("🔴")
        elif message_type == "warning":
            self.status_dot.setText("🟡")
        elif message_type == "success":
            self.status_dot.setText("🟢")
        else:
            self.status_dot.setText("🔵")

        self.log_text.append(message)
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

        if not self.has_been_shown:
            self.has_been_shown = True
            self.slide_right()
            if hasattr(self.parent(), "on_log_first_shown"):
                self.parent().on_log_first_shown()

    def clear_log(self):
        self.log_text.clear()
        self.status_dot.setText("🟢")
