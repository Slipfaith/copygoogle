"""Centralized GUI style definitions."""

WINDOW_STYLE = """
    QWidget {
        background-color: #ffffff;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }
"""

DROP_AREA_STYLE = """
    background-color: #f8f9fa;
    border: 2px dashed #dee2e6;
    border-radius: 8px;
"""

DROP_AREA_ACTIVE_STYLE = """
    background-color: #e7f3ff;
    border: 2px solid #0066cc;
    border-radius: 8px;
"""

def primary_button(color: str, hover_color: str) -> str:
    return f"""
        QPushButton {{
            background-color: {color};
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 500;
        }}
        QPushButton:hover {{
            background-color: {hover_color};
        }}
        QPushButton:pressed {{
            background-color: {hover_color};
        }}
        QPushButton:disabled {{
            background-color: #e9ecef;
            color: #adb5bd;
        }}
    """

def secondary_button(color: str) -> str:
    return f"""
        QPushButton {{
            background-color: white;
            color: {color};
            border: 1px solid {color};
            padding: 10px 20px;
            border-radius: 6px;
            font-size: 14px;
        }}
        QPushButton:hover {{
            background-color: {color};
            color: white;
        }}
        QPushButton:disabled {{
            border-color: #dee2e6;
            color: #adb5bd;
        }}
    """

SMALL_BUTTON_STYLE = """
    QPushButton {
        background-color: transparent;
        color: #6c757d;
        border: none;
        padding: 5px 10px;
        font-size: 12px;
    }
    QPushButton:hover {
        color: #495057;
        background-color: #f8f9fa;
        border-radius: 4px;
    }
"""

DROP_AREA_LABEL = """
    color: #6c757d;
    font-size: 13px;
"""

DROP_AREA_INFO = """
    color: #28a745;
    font-size: 12px;
    font-weight: 500;
"""

TITLE_LABEL_STYLE = """
    font-size: 24px;
    font-weight: 600;
    color: #212529;
"""

SUBTITLE_LABEL_STYLE = """
    font-size: 14px;
    color: #6c757d;
"""

URL_CONTAINER_STYLE = """
    QWidget {
        background-color: #f8f9fa;
        border-radius: 8px;
    }
"""

URL_LABEL_STYLE = """
    font-size: 12px;
    font-weight: 500;
    color: #495057;
    margin-bottom: 5px;
"""

URL_INPUT_STYLE = """
    QLineEdit {
        padding: 10px;
        border: 1px solid #ced4da;
        border-radius: 6px;
        font-size: 14px;
        background-color: white;
    }
    QLineEdit:focus {
        border-color: #0066cc;
        outline: none;
    }
"""

DOWNLOAD_BUTTON_STYLE = """
    QPushButton {
        background-color: #17a2b8;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 4px;
        font-weight: 500;
    }
    QPushButton:hover {
        background-color: #138496;
    }
    QPushButton:disabled {
        background-color: #e9ecef;
        color: #6c757d;
    }
"""

BACKUP_CHECKBOX_STYLE = """
    QCheckBox {
        color: #495057;
        font-size: 13px;
    }
    QCheckBox::indicator {
        width: 18px;
        height: 18px;
    }
"""

TAB_WIDGET_STYLE = """
    QTabWidget::pane {
        border: none;
        background-color: white;
    }
    QTabBar::tab {
        padding: 8px 20px;
        margin: 0 2px;
        background-color: #f8f9fa;
        border: none;
        border-radius: 6px 6px 0 0;
    }
    QTabBar::tab:selected {
        background-color: white;
        color: #0066cc;
        font-weight: 500;
    }
    QTabBar::tab:hover:!selected {
        background-color: #e9ecef;
    }
"""

FILES_LIST_STYLE = """
    QListWidget {
        border: 1px solid #dee2e6;
        border-radius: 6px;
        background-color: #f8f9fa;
        padding: 5px;
    }
    QListWidget::item {
        padding: 3px;
        border-radius: 3px;
    }
    QListWidget::item:selected {
        background-color: #e7f3ff;
        color: #0066cc;
    }
"""

PROGRESS_BAR_STYLE = """
    QProgressBar {
        border: none;
        border-radius: 6px;
        background-color: #e9ecef;
        text-align: center;
        height: 20px;
    }
    QProgressBar::chunk {
        background-color: #0066cc;
        border-radius: 6px;
    }
"""

STATUS_LABEL_STYLE = """
    color: #6c757d;
    font-size: 12px;
"""

LOG_TEXT_STYLE = """
    QTextEdit {
        border: 1px solid #dee2e6;
        border-radius: 6px;
        background-color: #f8f9fa;
        padding: 8px;
        font-family: 'SF Mono', Monaco, monospace;
        font-size: 11px;
        color: #495057;
    }
"""
