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

