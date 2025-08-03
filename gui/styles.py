"""Улучшенные стили с консистентным дизайном."""

# Цветовая схема
COLORS = {
    'primary': '#2563eb',      # Синий
    'primary_hover': '#1d4ed8',
    'primary_light': '#dbeafe',
    'secondary': '#64748b',    # Серый
    'secondary_hover': '#475569',
    'success': '#059669',      # Зеленый
    'success_hover': '#047857',
    'warning': '#d97706',      # Оранжевый
    'danger': '#dc2626',       # Красный
    'danger_hover': '#b91c1c',

    # Нейтральные цвета
    'white': '#ffffff',
    'gray_50': '#f9fafb',
    'gray_100': '#f3f4f6',
    'gray_200': '#e5e7eb',
    'gray_300': '#d1d5db',
    'gray_400': '#9ca3af',
    'gray_500': '#6b7280',
    'gray_600': '#4b5563',
    'gray_700': '#374151',
    'gray_800': '#1f2937',
    'gray_900': '#111827',
}

# Размеры и отступы
SPACING = {
    'xs': '4px',
    'sm': '8px',
    'md': '12px',
    'lg': '16px',
    'xl': '20px',
    'xxl': '24px',
    'xxxl': '32px',
}

BORDER_RADIUS = {
    'sm': '4px',
    'md': '6px',
    'lg': '8px',
    'xl': '12px',
}

# Базовые стили для окна
WINDOW_STYLE = f"""
QMainWindow {{
    background-color: {COLORS['white']};
    color: {COLORS['gray_800']};
}}

QWidget {{
    background-color: {COLORS['white']};
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
    font-size: 14px;
    color: {COLORS['gray_800']};
}}

* {{
    outline: none;
}}
"""

# Заголовки
TITLE_LABEL_STYLE = f"""
QLabel {{
    font-size: 28px;
    font-weight: 700;
    color: {COLORS['gray_900']};
    margin: 0;
    padding: 0;
}}
"""

SUBTITLE_LABEL_STYLE = f"""
QLabel {{
    font-size: 16px;
    font-weight: 400;
    color: {COLORS['gray_500']};
    margin: 0;
    padding: 0;
}}
"""

# Контейнеры и карточки
def card_container() -> str:
    return f"""
QWidget {{
    background-color: {COLORS['white']};
    border: 1px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['lg']};
    padding: {SPACING['xl']};
}}
"""

URL_CONTAINER_STYLE = f"""
QFrame {{
    background-color: {COLORS['gray_50']};
    border: 1px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['lg']};
    margin: 0px;
}}
"""

# Лейблы
URL_LABEL_STYLE = f"""
QLabel {{
    font-size: 14px;
    font-weight: 600;
    color: {COLORS['gray_700']};
    margin-bottom: {SPACING['sm']};
}}
"""

STATUS_LABEL_STYLE = f"""
QLabel {{
    font-size: 13px;
    color: {COLORS['gray_600']};
    margin: {SPACING['sm']} 0;
}}
"""

# Инпуты
URL_INPUT_STYLE = f"""
QLineEdit {{
    padding: {SPACING['md']} {SPACING['lg']};
    border: 2px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['md']};
    font-size: 14px;
    background-color: {COLORS['white']};
    color: {COLORS['gray_800']};
    min-height: 20px;
}}

QLineEdit:focus {{
    border-color: {COLORS['primary']};
    background-color: {COLORS['white']};
}}

QLineEdit:disabled {{
    background-color: {COLORS['gray_100']};
    color: {COLORS['gray_400']};
    border-color: {COLORS['gray_200']};
}}
"""

# Основные кнопки
def primary_button() -> str:
    return f"""
QPushButton {{
    background-color: {COLORS['primary']};
    color: {COLORS['white']};
    border: none;
    padding: {SPACING['md']} {SPACING['xxl']};
    border-radius: {BORDER_RADIUS['md']};
    font-size: 14px;
    font-weight: 600;
    min-height: 20px;
    min-width: 120px;
}}

QPushButton:hover {{
    background-color: {COLORS['primary_hover']};
}}

QPushButton:pressed {{
    background-color: {COLORS['primary_hover']};
    transform: translateY(1px);
}}

QPushButton:disabled {{
    background-color: {COLORS['gray_300']};
    color: {COLORS['gray_500']};
}}
"""

def success_button() -> str:
    return f"""
QPushButton {{
    background-color: {COLORS['success']};
    color: {COLORS['white']};
    border: none;
    padding: {SPACING['md']} {SPACING['xxl']};
    border-radius: {BORDER_RADIUS['md']};
    font-size: 14px;
    font-weight: 600;
    min-height: 20px;
    min-width: 120px;
}}

QPushButton:hover {{
    background-color: {COLORS['success_hover']};
}}

QPushButton:pressed {{
    background-color: {COLORS['success_hover']};
    transform: translateY(1px);
}}

QPushButton:disabled {{
    background-color: {COLORS['gray_300']};
    color: {COLORS['gray_500']};
}}
"""

def secondary_button() -> str:
    return f"""
QPushButton {{
    background-color: {COLORS['white']};
    color: {COLORS['gray_700']};
    border: 2px solid {COLORS['gray_300']};
    padding: {SPACING['md']} {SPACING['xl']};
    border-radius: {BORDER_RADIUS['md']};
    font-size: 14px;
    font-weight: 500;
    min-height: 20px;
    min-width: 100px;
}}

QPushButton:hover {{
    background-color: {COLORS['gray_50']};
    border-color: {COLORS['gray_400']};
}}

QPushButton:pressed {{
    background-color: {COLORS['gray_100']};
    transform: translateY(1px);
}}

QPushButton:disabled {{
    background-color: {COLORS['gray_100']};
    color: {COLORS['gray_400']};
    border-color: {COLORS['gray_200']};
}}
"""

def download_button() -> str:
    return f"""
QPushButton {{
    background-color: {COLORS['warning']};
    color: {COLORS['white']};
    border: none;
    border-radius: {BORDER_RADIUS['md']};
    font-size: 18px;
    font-weight: normal;
    text-align: center;
}}

QPushButton:hover {{
    background-color: #b45309;
    transform: scale(1.1);
}}

QPushButton:pressed {{
    transform: scale(0.95);
}}

QPushButton:disabled {{
    background-color: {COLORS['gray_300']};
    color: {COLORS['gray_500']};
    transform: none;
}}
"""

def small_button() -> str:
    return f"""
QPushButton {{
    background-color: transparent;
    color: {COLORS['gray_600']};
    border: 1px solid transparent;
    padding: {SPACING['sm']} {SPACING['md']};
    border-radius: {BORDER_RADIUS['sm']};
    font-size: 13px;
    font-weight: 500;
    min-height: 16px;
}}

QPushButton:hover {{
    background-color: {COLORS['gray_100']};
    color: {COLORS['gray_800']};
}}

QPushButton:pressed {{
    background-color: {COLORS['gray_200']};
}}
"""

# Дроп-области
DROP_AREA_STYLE = f"""
QWidget {{
    background-color: {COLORS['gray_50']};
    border: 2px dashed {COLORS['gray_300']};
    border-radius: {BORDER_RADIUS['lg']};
    padding: {SPACING['xl']};
}}
"""

DROP_AREA_ACTIVE_STYLE = f"""
QWidget {{
    background-color: {COLORS['primary_light']};
    border: 2px dashed {COLORS['primary']};
    border-radius: {BORDER_RADIUS['lg']};
    padding: {SPACING['xl']};
}}
"""

DROP_AREA_LABEL = f"""
QLabel {{
    color: {COLORS['gray_600']};
    font-size: 14px;
    font-weight: 500;
    text-align: center;
}}
"""

DROP_AREA_INFO = f"""
QLabel {{
    color: {COLORS['success']};
    font-size: 13px;
    font-weight: 600;
    text-align: center;
}}
"""

# Табы
TAB_WIDGET_STYLE = f"""
QTabWidget::pane {{
    border: 1px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['lg']};
    background-color: {COLORS['white']};
    padding: {SPACING['xl']};
    margin-top: -1px;
}}

QTabBar::tab {{
    background-color: {COLORS['gray_100']};
    border: 1px solid {COLORS['gray_200']};
    border-bottom: none;
    padding: {SPACING['md']} {SPACING['xxl']};
    margin-right: 2px;
    border-radius: {BORDER_RADIUS['md']} {BORDER_RADIUS['md']} 0 0;
    font-weight: 500;
    color: {COLORS['gray_600']};
    min-width: 80px;
}}

QTabBar::tab:selected {{
    background-color: {COLORS['white']};
    color: {COLORS['primary']};
    font-weight: 600;
    border-bottom: 1px solid {COLORS['white']};
}}

QTabBar::tab:hover:!selected {{
    background-color: {COLORS['gray_200']};
    color: {COLORS['gray_800']};
}}
"""

# Списки
FILES_LIST_STYLE = f"""
QListWidget {{
    border: 1px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['md']};
    background-color: {COLORS['white']};
    padding: {SPACING['sm']};
    font-size: 13px;
}}

QListWidget::item {{
    padding: {SPACING['sm']} {SPACING['md']};
    margin: 2px 0;
    border-radius: {BORDER_RADIUS['sm']};
    color: {COLORS['gray_700']};
}}

QListWidget::item:selected {{
    background-color: {COLORS['primary_light']};
    color: {COLORS['primary']};
    font-weight: 500;
}}

QListWidget::item:hover {{
    background-color: {COLORS['gray_100']};
}}
"""

# Прогресс-бар
PROGRESS_BAR_STYLE = f"""
QProgressBar {{
    border: none;
    border-radius: {BORDER_RADIUS['md']};
    background-color: {COLORS['gray_200']};
    text-align: center;
    font-weight: 600;
    color: {COLORS['gray_800']};
    height: 24px;
}}

QProgressBar::chunk {{
    background-color: {COLORS['primary']};
    border-radius: {BORDER_RADIUS['md']};
}}
"""

# Лог
LOG_TEXT_STYLE = f"""
QTextEdit {{
    border: 1px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['md']};
    background-color: {COLORS['gray_50']};
    padding: {SPACING['md']};
    font-family: 'SF Mono', 'Monaco', 'Cascadia Code', 'Roboto Mono', monospace;
    font-size: 12px;
    color: {COLORS['gray_700']};
    line-height: 1.4;
}}

QTextEdit QScrollBar:vertical {{
    border: none;
    background-color: {COLORS['gray_100']};
    width: 8px;
    border-radius: 4px;
}}

QTextEdit QScrollBar::handle:vertical {{
    background-color: {COLORS['gray_400']};
    border-radius: 4px;
    min-height: 20px;
}}

QTextEdit QScrollBar::handle:vertical:hover {{
    background-color: {COLORS['gray_500']};
}}
"""

# Диалоги и формы
DIALOG_STYLE = f"""
QDialog {{
    background-color: {COLORS['white']};
    border: 1px solid {COLORS['gray_300']};
    border-radius: {BORDER_RADIUS['xl']};
}}

QGroupBox {{
    font-weight: 600;
    color: {COLORS['gray_800']};
    border: 2px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['lg']};
    margin-top: {SPACING['lg']};
    padding-top: {SPACING['md']};
    background-color: {COLORS['white']};
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    left: {SPACING['lg']};
    padding: 0 {SPACING['sm']};
    background-color: {COLORS['white']};
    color: {COLORS['gray_800']};
}}

QComboBox {{
    border: 2px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['md']};
    padding: {SPACING['sm']} {SPACING['md']};
    background-color: {COLORS['white']};
    color: {COLORS['gray_800']};
    font-size: 14px;
    min-height: 20px;
}}

QComboBox:hover {{
    border-color: {COLORS['gray_400']};
}}

QComboBox:focus {{
    border-color: {COLORS['primary']};
}}

QComboBox::drop-down {{
    border: none;
    width: 20px;
}}

QComboBox::down-arrow {{
    image: none;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 4px solid {COLORS['gray_600']};
    margin-right: 4px;
}}

QSpinBox {{
    border: 2px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['md']};
    padding: {SPACING['sm']} {SPACING['md']};
    background-color: {COLORS['white']};
    color: {COLORS['gray_800']};
    font-size: 14px;
    min-height: 20px;
}}

QSpinBox:hover {{
    border-color: {COLORS['gray_400']};
}}

QSpinBox:focus {{
    border-color: {COLORS['primary']};
}}
"""

# Таблицы
TABLE_STYLE = f"""
QTableWidget {{
    border: 1px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['md']};
    background-color: {COLORS['white']};
    gridline-color: {COLORS['gray_200']};
    selection-background-color: {COLORS['primary_light']};
}}

QTableWidget::item {{
    padding: {SPACING['md']};
    border-bottom: 1px solid {COLORS['gray_100']};
    color: {COLORS['gray_800']};
}}

QTableWidget::item:selected {{
    background-color: {COLORS['primary_light']};
    color: {COLORS['primary']};
}}

QHeaderView::section {{
    background-color: {COLORS['gray_100']};
    padding: {SPACING['md']};
    border: none;
    border-right: 1px solid {COLORS['gray_200']};
    border-bottom: 1px solid {COLORS['gray_200']};
    font-weight: 600;
    color: {COLORS['gray_800']};
    text-align: left;
}}
"""

# Кастомные виджеты
FRAME_STYLE = f"""
QFrame {{
    background-color: {COLORS['gray_50']};
    border: 1px solid {COLORS['gray_200']};
    border-radius: {BORDER_RADIUS['md']};
    padding: {SPACING['lg']};
}}
"""