from typing import List
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget, QHeaderView, QComboBox,
    QTableWidgetItem, QSpinBox, QDialogButtonBox, QFrame, QLineEdit, QGroupBox,
    QGridLayout, QScrollArea, QWidget
)
from PySide6.QtCore import Qt
from .. import styles


class MappingDialog(QDialog):
    """Современный диалог настройки маппинга для одного файла."""

    def __init__(self, excel_sheets: List[str], google_sheets: List[str], parent=None):
        super().__init__(parent)
        self.excel_sheets = excel_sheets
        self.google_sheets = google_sheets

        self.setWindowTitle("Настройка маппинга данных")
        self.setModal(True)
        self.setFixedSize(750, 600)
        self.setStyleSheet(styles.DIALOG_STYLE)

        self.init_ui()

    def init_ui(self):
        """Инициализация пользовательского интерфейса"""
        layout = QVBoxLayout(self)
        layout.setSpacing(24)
        layout.setContentsMargins(32, 32, 32, 32)

        # Заголовок с описанием
        self.create_header(layout)

        # Секция маппинга листов
        self.create_sheets_mapping_section(layout)

        # Секция настройки колонок
        self.create_columns_mapping_section(layout)

        # Кнопки диалога
        self.create_dialog_buttons(layout)

    def create_header(self, parent_layout):
        """Создает заголовок диалога"""
        header_frame = QFrame()
        header_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {styles.COLORS['primary_light']};
                border: 1px solid {styles.COLORS['primary']};
                border-radius: {styles.BORDER_RADIUS['lg']};
                padding: {styles.SPACING['xl']};
            }}
        """)

        header_layout = QVBoxLayout(header_frame)
        header_layout.setSpacing(8)

        title = QLabel("🔗 Настройка маппинга данных")
        title.setStyleSheet(f"""
            QLabel {{
                font-size: 18px;
                font-weight: 700;
                color: {styles.COLORS['primary']};
                margin: 0;
            }}
        """)

        description = QLabel(
            "Настройте соответствие между листами Excel и Google Таблиц, "
            "а также укажите какие колонки копировать и с какой строки начинать."
        )
        description.setStyleSheet(f"""
            QLabel {{
                font-size: 14px;
                color: {styles.COLORS['gray_700']};
                line-height: 1.4;
                margin: 0;
            }}
        """)
        description.setWordWrap(True)

        header_layout.addWidget(title)
        header_layout.addWidget(description)

        parent_layout.addWidget(header_frame)

    def create_sheets_mapping_section(self, parent_layout):
        """Создает секцию маппинга листов"""
        sheets_group = QGroupBox("📋 Соответствие листов")
        sheets_group.setStyleSheet(f"""
            QGroupBox {{
                font-size: 16px;
                font-weight: 600;
                color: {styles.COLORS['gray_800']};
                border: 2px solid {styles.COLORS['gray_200']};
                border-radius: {styles.BORDER_RADIUS['lg']};
                margin-top: {styles.SPACING['lg']};
                padding-top: {styles.SPACING['lg']};
                background-color: {styles.COLORS['white']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {styles.SPACING['lg']};
                padding: 0 {styles.SPACING['sm']};
                background-color: {styles.COLORS['white']};
            }}
        """)

        sheets_layout = QVBoxLayout(sheets_group)
        sheets_layout.setSpacing(16)
        sheets_layout.setContentsMargins(20, 20, 20, 20)

        # Создаем таблицу маппинга
        self.sheet_table = QTableWidget()
        self.sheet_table.setColumnCount(3)
        self.sheet_table.setHorizontalHeaderLabels(["Excel лист", "→", "Google лист"])
        self.sheet_table.setStyleSheet(styles.TABLE_STYLE)
        self.sheet_table.setFixedHeight(min(250, len(self.excel_sheets) * 40 + 60))

        # Настройка размеров колонок
        header = self.sheet_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.sheet_table.setColumnWidth(1, 50)

        # Заполнение таблицы
        self.populate_sheets_table()

        sheets_layout.addWidget(self.sheet_table)
        parent_layout.addWidget(sheets_group)

    def populate_sheets_table(self):
        """Заполняет таблицу маппинга листов"""
        self.sheet_table.setRowCount(len(self.excel_sheets))

        for i, excel_sheet in enumerate(self.excel_sheets):
            # Excel лист (неизменяемый)
            excel_item = QTableWidgetItem(f"📄 {excel_sheet}")
            excel_item.setFlags(excel_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            excel_item.setStyleSheet(f"color: {styles.COLORS['gray_800']};")
            self.sheet_table.setItem(i, 0, excel_item)

            # Стрелка (неизменяемая)
            arrow_item = QTableWidgetItem("→")
            arrow_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            arrow_item.setFlags(arrow_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            arrow_item.setStyleSheet(f"color: {styles.COLORS['primary']}; font-weight: bold; font-size: 16px;")
            self.sheet_table.setItem(i, 1, arrow_item)

            # Google лист (выпадающий список)
            google_combo = QComboBox()
            google_combo.addItem("-- Не копировать --", "")

            for sheet in self.google_sheets:
                google_combo.addItem(f"📋 {sheet}", sheet)

            google_combo.setStyleSheet(f"""
                QComboBox {{
                    border: 1px solid {styles.COLORS['gray_300']};
                    border-radius: {styles.BORDER_RADIUS['sm']};
                    padding: {styles.SPACING['sm']};
                    background-color: {styles.COLORS['white']};
                    font-size: 14px;
                }}
                QComboBox:hover {{
                    border-color: {styles.COLORS['primary']};
                }}
            """)

            # Автоматический выбор совпадающего листа
            if excel_sheet in self.google_sheets:
                for j in range(google_combo.count()):
                    if google_combo.itemData(j) == excel_sheet:
                        google_combo.setCurrentIndex(j)
                        break

            self.sheet_table.setCellWidget(i, 2, google_combo)

    def create_columns_mapping_section(self, parent_layout):
        """Создает секцию настройки колонок"""
        columns_group = QGroupBox("📊 Настройка колонок и строк")
        columns_group.setStyleSheet(f"""
            QGroupBox {{
                font-size: 16px;
                font-weight: 600;
                color: {styles.COLORS['gray_800']};
                border: 2px solid {styles.COLORS['gray_200']};
                border-radius: {styles.BORDER_RADIUS['lg']};
                margin-top: {styles.SPACING['lg']};
                padding-top: {styles.SPACING['lg']};
                background-color: {styles.COLORS['white']};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: {styles.SPACING['lg']};
                padding: 0 {styles.SPACING['sm']};
                background-color: {styles.COLORS['white']};
            }}
        """)

        columns_layout = QGridLayout(columns_group)
        columns_layout.setSpacing(16)
        columns_layout.setContentsMargins(20, 20, 20, 20)

        # Исходные колонки
        columns_layout.addWidget(QLabel("Колонки из Excel:"), 0, 0)
        self.source_columns = QLineEdit("A")
        self.source_columns.setPlaceholderText("A, B, C или A-C")
        self.source_columns.setStyleSheet(styles.URL_INPUT_STYLE)
        columns_layout.addWidget(self.source_columns, 0, 1)

        # Стрелка
        arrow_label = QLabel("→")
        arrow_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        arrow_label.setStyleSheet(f"""
            QLabel {{
                font-size: 20px;
                font-weight: bold;
                color: {styles.COLORS['primary']};
            }}
        """)
        columns_layout.addWidget(arrow_label, 0, 2)

        # Целевые колонки
        columns_layout.addWidget(QLabel("Колонки в Google:"), 0, 3)
        self.target_columns = QLineEdit("A")
        self.target_columns.setPlaceholderText("A, B, C или A-C")
        self.target_columns.setStyleSheet(styles.URL_INPUT_STYLE)
        columns_layout.addWidget(self.target_columns, 0, 4)

        # Начальная строка
        columns_layout.addWidget(QLabel("Начать с строки:"), 1, 0)
        self.start_row = QSpinBox()
        self.start_row.setMinimum(1)
        self.start_row.setMaximum(10000)
        self.start_row.setValue(1)
        self.start_row.setSuffix(" строка")
        self.start_row.setStyleSheet(f"""
            QSpinBox {{
                border: 2px solid {styles.COLORS['gray_200']};
                border-radius: {styles.BORDER_RADIUS['md']};
                padding: {styles.SPACING['md']} {styles.SPACING['lg']};
                background-color: {styles.COLORS['white']};
                font-size: 14px;
                min-height: 20px;
            }}
            QSpinBox:focus {{
                border-color: {styles.COLORS['primary']};
            }}
        """)
        columns_layout.addWidget(self.start_row, 1, 1)

        # Подсказка
        self.create_hint_section(columns_layout)

        parent_layout.addWidget(columns_group)

    def create_hint_section(self, parent_layout):
        """Создает секцию с подсказками"""
        hint_frame = QFrame()
        hint_frame.setStyleSheet(f"""
            QFrame {{
                background-color: {styles.COLORS['warning']};
                background-color: #fef3c7;
                border: 1px solid #f59e0b;
                border-radius: {styles.BORDER_RADIUS['md']};
                padding: {styles.SPACING['lg']};
            }}
        """)

        hint_layout = QVBoxLayout(hint_frame)
        hint_layout.setSpacing(8)

        hint_title = QLabel("💡 Форматы колонок:")
        hint_title.setStyleSheet(f"""
            QLabel {{
                font-weight: 600;
                color: #92400e;
                font-size: 14px;
                margin: 0;
            }}
        """)

        hint_text = QLabel(
            "• Отдельные колонки: A, C, E\n"
            "• Диапазон колонок: A-E (от A до E включительно)\n"
            "• Количество исходных и целевых колонок должно совпадать\n"
            "• Пример: A,B,C → D,E,F или A-C → D-F"
        )
        hint_text.setStyleSheet(f"""
            QLabel {{
                color: #92400e;
                font-size: 13px;
                line-height: 1.4;
                margin: 0;
            }}
        """)

        hint_layout.addWidget(hint_title)
        hint_layout.addWidget(hint_text)

        parent_layout.addWidget(hint_frame, 2, 0, 1, 5)

    def create_dialog_buttons(self, parent_layout):
        """Создает кнопки диалога"""
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        # Стилизация кнопок
        ok_btn = buttons.button(QDialogButtonBox.StandardButton.Ok)
        ok_btn.setText("✅ Применить настройки")
        ok_btn.setStyleSheet(styles.success_button())

        cancel_btn = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        cancel_btn.setText("❌ Отмена")
        cancel_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {styles.COLORS['white']};
                color: {styles.COLORS['danger']};
                border: 2px solid {styles.COLORS['danger']};
                padding: {styles.SPACING['md']} {styles.SPACING['xxl']};
                border-radius: {styles.BORDER_RADIUS['md']};
                font-size: 14px;
                font-weight: 600;
                min-height: 20px;
                min-width: 120px;
            }}
            QPushButton:hover {{
                background-color: {styles.COLORS['danger']};
                color: {styles.COLORS['white']};
            }}
        """)

        parent_layout.addWidget(buttons)

    def get_config(self) -> dict:
        """Получает конфигурацию маппинга"""
        # Собираем маппинг листов
        sheet_mapping = {}
        for i in range(self.sheet_table.rowCount()):
            excel_sheet = self.sheet_table.item(i, 0).text().replace("📄 ", "")
            google_combo = self.sheet_table.cellWidget(i, 2)
            google_sheet = google_combo.currentData()

            if google_sheet:  # Если выбран Google лист
                sheet_mapping[excel_sheet] = google_sheet

        # Парсим колонки
        source_cols = self.parse_columns(self.source_columns.text())
        target_cols = self.parse_columns(self.target_columns.text())

        return {
            'sheet_mapping': sheet_mapping,
            'column_mapping': {
                'source': source_cols,
                'target': target_cols
            },
            'start_row': self.start_row.value()
        }

    def parse_columns(self, text: str) -> List[str]:
        """Парсит строку колонок в список"""
        text = text.strip().upper()
        if not text:
            return ['A']

        # Обработка диапазона (A-C)
        if '-' in text and ',' not in text:
            parts = text.split('-')
            if len(parts) == 2:
                start_col = parts[0].strip()
                end_col = parts[1].strip()

                if start_col.isalpha() and end_col.isalpha() and len(start_col) == 1 and len(end_col) == 1:
                    start_ord = ord(start_col)
                    end_ord = ord(end_col)
                    return [chr(i) for i in range(start_ord, end_ord + 1)]

        # Обработка списка (A,B,C)
        cols = [col.strip() for col in text.split(',') if col.strip()]

        # Фильтруем только валидные колонки
        valid_cols = []
        for col in cols:
            if col.isalpha() and len(col) == 1:
                valid_cols.append(col)

        return valid_cols if valid_cols else ['A']