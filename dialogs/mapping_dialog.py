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

class MappingDialog(QDialog):
    """Улучшенный диалог настройки маппинга для одного файла."""

    def __init__(self, excel_sheets: List[str], google_sheets: List[str], parent=None):
        super().__init__(parent)
        self.excel_sheets = excel_sheets
        self.google_sheets = google_sheets

        self.setWindowTitle("Настройка маппинга")
        self.setModal(True)
        self.resize(700, 500)
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
            }
            QLabel {
                color: #212529;
            }
            QLineEdit, QSpinBox {
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 8px;
                background-color: white;
                font-size: 13px;
            }
            QLineEdit:focus, QSpinBox:focus {
                border-color: #0066cc;
                outline: none;
            }
            QPushButton {
                background-color: #0066cc;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #0052a3;
            }
            QGroupBox {
                font-weight: 600;
                color: #495057;
                border: 2px solid #e9ecef;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 12px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 8px 0 8px;
                background-color: white;
            }
        """)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(20)

        # Заголовок с инструкцией
        header_frame = QFrame()
        header_frame.setStyleSheet("""
            QFrame {
                background-color: #e3f2fd;
                border: 1px solid #90caf9;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        header_layout = QVBoxLayout()

        title = QLabel("🔗 Настройка маппинга данных")
        title.setStyleSheet("font-size: 16px; font-weight: 600; color: #0066cc; margin-bottom: 5px;")

        instruction = QLabel(
            "Настройте соответствие между листами Excel и Google Таблиц,\n"
            "а также укажите какие колонки копировать."
        )
        instruction.setStyleSheet("color: #1565c0; line-height: 1.3;")

        header_layout.addWidget(title)
        header_layout.addWidget(instruction)
        header_frame.setLayout(header_layout)
        layout.addWidget(header_frame)

        # Группа маппинга листов
        sheets_group = QGroupBox("📋 Соответствие листов")
        sheets_layout = QVBoxLayout()

        # Таблица маппинга листов
        self.sheet_table = QTableWidget()
        self.sheet_table.setColumnCount(3)
        self.sheet_table.setHorizontalHeaderLabels(["Excel лист", "→", "Google лист"])
        self.sheet_table.horizontalHeader().setStretchLastSection(True)
        self.sheet_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #dee2e6;
                border-radius: 6px;
                background-color: white;
                gridline-color: #e9ecef;
            }
            QTableWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f1f3f4;
            }
            QHeaderView::section {
                background-color: #f8f9fa;
                padding: 10px;
                border: none;
                font-weight: 600;
                color: #495057;
                border-right: 1px solid #dee2e6;
            }
        """)

        # Настройка размеров колонок
        header = self.sheet_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.sheet_table.setColumnWidth(1, 40)

        # Заполнение таблицы
        self.sheet_table.setRowCount(len(self.excel_sheets))
        for i, excel_sheet in enumerate(self.excel_sheets):
            # Excel лист
            excel_item = QTableWidgetItem(f"📄 {excel_sheet}")
            excel_item.setFlags(excel_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.sheet_table.setItem(i, 0, excel_item)

            # Стрелка
            arrow_item = QTableWidgetItem("→")
            arrow_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            arrow_item.setFlags(arrow_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.sheet_table.setItem(i, 1, arrow_item)

            # Google лист
            google_combo = QComboBox()
            google_combo.addItem("-- Не копировать --", "")
            for sheet in self.google_sheets:
                google_combo.addItem(f"📋 {sheet}", sheet)

            google_combo.setStyleSheet("""
                QComboBox {
                    border: 1px solid #e0e0e0;
                    border-radius: 4px;
                    padding: 6px;
                    background-color: white;
                }
                QComboBox:hover {
                    border-color: #90caf9;
                }
            """)

            # Автоматический выбор одинакового листа
            if excel_sheet in self.google_sheets:
                for j in range(google_combo.count()):
                    if google_combo.itemData(j) == excel_sheet:
                        google_combo.setCurrentIndex(j)
                        break

            self.sheet_table.setCellWidget(i, 2, google_combo)

        sheets_layout.addWidget(self.sheet_table)
        sheets_group.setLayout(sheets_layout)
        layout.addWidget(sheets_group)

        # Группа настройки колонок
        columns_group = QGroupBox("📊 Настройка колонок")
        columns_layout = QGridLayout()
        columns_layout.setSpacing(15)

        # Исходные колонки
        columns_layout.addWidget(QLabel("Колонки из Excel:"), 0, 0)
        self.source_columns = QLineEdit("A")
        self.source_columns.setPlaceholderText("A, B, C или A-C")
        columns_layout.addWidget(self.source_columns, 0, 1)

        # Стрелка
        arrow_label = QLabel("→")
        arrow_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        arrow_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #0066cc;")
        columns_layout.addWidget(arrow_label, 0, 2)

        # Целевые колонки
        columns_layout.addWidget(QLabel("Колонки в Google:"), 0, 3)
        self.target_columns = QLineEdit("A")
        self.target_columns.setPlaceholderText("A, B, C или A-C")
        columns_layout.addWidget(self.target_columns, 0, 4)

        # Начальная строка
        columns_layout.addWidget(QLabel("Начать с строки:"), 1, 0)
        self.start_row = QSpinBox()
        self.start_row.setMinimum(1)
        self.start_row.setMaximum(10000)
        self.start_row.setValue(1)
        self.start_row.setSuffix(" строка")
        columns_layout.addWidget(self.start_row, 1, 1)

        # Подсказка по формату
        hint_frame = QFrame()
        hint_frame.setStyleSheet("""
            QFrame {
                background-color: #fff3cd;
                border: 1px solid #ffeaa7;
                border-radius: 6px;
                padding: 10px;
            }
        """)
        hint_layout = QVBoxLayout()

        hint_title = QLabel("💡 Примеры форматов колонок:")
        hint_title.setStyleSheet("font-weight: 600; color: #856404;")

        hint_text = QLabel(
            "• Отдельные колонки: A, C, E\n"
            "• Диапазон колонок: A-E (от A до E)\n"
            "• Количество колонок слева и справа должно совпадать"
        )
        hint_text.setStyleSheet("color: #856404; font-size: 12px; line-height: 1.4;")

        hint_layout.addWidget(hint_title)
        hint_layout.addWidget(hint_text)
        hint_frame.setLayout(hint_layout)
        columns_layout.addWidget(hint_frame, 2, 0, 1, 5)

        columns_group.setLayout(columns_layout)
        layout.addWidget(columns_group)

        # Кнопки
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        # Стилизация кнопок
        ok_btn = buttons.button(QDialogButtonBox.StandardButton.Ok)
        ok_btn.setText("✓ Применить")
        ok_btn.setStyleSheet("QPushButton { background-color: #28a745; min-width: 100px; }")

        cancel_btn = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        cancel_btn.setText("✕ Отмена")
        cancel_btn.setStyleSheet("QPushButton { background-color: #dc3545; }")

        layout.addWidget(buttons)
        self.setLayout(layout)

    def get_config(self) -> dict:
        """Получение конфигурации"""
        sheet_mapping = {}
        for i in range(self.sheet_table.rowCount()):
            excel_sheet = self.sheet_table.item(i, 0).text().replace("📄 ", "")
            google_combo = self.sheet_table.cellWidget(i, 2)
            google_sheet = google_combo.currentData()
            if google_sheet:  # Не пустая строка
                sheet_mapping[excel_sheet] = google_sheet

        source_cols = [col.strip().upper() for col in self.source_columns.text().split(',') if col.strip()]
        target_cols = [col.strip().upper() for col in self.target_columns.text().split(',') if col.strip()]

        return {
            'sheet_mapping': sheet_mapping,
            'column_mapping': {
                'source': source_cols,
                'target': target_cols
            },
            'start_row': self.start_row.value()
        }

