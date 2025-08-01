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


class BatchMappingDialog(QDialog):
    """Улучшенный диалог настройки маппинга для пакетной обработки."""

    def __init__(self, excel_files: List[str], google_sheets: List[str], parent=None):
        super().__init__(parent)
        self.excel_files = excel_files
        self.google_sheets = google_sheets
        self.mappings = []

        self.setWindowTitle("Настройка пакетного маппинга")
        self.setModal(True)
        self.resize(900, 650)
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
            }
            QLabel {
                color: #212529;
            }
            QPushButton {
                background-color: #0066cc;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: 500;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #0052a3;
            }
            QPushButton:pressed {
                background-color: #004080;
            }
            QPushButton:disabled {
                background-color: #e9ecef;
                color: #6c757d;
            }
            QGroupBox {
                font-weight: 600;
                color: #495057;
                border: 2px solid #e9ecef;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
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
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        header_layout = QVBoxLayout()

        title = QLabel("🔗 Настройка маппинга файлов")
        title.setStyleSheet("font-size: 18px; font-weight: 600; color: #0066cc; margin-bottom: 5px;")

        instruction = QLabel(
            "Для каждого Excel файла настройте:\n"
            "• Какой лист из Excel копировать\n"
            "• В какой лист Google Таблицы вставлять\n"
            "• Какие колонки копировать (формат: A,B,C → D,E,F)\n"
            "• С какой строки начинать"
        )
        instruction.setStyleSheet("color: #6c757d; line-height: 1.4;")

        header_layout.addWidget(title)
        header_layout.addWidget(instruction)
        header_frame.setLayout(header_layout)
        layout.addWidget(header_frame)

        # Скроллируемая область для файлов
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout()
        scroll_layout.setSpacing(15)

        # Создаем карточку для каждого файла
        self.file_widgets = []
        for i, excel_file in enumerate(self.excel_files):
            file_widget = self.create_file_mapping_widget(excel_file, i)
            self.file_widgets.append(file_widget)
            scroll_layout.addWidget(file_widget)

        scroll_layout.addStretch()
        scroll_widget.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)

        # Быстрые действия
        quick_actions_group = QGroupBox("⚡ Быстрые действия")
        quick_layout = QHBoxLayout()

        select_all_btn = QPushButton("Выбрать все Google листы")
        select_all_btn.clicked.connect(self.select_all_sheets)
        select_all_btn.setStyleSheet("QPushButton { background-color: #28a745; }")

        auto_map_btn = QPushButton("Авто-маппинг по именам")
        auto_map_btn.clicked.connect(self.auto_map_by_names)
        auto_map_btn.setStyleSheet("QPushButton { background-color: #17a2b8; }")

        reset_btn = QPushButton("Сбросить все")
        reset_btn.clicked.connect(self.reset_all_mappings)
        reset_btn.setStyleSheet("QPushButton { background-color: #6c757d; }")

        quick_layout.addWidget(select_all_btn)
        quick_layout.addWidget(auto_map_btn)
        quick_layout.addWidget(reset_btn)
        quick_layout.addStretch()

        quick_actions_group.setLayout(quick_layout)
        layout.addWidget(quick_actions_group)

        # Кнопки диалога
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)

        # Стилизуем кнопки
        ok_btn = buttons.button(QDialogButtonBox.StandardButton.Ok)
        ok_btn.setText("✓ Применить настройки")
        ok_btn.setStyleSheet("QPushButton { background-color: #28a745; min-width: 140px; }")

        cancel_btn = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        cancel_btn.setText("✕ Отмена")
        cancel_btn.setStyleSheet("QPushButton { background-color: #dc3545; }")

        layout.addWidget(buttons)
        self.setLayout(layout)

    def create_file_mapping_widget(self, excel_file: str, index: int) -> QGroupBox:
        """Создание виджета настройки для одного файла"""
        file_name = os.path.basename(excel_file)
        group = QGroupBox(f"📄 {file_name}")

        layout = QGridLayout()
        layout.setSpacing(10)

        # Excel лист
        layout.addWidget(QLabel("Excel лист:"), 0, 0)
        excel_sheet_input = QLineEdit("Sheet1")
        excel_sheet_input.setPlaceholderText("Имя листа в Excel файле")
        excel_sheet_input.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-size: 13px;
            }
            QLineEdit:focus { border-color: #0066cc; }
        """)
        layout.addWidget(excel_sheet_input, 0, 1)

        # Стрелка
        arrow_label = QLabel("→")
        arrow_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        arrow_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #0066cc;")
        layout.addWidget(arrow_label, 0, 2)

        # Google лист
        layout.addWidget(QLabel("Google лист:"), 0, 3)
        google_combo = QComboBox()
        google_combo.addItem("-- Не копировать --", "")
        for sheet in self.google_sheets:
            google_combo.addItem(f"📋 {sheet}", sheet)

        google_combo.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-size: 13px;
                min-width: 150px;
            }
            QComboBox:hover { border-color: #90caf9; }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #6c757d;
                margin-right: 5px;
            }
        """)

        # Автоматический выбор похожего листа
        file_name_without_ext = os.path.splitext(file_name)[0]
        for i in range(1, google_combo.count()):
            sheet_name = google_combo.itemData(i)
            if (file_name_without_ext.lower() in sheet_name.lower() or
                    sheet_name.lower() in file_name_without_ext.lower()):
                google_combo.setCurrentIndex(i)
                break

        layout.addWidget(google_combo, 0, 4)

        # Настройки колонок
        columns_frame = QFrame()
        columns_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 1px solid #e9ecef;
                border-radius: 6px;
                padding: 10px;
            }
        """)
        columns_layout = QHBoxLayout()

        columns_layout.addWidget(QLabel("Колонки:"))

        columns_input = QLineEdit("A → A")
        columns_input.setPlaceholderText("Например: A,B,C → D,E,F или A-C → D-F")
        columns_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-family: monospace;
                font-size: 12px;
            }
        """)
        columns_layout.addWidget(columns_input)

        columns_layout.addWidget(QLabel("Начать со строки:"))

        start_row_spin = QSpinBox()
        start_row_spin.setMinimum(1)
        start_row_spin.setMaximum(10000)
        start_row_spin.setValue(1)
        start_row_spin.setStyleSheet("""
            QSpinBox {
                padding: 6px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                min-width: 60px;
            }
        """)
        columns_layout.addWidget(start_row_spin)

        columns_frame.setLayout(columns_layout)
        layout.addWidget(columns_frame, 1, 0, 1, 5)

        # Сохраняем ссылки на виджеты
        group.excel_file = excel_file
        group.excel_sheet_input = excel_sheet_input
        group.google_combo = google_combo
        group.columns_input = columns_input
        group.start_row_spin = start_row_spin

        group.setLayout(layout)
        return group

    def select_all_sheets(self):
        """Выбрать Google лист для всех файлов"""
        for widget in self.file_widgets:
            combo = widget.google_combo
            if combo.count() > 1:
                combo.setCurrentIndex(1)

    def auto_map_by_names(self):
        """Автоматический маппинг по именам файлов"""
        for widget in self.file_widgets:
            file_path = widget.excel_file
            file_name = os.path.splitext(os.path.basename(file_path))[0].lower()
            combo = widget.google_combo

            best_match_index = 0
            best_score = 0

            for i in range(1, combo.count()):
                sheet_name = combo.itemData(i).lower()

                # Точное совпадение
                if file_name == sheet_name:
                    combo.setCurrentIndex(i)
                    break

                # Частичное совпадение
                elif file_name in sheet_name or sheet_name in file_name:
                    # Подсчет общих символов
                    score = len(set(file_name) & set(sheet_name))
                    if score > best_score:
                        best_score = score
                        best_match_index = i
            else:
                if best_match_index > 0:
                    combo.setCurrentIndex(best_match_index)

    def reset_all_mappings(self):
        """Сброс всех настроек"""
        for widget in self.file_widgets:
            widget.google_combo.setCurrentIndex(0)
            widget.excel_sheet_input.setText("Sheet1")
            widget.columns_input.setText("A → A")
            widget.start_row_spin.setValue(1)

    def validate_and_accept(self):
        """Валидация и принятие настроек"""
        self.mappings = []
        errors = []

        for i, widget in enumerate(self.file_widgets):
            if widget.google_combo.currentData() == "":
                continue  # Пропускаем файлы с "Не копировать"

            file_name = os.path.basename(widget.excel_file)

            try:
                source_cols, target_cols = self.parse_column_mapping(widget.columns_input.text())

                self.mappings.append({
                    'excel_path': widget.excel_file,
                    'excel_sheet': widget.excel_sheet_input.text(),
                    'google_sheet': widget.google_combo.currentData(),
                    'column_mapping': {
                        'source': source_cols,
                        'target': target_cols
                    },
                    'start_row': widget.start_row_spin.value()
                })

            except ValueError as e:
                errors.append(f"Файл '{file_name}': {e}")

        if errors:
            from PySide6.QtWidgets import QMessageBox
            error_text = "Найдены ошибки в настройках:\n\n" + "\n".join(errors)
            QMessageBox.warning(self, "Ошибки настройки", error_text)
            return

        if not self.mappings:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.information(
                self,
                "Внимание",
                "Не выбрано ни одного файла для копирования.\n"
                "Выберите Google листы для файлов, которые хотите обработать."
            )
            return

        self.accept()

    def parse_column_mapping(self, text: str) -> Tuple[List[str], List[str]]:
        """Парсинг маппинга колонок"""
        if '→' not in text:
            raise ValueError("Используйте формат: 'A,B,C → D,E,F' или 'A-C → D-F'")

        parts = text.split('→')
        if len(parts) != 2:
            raise ValueError("Используйте формат: 'A,B,C → D,E,F'")

        source_part = parts[0].strip()
        target_part = parts[1].strip()

        source_cols = self.parse_column_range(source_part)
        target_cols = self.parse_column_range(target_part)

        if len(source_cols) != len(target_cols):
            raise ValueError(f"Количество колонок должно совпадать: {len(source_cols)} ≠ {len(target_cols)}")

        return source_cols, target_cols

    def parse_column_range(self, text: str) -> List[str]:
        """Парсинг диапазона колонок"""
        text = text.strip()

        # Диапазон вида A-C
        if '-' in text and ',' not in text:
            parts = text.split('-')
            if len(parts) != 2:
                raise ValueError(f"Неверный диапазон: {text}")

            start_col = parts[0].strip().upper()
            end_col = parts[1].strip().upper()

            if not start_col.isalpha() or not end_col.isalpha() or len(start_col) != 1 or len(end_col) != 1:
                raise ValueError(f"Неверные колонки в диапазоне: {text}")

            start_ord = ord(start_col)
            end_ord = ord(end_col)

            if start_ord > end_ord:
                raise ValueError(f"Неверный порядок в диапазоне: {text}")

            return [chr(i) for i in range(start_ord, end_ord + 1)]

        # Список колонок вида A,B,C
        else:
            cols = [col.strip().upper() for col in text.split(',') if col.strip()]
            if not cols:
                raise ValueError("Не указаны колонки")

            for col in cols:
                if not col.isalpha() or len(col) != 1:
                    raise ValueError(f"Неверная колонка: {col}")

            return cols


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