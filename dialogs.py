import os
from typing import List, Tuple

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QTableWidget, QHeaderView, QComboBox,
    QTableWidgetItem, QSpinBox, QDialogButtonBox, QPushButton, QFrame,
    QFileDialog, QLineEdit
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont


class BatchMappingDialog(QDialog):
    """Диалог настройки маппинга для пакетной обработки."""

    def __init__(self, excel_files: List[str], google_sheets: List[str], parent=None):
        super().__init__(parent)
        self.excel_files = excel_files
        self.google_sheets = google_sheets
        self.mappings = []

        self.setWindowTitle("Настройка пакетного маппинга")
        self.setModal(True)
        self.setMinimumSize(800, 600)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        info = QLabel("Настройте маппинг для каждого Excel файла:")
        info.setStyleSheet("font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(info)

        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(6)
        self.mapping_table.setHorizontalHeaderLabels([
            "Excel файл", "Excel лист", "→", "Google лист",
            "Колонки (из → в)", "Начальная строка"
        ])

        header = self.mapping_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)

        self.mapping_table.setColumnWidth(1, 100)
        self.mapping_table.setColumnWidth(2, 30)
        self.mapping_table.setColumnWidth(3, 150)
        self.mapping_table.setColumnWidth(4, 150)
        self.mapping_table.setColumnWidth(5, 120)

        self.mapping_table.setRowCount(len(self.excel_files))
        for i, excel_file in enumerate(self.excel_files):
            file_item = QTableWidgetItem(os.path.basename(excel_file))
            file_item.setData(Qt.ItemDataRole.UserRole, excel_file)
            file_item.setFlags(file_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.mapping_table.setItem(i, 0, file_item)

            sheet_item = QTableWidgetItem("Sheet1")
            self.mapping_table.setItem(i, 1, sheet_item)

            arrow_item = QTableWidgetItem("→")
            arrow_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            arrow_item.setFlags(arrow_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.mapping_table.setItem(i, 2, arrow_item)

            google_combo = QComboBox()
            google_combo.addItem("-- Не копировать --")
            google_combo.addItems(self.google_sheets)

            file_name_without_ext = os.path.splitext(os.path.basename(excel_file))[0]
            for sheet in self.google_sheets:
                if file_name_without_ext.lower() in sheet.lower() or sheet.lower() in file_name_without_ext.lower():
                    google_combo.setCurrentText(sheet)
                    break
            self.mapping_table.setCellWidget(i, 3, google_combo)

            columns_item = QTableWidgetItem("A → A")
            self.mapping_table.setItem(i, 4, columns_item)

            start_row_spin = QSpinBox()
            start_row_spin.setMinimum(1)
            start_row_spin.setMaximum(10000)
            start_row_spin.setValue(1)
            self.mapping_table.setCellWidget(i, 5, start_row_spin)

        layout.addWidget(self.mapping_table)

        hint = QLabel("Формат колонок: 'A,B,C → D,E,F' или 'A-C → D-F'")
        hint.setStyleSheet("color: #666; font-style: italic; margin-top: 5px;")
        layout.addWidget(hint)

        quick_actions = QVBoxLayout()

        select_all_btn = QPushButton("Выбрать все")
        select_all_btn.clicked.connect(self.select_all_sheets)
        quick_actions.addWidget(select_all_btn)

        auto_map_btn = QPushButton("Авто-маппинг по именам")
        auto_map_btn.clicked.connect(self.auto_map_by_names)
        quick_actions.addWidget(auto_map_btn)

        quick_actions.addStretch()
        layout.addLayout(quick_actions)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(buttons)

        self.setLayout(layout)

    def select_all_sheets(self):
        for i in range(self.mapping_table.rowCount()):
            combo = self.mapping_table.cellWidget(i, 3)
            if combo and combo.count() > 1:
                combo.setCurrentIndex(1)

    def auto_map_by_names(self):
        for i in range(self.mapping_table.rowCount()):
            file_path = self.mapping_table.item(i, 0).data(Qt.ItemDataRole.UserRole)
            file_name = os.path.splitext(os.path.basename(file_path))[0].lower()
            combo = self.mapping_table.cellWidget(i, 3)
            if combo:
                best_match = None
                best_score = 0
                for j in range(1, combo.count()):
                    sheet_name = combo.itemText(j).lower()
                    if file_name == sheet_name:
                        best_match = j
                        break
                    elif file_name in sheet_name or sheet_name in file_name:
                        score = len(set(file_name) & set(sheet_name))
                        if score > best_score:
                            best_score = score
                            best_match = j
                if best_match:
                    combo.setCurrentIndex(best_match)

    def validate_and_accept(self):
        self.mappings = []
        for i in range(self.mapping_table.rowCount()):
            google_combo = self.mapping_table.cellWidget(i, 3)
            if google_combo.currentText() == "-- Не копировать --":
                continue

            excel_path = self.mapping_table.item(i, 0).data(Qt.ItemDataRole.UserRole)
            excel_sheet = self.mapping_table.item(i, 1).text()
            google_sheet = google_combo.currentText()
            columns_text = self.mapping_table.item(i, 4).text()
            start_row = self.mapping_table.cellWidget(i, 5).value()

            try:
                source_cols, target_cols = self.parse_column_mapping(columns_text)
            except ValueError as e:
                from PySide6.QtWidgets import QMessageBox
                QMessageBox.warning(self, "Ошибка", f"Неверный формат колонок в строке {i+1}: {e}")
                return

            self.mappings.append({
                'excel_path': excel_path,
                'excel_sheet': excel_sheet,
                'google_sheet': google_sheet,
                'column_mapping': {
                    'source': source_cols,
                    'target': target_cols
                },
                'start_row': start_row
            })

        if not self.mappings:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.warning(self, "Внимание", "Не выбрано ни одного файла для копирования")
            return

        self.accept()

    def parse_column_mapping(self, text: str) -> Tuple[List[str], List[str]]:
        parts = text.split('→')
        if len(parts) != 2:
            raise ValueError("Используйте формат: 'A,B,C → D,E,F'")

        source_part = parts[0].strip()
        target_part = parts[1].strip()
        source_cols = self.parse_column_range(source_part)
        target_cols = self.parse_column_range(target_part)

        if len(source_cols) != len(target_cols):
            raise ValueError("Количество исходных и целевых колонок должно совпадать")

        return source_cols, target_cols

    def parse_column_range(self, text: str) -> List[str]:
        text = text.strip()
        if '-' in text and ',' not in text and text.replace('-', '').isalpha():
            parts = text.split('-')
            if len(parts) != 2:
                raise ValueError(f"Неверный диапазон: {text}")
            start_col = parts[0].strip().upper()
            end_col = parts[1].strip().upper()
            if not start_col.isalpha() or not end_col.isalpha():
                raise ValueError(f"Неверные колонки: {text}")
            start_ord = ord(start_col)
            end_ord = ord(end_col)
            if start_ord > end_ord:
                raise ValueError(f"Неверный диапазон: {text}")
            return [chr(i) for i in range(start_ord, end_ord + 1)]
        else:
            cols = [col.strip() for col in text.split(',') if col.strip()]
            return cols


class MappingDialog(QDialog):
    """Диалог настройки маппинга для одного файла."""

    def __init__(self, excel_sheets: List[str], google_sheets: List[str], parent=None):
        super().__init__(parent)
        self.excel_sheets = excel_sheets
        self.google_sheets = google_sheets
        self.setWindowTitle("Настройка маппинга")
        self.setModal(True)
        self.setMinimumWidth(600)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        sheets_group = QFrame()
        sheets_layout = QVBoxLayout()

        self.sheet_table = QTableWidget()
        self.sheet_table.setColumnCount(2)
        self.sheet_table.setHorizontalHeaderLabels(["Excel лист", "Google лист"])
        self.sheet_table.horizontalHeader().setStretchLastSection(True)

        self.sheet_table.setRowCount(len(self.excel_sheets))
        for i, excel_sheet in enumerate(self.excel_sheets):
            excel_item = QTableWidgetItem(excel_sheet)
            excel_item.setFlags(excel_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.sheet_table.setItem(i, 0, excel_item)

            google_combo = QComboBox()
            google_combo.addItem("-- Не копировать --")
            google_combo.addItems(self.google_sheets)
            if excel_sheet in self.google_sheets:
                google_combo.setCurrentText(excel_sheet)
            self.sheet_table.setCellWidget(i, 1, google_combo)

        sheets_layout.addWidget(self.sheet_table)
        sheets_group.setLayout(sheets_layout)
        layout.addWidget(sheets_group)

        columns_group = QFrame()
        columns_layout = QVBoxLayout()

        columns_info = QLabel("Укажите какие колонки копировать:")
        columns_layout.addWidget(columns_info)

        columns_input_layout = QVBoxLayout()

        columns_input_layout.addWidget(QLabel("Из Excel:"))
        self.source_columns = QLineEdit("A")
        self.source_columns.setPlaceholderText("Например: A, C, E")
        columns_input_layout.addWidget(self.source_columns)

        columns_input_layout.addWidget(QLabel("В Google:"))
        self.target_columns = QLineEdit("A")
        self.target_columns.setPlaceholderText("Например: B, D, F")
        columns_input_layout.addWidget(self.target_columns)

        columns_layout.addLayout(columns_input_layout)

        row_layout = QVBoxLayout()
        row_layout.addWidget(QLabel("Начать с строки:"))
        self.start_row = QSpinBox()
        self.start_row.setMinimum(1)
        self.start_row.setMaximum(10000)
        self.start_row.setValue(1)
        row_layout.addWidget(self.start_row)
        row_layout.addStretch()

        columns_layout.addLayout(row_layout)
        columns_group.setLayout(columns_layout)
        layout.addWidget(columns_group)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def get_config(self) -> dict:
        sheet_mapping = {}
        for i in range(self.sheet_table.rowCount()):
            excel_sheet = self.sheet_table.item(i, 0).text()
            google_combo = self.sheet_table.cellWidget(i, 1)
            google_sheet = google_combo.currentText()
            if google_sheet != "-- Не копировать --":
                sheet_mapping[excel_sheet] = google_sheet

        source_cols = [col.strip() for col in self.source_columns.text().split(',') if col.strip()]
        target_cols = [col.strip() for col in self.target_columns.text().split(',') if col.strip()]

        return {
            'sheet_mapping': sheet_mapping,
            'column_mapping': {
                'source': source_cols,
                'target': target_cols
            },
            'start_row': self.start_row.value()
        }


class DropArea(QFrame):
    """Область для drag & drop файлов."""

    file_dropped = Signal(str)
    files_dropped = Signal(list)

    def __init__(self, accept_multiple=False):
        super().__init__()
        self.accept_multiple = accept_multiple
        self.setAcceptDrops(True)
        # Простая рамка без разделения на внутренние части
        self.setFrameStyle(QFrame.NoFrame)
        self.setStyleSheet(
            """
            QFrame {
                border: 2px dashed #aaa;
                border-radius: 10px;
                background-color: #f5f5f5;
                min-height: 120px;
            }
            QFrame:hover {
                background-color: #e8e8e8;
                border-color: #666;
            }
            """
        )

        layout = QVBoxLayout()

        if accept_multiple:
            self.label = QLabel("📁 Перетащите Excel файлы сюда")
            self.file_label = QLabel("можно выбрать несколько файлов")
        else:
            self.label = QLabel("📁 Перетащите Excel файл сюда")
            self.file_label = QLabel("или нажмите для выбора файла")

        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        font = QFont()
        font.setPointSize(11)
        self.label.setFont(font)
        self.label.setStyleSheet("color: #666;")

        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_label.setStyleSheet("color: #999; font-size: 9pt;")

        layout.addStretch()
        layout.addWidget(self.label)
        layout.addWidget(self.file_label)
        layout.addStretch()

        self.setLayout(layout)

    def mouseDoubleClickEvent(self, event):
        self.open_file_dialog()
        super().mouseDoubleClickEvent(event)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            valid_files = [u for u in urls if u.toLocalFile().lower().endswith(('.xlsx', '.xls'))]
            if valid_files:
                event.acceptProposedAction()
                self.setStyleSheet(
                    """
                    QFrame {
                        border: 2px solid #4CAF50;
                        border-radius: 10px;
                        background-color: #e8f5e9;
                        min-height: 120px;
                    }
                    """
                )

    def dragLeaveEvent(self, event):
        self.setStyleSheet(
            """
            QFrame {
                border: 2px dashed #aaa;
                border-radius: 10px;
                background-color: #f5f5f5;
                min-height: 120px;
            }
            """
        )

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile().lower().endswith(('.xlsx', '.xls'))]
        if files:
            if self.accept_multiple:
                self.files_dropped.emit(files)
                self.update_files_info(files)
            else:
                self.file_dropped.emit(files[0])
                self.update_file_info(files[0])

        self.setStyleSheet(
            """
            QFrame {
                border: 2px dashed #aaa;
                border-radius: 10px;
                background-color: #f5f5f5;
                min-height: 120px;
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
        self.label.setText(f"📄 {file_name}")
        self.file_label.setText(f"Размер: {self._get_file_size(file_path)}")

    def update_files_info(self, files: List[str]):
        self.label.setText(f"📄 Выбрано файлов: {len(files)}")
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

