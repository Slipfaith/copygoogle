"""
GUI для приложения копирования данных из Excel в Google Таблицы
"""

import sys
import os
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, List, Tuple

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QProgressBar, QLabel, QFrame,
    QMessageBox, QFileDialog, QLineEdit, QDialog, QDialogButtonBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox,
    QGroupBox, QSpinBox, QTabWidget, QListWidget, QListWidgetItem,
    QSplitter, QCheckBox
)
from PySide6.QtCore import Qt, QThread, Signal, QMimeData, QTimer
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QPalette, QColor, QFont, QIcon

from excel_to_google_sheets import ExcelToGoogleSheets

BASE_DIR = Path(__file__).resolve().parent


class WorkerThread(QThread):
    """Поток для выполнения операций копирования"""
    
    progress_update = Signal(int, int, str)  # current, total, sheet_name
    log_message = Signal(str)
    finished_successfully = Signal()
    error_occurred = Signal(str)
    
    def __init__(self, mode: str, **kwargs):
        super().__init__()
        self.mode = mode
        self.kwargs = kwargs
        self.processor = None
        
    def run(self):
        try:
            # Создание процессора
            self.processor = ExcelToGoogleSheets(str(BASE_DIR / "config.yaml"))
            
            if self.mode == "single":
                # Одиночный файл
                excel_path = self.kwargs['excel_path']
                google_sheet_url = self.kwargs['google_sheet_url']
                config = self.kwargs['config']
                
                # Обновление конфигурации
                self.processor.update_config(
                    sheet_mapping=config['sheet_mapping'],
                    column_mapping=config['column_mapping'],
                    start_row=config['start_row']
                )
                
                # Подключение к Google Таблицам
                self.log_message.emit("Подключение к Google Таблицам...")
                self.processor.connect_to_google_sheets(google_sheet_url)
                
                # Обработка файла
                self.processor.process_excel_file(
                    excel_path,
                    progress_callback=self.progress_update.emit,
                    log_callback=self.log_message.emit
                )
                
            elif self.mode == "batch":
                # Пакетная обработка
                file_mappings = self.kwargs['file_mappings']
                google_sheet_url = self.kwargs['google_sheet_url']
                
                # Обработка нескольких файлов
                self.processor.process_multiple_excel_files(
                    file_mappings,
                    google_sheet_url,
                    progress_callback=self.progress_update.emit,
                    log_callback=self.log_message.emit
                )
            
            self.finished_successfully.emit()
            
        except Exception as e:
            self.error_occurred.emit(str(e))


class BatchMappingDialog(QDialog):
    """Диалог настройки маппинга для пакетной обработки"""
    
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
        
        # Инструкция
        info = QLabel("Настройте маппинг для каждого Excel файла:")
        info.setStyleSheet("font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(info)
        
        # Таблица маппингов
        self.mapping_table = QTableWidget()
        self.mapping_table.setColumnCount(6)
        self.mapping_table.setHorizontalHeaderLabels([
            "Excel файл", "Excel лист", "→", "Google лист", 
            "Колонки (из → в)", "Начальная строка"
        ])
        
        # Настройка ширины колонок
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
        
        # Добавляем строки для каждого Excel файла
        self.mapping_table.setRowCount(len(self.excel_files))
        
        for i, excel_file in enumerate(self.excel_files):
            # Excel файл (не редактируемый)
            file_item = QTableWidgetItem(os.path.basename(excel_file))
            file_item.setData(Qt.ItemDataRole.UserRole, excel_file)  # Сохраняем полный путь
            file_item.setFlags(file_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.mapping_table.setItem(i, 0, file_item)
            
            # Excel лист (по умолчанию Sheet1)
            sheet_item = QTableWidgetItem("Sheet1")
            self.mapping_table.setItem(i, 1, sheet_item)
            
            # Стрелка
            arrow_item = QTableWidgetItem("→")
            arrow_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            arrow_item.setFlags(arrow_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.mapping_table.setItem(i, 2, arrow_item)
            
            # Google лист (ComboBox)
            google_combo = QComboBox()
            google_combo.addItem("-- Не копировать --")
            google_combo.addItems(self.google_sheets)
            
            # Пытаемся найти похожее название
            file_name_without_ext = os.path.splitext(os.path.basename(excel_file))[0]
            for sheet in self.google_sheets:
                if file_name_without_ext.lower() in sheet.lower() or sheet.lower() in file_name_without_ext.lower():
                    google_combo.setCurrentText(sheet)
                    break
            
            self.mapping_table.setCellWidget(i, 3, google_combo)
            
            # Маппинг колонок
            columns_item = QTableWidgetItem("A → A")
            self.mapping_table.setItem(i, 4, columns_item)
            
            # Начальная строка
            start_row_spin = QSpinBox()
            start_row_spin.setMinimum(1)
            start_row_spin.setMaximum(10000)
            start_row_spin.setValue(1)
            self.mapping_table.setCellWidget(i, 5, start_row_spin)
        
        layout.addWidget(self.mapping_table)
        
        # Подсказка
        hint = QLabel("Формат колонок: 'A,B,C → D,E,F' или 'A-C → D-F'")
        hint.setStyleSheet("color: #666; font-style: italic; margin-top: 5px;")
        layout.addWidget(hint)
        
        # Кнопки быстрых действий
        quick_actions = QHBoxLayout()
        
        select_all_btn = QPushButton("Выбрать все")
        select_all_btn.clicked.connect(self.select_all_sheets)
        quick_actions.addWidget(select_all_btn)
        
        auto_map_btn = QPushButton("Авто-маппинг по именам")
        auto_map_btn.clicked.connect(self.auto_map_by_names)
        quick_actions.addWidget(auto_map_btn)
        
        quick_actions.addStretch()
        layout.addLayout(quick_actions)
        
        # Кнопки диалога
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        
        layout.addWidget(buttons)
        self.setLayout(layout)
    
    def select_all_sheets(self):
        """Выбрать первый доступный Google лист для всех файлов"""
        for i in range(self.mapping_table.rowCount()):
            combo = self.mapping_table.cellWidget(i, 3)
            if combo and combo.count() > 1:
                combo.setCurrentIndex(1)  # Первый лист после "-- Не копировать --"
    
    def auto_map_by_names(self):
        """Автоматический маппинг по совпадению имен"""
        for i in range(self.mapping_table.rowCount()):
            file_path = self.mapping_table.item(i, 0).data(Qt.ItemDataRole.UserRole)
            file_name = os.path.splitext(os.path.basename(file_path))[0].lower()
            
            combo = self.mapping_table.cellWidget(i, 3)
            if combo:
                # Ищем наилучшее совпадение
                best_match = None
                best_score = 0
                
                for j in range(1, combo.count()):  # Пропускаем "-- Не копировать --"
                    sheet_name = combo.itemText(j).lower()
                    
                    # Различные стратегии совпадения
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
        """Валидация и сохранение маппингов"""
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
            
            # Парсинг маппинга колонок
            try:
                source_cols, target_cols = self.parse_column_mapping(columns_text)
            except ValueError as e:
                QMessageBox.warning(
                    self, 
                    "Ошибка", 
                    f"Неверный формат колонок в строке {i+1}: {e}"
                )
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
            QMessageBox.warning(self, "Внимание", "Не выбрано ни одного файла для копирования")
            return
        
        self.accept()
    
    def parse_column_mapping(self, text: str) -> Tuple[List[str], List[str]]:
        """Парсинг текста маппинга колонок"""
        parts = text.split('→')
        if len(parts) != 2:
            raise ValueError("Используйте формат: 'A,B,C → D,E,F'")
        
        source_part = parts[0].strip()
        target_part = parts[1].strip()
        
        # Парсинг исходных колонок
        source_cols = self.parse_column_range(source_part)
        target_cols = self.parse_column_range(target_part)
        
        if len(source_cols) != len(target_cols):
            raise ValueError("Количество исходных и целевых колонок должно совпадать")
        
        return source_cols, target_cols
    
    def parse_column_range(self, text: str) -> List[str]:
        """Парсинг диапазона колонок (поддержка A-C и A,B,C)"""
        text = text.strip().upper()
        
        if '-' in text and ',' not in text:
            # Диапазон типа A-C
            parts = text.split('-')
            if len(parts) != 2:
                raise ValueError(f"Неверный диапазон: {text}")
            
            start_col = parts[0].strip()
            end_col = parts[1].strip()
            
            if not start_col.isalpha() or not end_col.isalpha():
                raise ValueError(f"Неверные колонки: {text}")
            
            start_ord = ord(start_col)
            end_ord = ord(end_col)
            
            if start_ord > end_ord:
                raise ValueError(f"Неверный диапазон: {text}")
            
            return [chr(i) for i in range(start_ord, end_ord + 1)]
        else:
            # Список типа A,B,C
            cols = [col.strip() for col in text.split(',')]
            for col in cols:
                if not col.isalpha():
                    raise ValueError(f"Неверная колонка: {col}")
            return cols


class MappingDialog(QDialog):
    """Диалог настройки маппинга для одного файла"""
    
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
        
        # Маппинг листов
        sheets_group = QGroupBox("Маппинг листов")
        sheets_layout = QVBoxLayout()
        
        self.sheet_table = QTableWidget()
        self.sheet_table.setColumnCount(2)
        self.sheet_table.setHorizontalHeaderLabels(["Excel лист", "Google лист"])
        self.sheet_table.horizontalHeader().setStretchLastSection(True)
        
        # Добавляем строки для каждого Excel листа
        self.sheet_table.setRowCount(len(self.excel_sheets))
        for i, excel_sheet in enumerate(self.excel_sheets):
            # Excel лист (не редактируемый)
            excel_item = QTableWidgetItem(excel_sheet)
            excel_item.setFlags(excel_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.sheet_table.setItem(i, 0, excel_item)
            
            # Google лист (ComboBox)
            google_combo = QComboBox()
            google_combo.addItem("-- Не копировать --")
            google_combo.addItems(self.google_sheets)
            # Пытаемся найти соответствие по имени
            if excel_sheet in self.google_sheets:
                google_combo.setCurrentText(excel_sheet)
            self.sheet_table.setCellWidget(i, 1, google_combo)
        
        sheets_layout.addWidget(self.sheet_table)
        sheets_group.setLayout(sheets_layout)
        layout.addWidget(sheets_group)
        
        # Маппинг колонок
        columns_group = QGroupBox("Маппинг колонок")
        columns_layout = QVBoxLayout()
        
        columns_info = QLabel("Укажите какие колонки копировать:")
        columns_layout.addWidget(columns_info)
        
        # Поля для ввода колонок
        columns_input_layout = QHBoxLayout()
        
        columns_input_layout.addWidget(QLabel("Из Excel:"))
        self.source_columns = QLineEdit("A")
        self.source_columns.setPlaceholderText("Например: A, C, E")
        columns_input_layout.addWidget(self.source_columns)
        
        columns_input_layout.addWidget(QLabel("В Google:"))
        self.target_columns = QLineEdit("A")
        self.target_columns.setPlaceholderText("Например: B, D, F")
        columns_input_layout.addWidget(self.target_columns)
        
        columns_layout.addLayout(columns_input_layout)
        
        # Начальная строка
        row_layout = QHBoxLayout()
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
        
        # Кнопки
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        
        layout.addWidget(buttons)
        self.setLayout(layout)
    
    def get_config(self) -> Dict:
        """Получение конфигурации из диалога"""
        # Маппинг листов
        sheet_mapping = {}
        for i in range(self.sheet_table.rowCount()):
            excel_sheet = self.sheet_table.item(i, 0).text()
            google_combo = self.sheet_table.cellWidget(i, 1)
            google_sheet = google_combo.currentText()
            
            if google_sheet != "-- Не копировать --":
                sheet_mapping[excel_sheet] = google_sheet
        
        # Маппинг колонок
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
    """Область для drag & drop файлов"""
    
    file_dropped = Signal(str)
    files_dropped = Signal(list)  # Для множественных файлов
    
    def __init__(self, accept_multiple=False):
        super().__init__()
        self.accept_multiple = accept_multiple
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Sunken)
        self.setStyleSheet("""
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
        """)
        
        layout = QVBoxLayout()
        
        # Иконка и текст
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
        
        # Делаем область кликабельной
        self.mousePressEvent = self.open_file_dialog
        
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            valid_files = [u for u in urls if u.toLocalFile().lower().endswith(('.xlsx', '.xls'))]
            if valid_files:
                event.acceptProposedAction()
                self.setStyleSheet("""
                    QFrame {
                        border: 2px solid #4CAF50;
                        border-radius: 10px;
                        background-color: #e8f5e9;
                        min-height: 120px;
                    }
                """)
    
    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            QFrame {
                border: 2px dashed #aaa;
                border-radius: 10px;
                background-color: #f5f5f5;
                min-height: 120px;
            }
        """)
    
    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls() 
                if u.toLocalFile().lower().endswith(('.xlsx', '.xls'))]
        
        if files:
            if self.accept_multiple:
                self.files_dropped.emit(files)
                self.update_files_info(files)
            else:
                self.file_dropped.emit(files[0])
                self.update_file_info(files[0])
        
        self.setStyleSheet("""
            QFrame {
                border: 2px dashed #aaa;
                border-radius: 10px;
                background-color: #f5f5f5;
                min-height: 120px;
            }
        """)
    
    def open_file_dialog(self, event):
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
        """Обновление информации о выбранном файле"""
        file_name = os.path.basename(file_path)
        self.label.setText(f"📄 {file_name}")
        self.file_label.setText(f"Размер: {self._get_file_size(file_path)}")
    
    def update_files_info(self, files: List[str]):
        """Обновление информации о выбранных файлах"""
        self.label.setText(f"📄 Выбрано файлов: {len(files)}")
        total_size = sum(os.path.getsize(f) for f in files)
        self.file_label.setText(f"Общий размер: {self._format_size(total_size)}")
    
    def _get_file_size(self, file_path: str) -> str:
        """Получение размера файла в читаемом формате"""
        size = os.path.getsize(file_path)
        return self._format_size(size)
    
    def _format_size(self, size: int) -> str:
        """Форматирование размера в читаемый вид"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"


class MainWindow(QMainWindow):
    """Главное окно приложения"""
    
    def __init__(self):
        super().__init__()
        self.processor = ExcelToGoogleSheets(str(BASE_DIR / "config.yaml"))
        self.worker_thread = None
        
        # Данные для разных режимов
        self.single_file = None
        self.single_config = None
        self.batch_files = []
        self.batch_mappings = []
        
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Excel → Google Таблицы")
        self.setMinimumSize(800, 700)
        
        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Основной layout
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Заголовок
        title = QLabel("Копирование данных из Excel в Google Таблицы")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        # Поле для ввода ссылки на Google Таблицу
        google_group = QGroupBox("Google Таблица")
        google_layout = QVBoxLayout()
        
        self.google_url_input = QLineEdit()
        self.google_url_input.setPlaceholderText("Вставьте ссылку на Google Таблицу...")
        self.google_url_input.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                font-size: 11pt;
                border: 2px solid #ddd;
                border-radius: 5px;
            }
            QLineEdit:focus {
                border-color: #4CAF50;
            }
        """)
        google_layout.addWidget(self.google_url_input)
        
        google_group.setLayout(google_layout)
        layout.addWidget(google_group)
        
        # Табы для разных режимов
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ddd;
                background-color: white;
            }
            QTabBar::tab {
                padding: 8px 16px;
                margin-right: 4px;
            }
            QTabBar::tab:selected {
                background-color: #4CAF50;
                color: white;
            }
        """)
        
        # Вкладка одиночного файла
        single_tab = QWidget()
        single_layout = QVBoxLayout()
        
        self.single_drop_area = DropArea(accept_multiple=False)
        self.single_drop_area.file_dropped.connect(self.on_single_file_dropped)
        single_layout.addWidget(self.single_drop_area)
        
        # Кнопки для одиночного режима
        single_buttons = QHBoxLayout()
        
        self.single_mapping_btn = QPushButton("⚙️ Настроить маппинг")
        self.single_mapping_btn.setEnabled(False)
        self.single_mapping_btn.clicked.connect(self.configure_single_mapping)
        self.single_mapping_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        
        self.single_process_btn = QPushButton("▶️ Начать копирование")
        self.single_process_btn.setEnabled(False)
        self.single_process_btn.clicked.connect(self.start_single_processing)
        self.single_process_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
                border-radius: 5px;
                font-size: 12pt;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        
        single_buttons.addWidget(self.single_mapping_btn)
        single_buttons.addWidget(self.single_process_btn)
        single_buttons.addStretch()
        
        single_layout.addLayout(single_buttons)
        single_tab.setLayout(single_layout)
        
        # Вкладка пакетной обработки
        batch_tab = QWidget()
        batch_layout = QVBoxLayout()
        
        batch_info = QLabel("🔄 Режим пакетной обработки: каждый Excel файл → отдельный лист Google")
        batch_info.setStyleSheet("color: #1976D2; font-weight: bold; margin-bottom: 10px;")
        batch_layout.addWidget(batch_info)
        
        # Разделитель для файлов и списка
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # Левая часть - drag&drop
        left_widget = QWidget()
        left_layout = QVBoxLayout()
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        self.batch_drop_area = DropArea(accept_multiple=True)
        self.batch_drop_area.files_dropped.connect(self.on_batch_files_dropped)
        left_layout.addWidget(self.batch_drop_area)
        
        left_widget.setLayout(left_layout)
        
        # Правая часть - список файлов
        right_widget = QWidget()
        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        files_label = QLabel("📁 Выбранные файлы:")
        files_label.setStyleSheet("font-weight: bold;")
        right_layout.addWidget(files_label)
        
        self.files_list = QListWidget()
        self.files_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 5px;
                background-color: #f9f9f9;
            }
        """)
        right_layout.addWidget(self.files_list)
        
        # Кнопки управления списком
        list_buttons = QHBoxLayout()
        
        clear_btn = QPushButton("Очистить")
        clear_btn.clicked.connect(self.clear_batch_files)
        list_buttons.addWidget(clear_btn)
        
        remove_btn = QPushButton("Удалить выбранные")
        remove_btn.clicked.connect(self.remove_selected_files)
        list_buttons.addWidget(remove_btn)
        
        list_buttons.addStretch()
        right_layout.addLayout(list_buttons)
        
        right_widget.setLayout(right_layout)
        
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([400, 300])
        
        batch_layout.addWidget(splitter)
        
        # Кнопки для пакетного режима
        batch_buttons = QHBoxLayout()
        
        self.batch_mapping_btn = QPushButton("⚙️ Настроить маппинг файлов")
        self.batch_mapping_btn.setEnabled(False)
        self.batch_mapping_btn.clicked.connect(self.configure_batch_mapping)
        self.batch_mapping_btn.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        
        self.batch_process_btn = QPushButton("▶️ Начать пакетное копирование")
        self.batch_process_btn.setEnabled(False)
        self.batch_process_btn.clicked.connect(self.start_batch_processing)
        self.batch_process_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
                border-radius: 5px;
                font-size: 12pt;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        
        batch_buttons.addWidget(self.batch_mapping_btn)
        batch_buttons.addWidget(self.batch_process_btn)
        batch_buttons.addStretch()
        
        batch_layout.addLayout(batch_buttons)
        batch_tab.setLayout(batch_layout)
        
        # Добавление вкладок
        self.tabs.addTab(single_tab, "📄 Один файл")
        self.tabs.addTab(batch_tab, "📚 Пакетная обработка")
        
        layout.addWidget(self.tabs)
        
        # Прогресс бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)
        
        # Текущий статус
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #666; font-size: 10pt;")
        layout.addWidget(self.status_label)
        
        # Лог
        log_label = QLabel("📋 Журнал операций:")
        log_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f9f9f9;
                border: 1px solid #ddd;
                border-radius: 5px;
                padding: 5px;
                font-family: 'Consolas', 'Monaco', monospace;
            }
        """)
        layout.addWidget(self.log_text)
        
        central_widget.setLayout(layout)
        
        # Подключение сигналов
        self.google_url_input.textChanged.connect(self.check_ready_state)
        self.tabs.currentChanged.connect(self.check_ready_state)
        
        # Проверка наличия config.yaml
        self.check_config()
    
    def check_config(self):
        """Проверка наличия конфигурационного файла"""
        config_path = BASE_DIR / "config.yaml"
        if not config_path.exists():
            self.log_message("⚠️ Файл config.yaml не найден. Создаю пример...")
            try:
                from excel_to_google_sheets import create_sample_config
                create_sample_config(config_path)
                self.log_message("✓ Создан пример config.yaml")
            except Exception as e:
                self.log_message(f"❌ Ошибка создания config.yaml: {e}")

        # Проверка credentials.json
        creds_path = BASE_DIR / "credentials.json"
        if not creds_path.exists():
            self.log_message("⚠️ Файл credentials.json не найден!")
            self.log_message("❗ Необходимо настроить Google Sheets API и получить credentials.json")
    
    def check_ready_state(self):
        """Проверка готовности к работе"""
        has_google_url = bool(self.google_url_input.text().strip())
        current_tab = self.tabs.currentIndex()
        
        if current_tab == 0:  # Одиночный файл
            has_file = self.single_file is not None
            self.single_mapping_btn.setEnabled(has_google_url and has_file)
            self.single_process_btn.setEnabled(has_google_url and has_file and self.single_config is not None)
        else:  # Пакетная обработка
            has_files = len(self.batch_files) > 0
            self.batch_mapping_btn.setEnabled(has_google_url and has_files)
            self.batch_process_btn.setEnabled(has_google_url and has_files and len(self.batch_mappings) > 0)
    
    # Методы для одиночного режима
    def on_single_file_dropped(self, file_path: str):
        """Обработка выбранного файла"""
        self.single_file = file_path
        self.single_config = None
        self.check_ready_state()
        self.log_message(f"✓ Выбран файл: {os.path.basename(file_path)}")
    
    def configure_single_mapping(self):
        """Настройка маппинга для одного файла"""
        if not self.single_file or not self.google_url_input.text().strip():
            return
        
        try:
            self.log_message("Получение информации о листах...")
            
            # Получение листов из Excel
            excel_sheets = self.processor.get_excel_sheets(self.single_file)
            if not excel_sheets:
                raise Exception("Не удалось получить список листов из Excel файла")
            
            # Подключение к Google Таблице
            self.processor.connect_to_google_sheets(self.google_url_input.text().strip())
            
            # Получение листов из Google
            google_sheets = self.processor.get_google_sheets()
            if not google_sheets:
                raise Exception("Не удалось получить список листов из Google Таблицы")
            
            # Открытие диалога маппинга
            dialog = MappingDialog(excel_sheets, google_sheets, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.single_config = dialog.get_config()
                self.log_message("✓ Маппинг настроен")
                self.check_ready_state()
            
        except Exception as e:
            self.log_message(f"❌ Ошибка: {e}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось получить информацию о листах:\n\n{e}")
    
    def start_single_processing(self):
        """Запуск обработки одного файла"""
        if not self.single_file or not self.google_url_input.text().strip() or not self.single_config:
            return
        
        self.disable_ui()
        self.show_progress()
        
        # Очистка лога
        self.log_text.clear()
        self.log_message(f"{'='*50}")
        self.log_message(f"Начало обработки: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.log_message(f"Excel файл: {os.path.basename(self.single_file)}")
        self.log_message(f"Google Таблица: {self.google_url_input.text().strip()}")
        self.log_message(f"{'='*50}")
        
        # Создание и запуск потока
        self.worker_thread = WorkerThread(
            mode="single",
            excel_path=self.single_file,
            google_sheet_url=self.google_url_input.text().strip(),
            config=self.single_config
        )
        self.connect_worker_signals()
        self.worker_thread.start()
    
    # Методы для пакетного режима
    def on_batch_files_dropped(self, files: List[str]):
        """Обработка выбранных файлов для пакетной обработки"""
        # Добавляем только новые файлы
        for file in files:
            if file not in self.batch_files:
                self.batch_files.append(file)
                self.files_list.addItem(os.path.basename(file))
        
        self.batch_mappings = []  # Сброс маппингов при добавлении новых файлов
        self.check_ready_state()
        self.log_message(f"✓ Добавлено файлов: {len(files)}")
    
    def clear_batch_files(self):
        """Очистка списка файлов"""
        self.batch_files = []
        self.batch_mappings = []
        self.files_list.clear()
        self.check_ready_state()
    
    def remove_selected_files(self):
        """Удаление выбранных файлов из списка"""
        for item in self.files_list.selectedItems():
            row = self.files_list.row(item)
            self.files_list.takeItem(row)
            if row < len(self.batch_files):
                self.batch_files.pop(row)
        
        self.batch_mappings = []  # Сброс маппингов при изменении списка файлов
        self.check_ready_state()
    
    def configure_batch_mapping(self):
        """Настройка маппинга для пакетной обработки"""
        if not self.batch_files or not self.google_url_input.text().strip():
            return
        
        try:
            self.log_message("Подключение к Google Таблице...")
            
            # Подключение к Google Таблице
            self.processor.connect_to_google_sheets(self.google_url_input.text().strip())
            
            # Получение листов из Google
            google_sheets = self.processor.get_google_sheets()
            if not google_sheets:
                raise Exception("Не удалось получить список листов из Google Таблицы")
            
            # Открытие диалога маппинга
            dialog = BatchMappingDialog(self.batch_files, google_sheets, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.batch_mappings = dialog.mappings
                self.log_message(f"✓ Настроен маппинг для {len(self.batch_mappings)} файлов")
                self.check_ready_state()
            
        except Exception as e:
            self.log_message(f"❌ Ошибка: {e}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось получить информацию о листах:\n\n{e}")
    
    def start_batch_processing(self):
        """Запуск пакетной обработки"""
        if not self.batch_mappings or not self.google_url_input.text().strip():
            return
        
        self.disable_ui()
        self.show_progress()
        
        # Очистка лога
        self.log_text.clear()
        self.log_message(f"{'='*50}")
        self.log_message(f"Начало пакетной обработки: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.log_message(f"Файлов для обработки: {len(self.batch_mappings)}")
        self.log_message(f"Google Таблица: {self.google_url_input.text().strip()}")
        self.log_message(f"{'='*50}")
        
        # Создание и запуск потока
        self.worker_thread = WorkerThread(
            mode="batch",
            file_mappings=self.batch_mappings,
            google_sheet_url=self.google_url_input.text().strip()
        )
        self.connect_worker_signals()
        self.worker_thread.start()
    
    # Общие методы
    def connect_worker_signals(self):
        """Подключение сигналов рабочего потока"""
        self.worker_thread.progress_update.connect(self.update_progress)
        self.worker_thread.log_message.connect(self.log_message)
        self.worker_thread.finished_successfully.connect(self.on_processing_finished)
        self.worker_thread.error_occurred.connect(self.on_processing_error)
    
    def disable_ui(self):
        """Отключение элементов управления во время обработки"""
        self.tabs.setEnabled(False)
        self.google_url_input.setEnabled(False)
        self.single_mapping_btn.setEnabled(False)
        self.single_process_btn.setEnabled(False)
        self.batch_mapping_btn.setEnabled(False)
        self.batch_process_btn.setEnabled(False)
    
    def enable_ui(self):
        """Включение элементов управления после обработки"""
        self.tabs.setEnabled(True)
        self.google_url_input.setEnabled(True)
        self.check_ready_state()
    
    def show_progress(self):
        """Показ прогресс-бара"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
    
    def hide_progress(self):
        """Скрытие прогресс-бара"""
        self.progress_bar.setVisible(False)
        self.status_label.setText("")
    
    def update_progress(self, current: int, total: int, item_name: str):
        """Обновление прогресс-бара"""
        progress = int((current / total) * 100)
        self.progress_bar.setValue(progress)
        
        if self.tabs.currentIndex() == 0:  # Одиночный режим
            self.progress_bar.setFormat(f"{progress}% - Обработка листа: {item_name}")
            self.status_label.setText(f"Обработано листов: {current} из {total}")
        else:  # Пакетный режим
            self.progress_bar.setFormat(f"{progress}% - Обработка файла: {item_name}")
            self.status_label.setText(f"Обработано файлов: {current} из {total}")
    
    def log_message(self, message: str):
        """Добавление сообщения в лог"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        
        # Автопрокрутка вниз
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def on_processing_finished(self):
        """Обработка успешного завершения"""
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("✓ Обработка завершена")
        self.status_label.setText("✓ Все операции выполнены успешно")
        
        self.log_message(f"{'='*50}")
        self.log_message(f"✓ Обработка завершена: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Скрытие прогресс-бара через 3 секунды
        QTimer.singleShot(3000, self.hide_progress)
        
        # Восстановление UI
        self.enable_ui()
        
        # Показ уведомления
        mode = "Данные" if self.tabs.currentIndex() == 0 else "Все файлы"
        QMessageBox.information(
            self,
            "Успешно",
            f"{mode} успешно скопированы в Google Таблицы!"
        )
    
    def on_processing_error(self, error_message: str):
        """Обработка ошибки"""
        self.log_message(f"❌ ОШИБКА: {error_message}")
        self.hide_progress()
        self.status_label.setText("❌ Произошла ошибка")
        
        # Восстановление UI
        self.enable_ui()
        
        # Показ сообщения об ошибке
        QMessageBox.critical(
            self,
            "Ошибка",
            f"Произошла ошибка при обработке:\n\n{error_message}"
        )


def main():
    app = QApplication(sys.argv)
    
    # Установка стиля приложения
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()