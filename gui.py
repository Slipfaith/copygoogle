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

from processor import ExcelToGoogleSheets
from config import BASE_DIR, create_sample_config
from dialogs import BatchMappingDialog, MappingDialog, DropArea

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