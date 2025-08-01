"""
GUI для приложения копирования данных из Excel в Google Таблицы
ОБНОВЛЕННАЯ ВЕРСИЯ с улучшенными диалогами
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
    QGraphicsDropShadowEffect, QSizePolicy
)
from PySide6.QtCore import Qt, QThread, Signal, QMimeData, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QPalette, QColor, QFont, QIcon

from business.processor import ExcelToGoogleSheets
from config import BASE_DIR, create_sample_config

# Импортируем обычные диалоги (НЕ improved_dialogs!)
from dialogs import BatchMappingDialog, MappingDialog, DropArea

BASE_DIR = Path(__file__).resolve().parent


class ModernDropArea(QWidget):
    """Современная компактная область для drag & drop файлов."""

    file_dropped = Signal(str)
    files_dropped = Signal(list)

    def __init__(self, accept_multiple=False):
        super().__init__()
        self.accept_multiple = accept_multiple
        self.setAcceptDrops(True)
        self.setFixedHeight(80)
        self.setMaximumWidth(400)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)

        self.setStyleSheet("""
            ModernDropArea {
                background-color: #f8f9fa;
                border: 2px dashed #dee2e6;
                border-radius: 8px;
            }
        """)

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 15, 20, 15)

        # Текст
        if accept_multiple:
            text = "Перетащите файлы или нажмите для выбора"
        else:
            text = "Перетащите файл или нажмите для выбора"

        self.label = QLabel(text)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet("""
            color: #6c757d;
            font-size: 13px;
        """)

        self.file_info = QLabel("")
        self.file_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_info.setStyleSheet("""
            color: #28a745;
            font-size: 12px;
            font-weight: 500;
        """)
        self.file_info.hide()

        layout.addWidget(self.label)
        layout.addWidget(self.file_info)

        self.setLayout(layout)

        # Добавляем эффект тени
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(10)
        shadow.setColor(QColor(0, 0, 0, 30))
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)

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
                self.setStyleSheet("""
                    ModernDropArea {
                        background-color: #e7f3ff;
                        border: 2px solid #0066cc;
                        border-radius: 8px;
                    }
                """)

    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            ModernDropArea {
                background-color: #f8f9fa;
                border: 2px dashed #dee2e6;
                border-radius: 8px;
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
            ModernDropArea {
                background-color: #f8f9fa;
                border: 2px dashed #dee2e6;
                border-radius: 8px;
            }
        """)

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
        self.label.hide()
        self.file_info.show()
        self.file_info.setText(f"✓ {os.path.basename(file_path)}")

    def update_files_info(self, files: List[str]):
        self.label.hide()
        self.file_info.show()
        self.file_info.setText(f"✓ Выбрано файлов: {len(files)}")

    def reset(self):
        self.label.show()
        self.file_info.hide()
        self.file_info.setText("")


class WorkerThread(QThread):
    """Поток для выполнения операций копирования"""

    progress_update = Signal(int, int, str)
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
            self.processor = ExcelToGoogleSheets(str(BASE_DIR / "config.yaml"))

            if self.mode == "single":
                excel_path = self.kwargs['excel_path']
                google_sheet_url = self.kwargs['google_sheet_url']
                config = self.kwargs['config']

                self.processor.update_config(
                    sheet_mapping=config['sheet_mapping'],
                    column_mapping=config['column_mapping'],
                    start_row=config['start_row']
                )

                self.log_message.emit("Подключение к Google Таблицам...")
                self.processor.connect_to_google_sheets(google_sheet_url)

                self.processor.process_excel_file(
                    excel_path,
                    progress_callback=self.progress_update.emit,
                    log_callback=self.log_message.emit
                )

            elif self.mode == "batch":
                file_mappings = self.kwargs['file_mappings']
                google_sheet_url = self.kwargs['google_sheet_url']

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
    """Главное окно приложения с современным дизайном"""

    def __init__(self):
        super().__init__()
        self.processor = ExcelToGoogleSheets(str(BASE_DIR / "config.yaml"))
        self.worker_thread = None

        self.single_file = None
        self.single_config = None
        self.batch_files = []
        self.batch_mappings = []
        self.log_file = None
        self.log_file_path = None

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Excel to Google Sheets - Улучшенная версия")
        self.setFixedSize(500, 600)

        # Центральный виджет
        central_widget = QWidget()
        central_widget.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            }
        """)
        self.setCentralWidget(central_widget)

        # Основной layout
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        # Заголовок
        title_layout = QVBoxLayout()
        title_layout.setSpacing(5)

        title = QLabel("Excel → Google Sheets")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            font-size: 24px;
            font-weight: 600;
            color: #212529;
        """)

        subtitle = QLabel("Улучшенное копирование данных")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("""
            font-size: 14px;
            color: #6c757d;
        """)

        title_layout.addWidget(title)
        title_layout.addWidget(subtitle)
        layout.addLayout(title_layout)

        # Поле для Google Sheets URL
        url_container = QWidget()
        url_container.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
                border-radius: 8px;
            }
        """)
        url_layout = QVBoxLayout()
        url_layout.setContentsMargins(15, 15, 15, 15)

        url_label = QLabel("🔗 Google Таблица")
        url_label.setStyleSheet("""
            font-size: 12px;
            font-weight: 500;
            color: #495057;
            margin-bottom: 5px;
        """)

        self.google_url_input = QLineEdit()
        self.google_url_input.setPlaceholderText("Вставьте ссылку на Google Таблицу")
        self.google_url_input.setStyleSheet("""
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
        """)

        url_layout.addWidget(url_label)
        url_layout.addWidget(self.google_url_input)
        url_container.setLayout(url_layout)
        layout.addWidget(url_container)

        # Табы
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
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
        """)

        # Вкладка одиночного файла
        single_tab = QWidget()
        single_layout = QVBoxLayout()
        single_layout.setSpacing(15)

        # Drop area для одиночного файла
        drop_container = QHBoxLayout()
        drop_container.addStretch()
        self.single_drop_area = ModernDropArea(accept_multiple=False)
        self.single_drop_area.file_dropped.connect(self.on_single_file_dropped)
        drop_container.addWidget(self.single_drop_area)
        drop_container.addStretch()
        single_layout.addLayout(drop_container)

        # Кнопки
        self.single_mapping_btn = self.create_button("⚙️ Настроить маппинг", "#6c757d", "#495057")
        self.single_mapping_btn.setEnabled(False)
        self.single_mapping_btn.clicked.connect(self.configure_single_mapping)

        self.single_process_btn = self.create_button("🚀 Начать копирование", "#28a745", "#218838", primary=True)
        self.single_process_btn.setEnabled(False)
        self.single_process_btn.clicked.connect(self.start_single_processing)

        single_layout.addWidget(self.single_mapping_btn)
        single_layout.addWidget(self.single_process_btn)
        single_layout.addStretch()

        single_tab.setLayout(single_layout)

        # Вкладка пакетной обработки
        batch_tab = QWidget()
        batch_layout = QVBoxLayout()
        batch_layout.setSpacing(15)

        # Drop area для множественных файлов
        batch_drop_container = QHBoxLayout()
        batch_drop_container.addStretch()
        self.batch_drop_area = ModernDropArea(accept_multiple=True)
        self.batch_drop_area.files_dropped.connect(self.on_batch_files_dropped)
        batch_drop_container.addWidget(self.batch_drop_area)
        batch_drop_container.addStretch()
        batch_layout.addLayout(batch_drop_container)

        # Список файлов
        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(100)
        self.files_list.setStyleSheet("""
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
        """)
        batch_layout.addWidget(self.files_list)

        # Кнопки управления списком
        list_btns_layout = QHBoxLayout()

        clear_btn = self.create_small_button("🗑️ Очистить")
        clear_btn.clicked.connect(self.clear_batch_files)

        remove_btn = self.create_small_button("➖ Удалить выбранные")
        remove_btn.clicked.connect(self.remove_selected_files)

        list_btns_layout.addWidget(clear_btn)
        list_btns_layout.addWidget(remove_btn)
        list_btns_layout.addStretch()
        batch_layout.addLayout(list_btns_layout)

        # Кнопки действий
        self.batch_mapping_btn = self.create_button("⚙️ Настроить маппинг", "#6c757d", "#495057")
        self.batch_mapping_btn.setEnabled(False)
        self.batch_mapping_btn.clicked.connect(self.configure_batch_mapping)

        self.batch_process_btn = self.create_button("🚀 Начать копирование", "#28a745", "#218838", primary=True)
        self.batch_process_btn.setEnabled(False)
        self.batch_process_btn.clicked.connect(self.start_batch_processing)

        batch_layout.addWidget(self.batch_mapping_btn)
        batch_layout.addWidget(self.batch_process_btn)
        batch_layout.addStretch()

        batch_tab.setLayout(batch_layout)

        # Добавляем вкладки
        self.tabs.addTab(single_tab, "📄 Один файл")
        self.tabs.addTab(batch_tab, "📁 Несколько файлов")
        layout.addWidget(self.tabs)

        # Прогресс
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
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
        """)
        layout.addWidget(self.progress_bar)

        # Статус
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("""
            color: #6c757d;
            font-size: 12px;
        """)
        self.status_label.hide()
        layout.addWidget(self.status_label)

        # Лог (компактный)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(80)
        self.log_text.setStyleSheet("""
            QTextEdit {
                border: 1px solid #dee2e6;
                border-radius: 6px;
                background-color: #f8f9fa;
                padding: 8px;
                font-family: 'SF Mono', Monaco, monospace;
                font-size: 11px;
                color: #495057;
            }
        """)
        self.log_text.hide()
        layout.addWidget(self.log_text)

        # Кнопка показать/скрыть лог
        self.toggle_log_btn = self.create_small_button("📋 Показать журнал")
        self.toggle_log_btn.clicked.connect(self.toggle_log)
        layout.addWidget(self.toggle_log_btn)

        central_widget.setLayout(layout)

        # Подключение сигналов
        self.google_url_input.textChanged.connect(self.check_ready_state)
        self.tabs.currentChanged.connect(self.check_ready_state)

        # Проверка конфигурации
        self.check_config()

    def create_button(self, text: str, color: str, hover_color: str, primary: bool = False) -> QPushButton:
        """Создание стилизованной кнопки"""
        btn = QPushButton(text)
        if primary:
            btn.setStyleSheet(f"""
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
            """)
        else:
            btn.setStyleSheet(f"""
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
            """)
        return btn

    def create_small_button(self, text: str) -> QPushButton:
        """Создание маленькой кнопки"""
        btn = QPushButton(text)
        btn.setStyleSheet("""
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
        """)
        return btn

    def toggle_log(self):
        """Переключение видимости журнала"""
        if self.log_text.isVisible():
            self.log_text.hide()
            self.toggle_log_btn.setText("📋 Показать журнал")
        else:
            self.log_text.show()
            self.toggle_log_btn.setText("📋 Скрыть журнал")

    def check_config(self):
        """Проверка наличия конфигурационного файла"""
        config_path = BASE_DIR / "config.yaml"
        if not config_path.exists():
            self.log_message("⚠️ Создаю config.yaml...")
            try:
                create_sample_config(config_path)
                self.log_message("✓ Создан config.yaml")
            except Exception as e:
                self.log_message(f"❌ Ошибка: {e}")

        creds_path = BASE_DIR / "credentials.json"
        if not creds_path.exists():
            self.log_message("⚠️ Нужен credentials.json!")

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

    def on_single_file_dropped(self, file_path: str):
        """Обработка выбранного файла"""
        self.single_file = file_path
        self.single_config = None
        self.check_ready_state()
        self.log_message(f"✓ Файл: {os.path.basename(file_path)}")

    def configure_single_mapping(self):
        """Настройка маппинга для одного файла"""
        if not self.single_file or not self.google_url_input.text().strip():
            return

        try:
            self.log_message("🔍 Анализ файла...")

            excel_sheets = self.processor.get_excel_sheets(self.single_file)
            if not excel_sheets:
                raise Exception("Не удалось получить листы Excel")

            self.log_message("🔗 Подключение к Google Таблицам...")
            self.processor.connect_to_google_sheets(self.google_url_input.text().strip())
            google_sheets = self.processor.get_google_sheets()
            if not google_sheets:
                raise Exception("Не удалось получить листы Google")

            self.log_message("⚙️ Открытие настроек...")
            dialog = MappingDialog(excel_sheets, google_sheets, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.single_config = dialog.get_config()
                self.log_message("✅ Маппинг настроен успешно!")
                self.check_ready_state()

        except Exception as e:
            self.log_message(f"❌ Ошибка: {e}")
            QMessageBox.critical(
                self,
                "Ошибка настройки",
                f"Произошла ошибка при настройке маппинга:\n\n{str(e)}\n\nПроверьте:\n"
                "• Корректность ссылки на Google Таблицу\n"
                "• Наличие файла credentials.json\n"
                "• Доступ к интернету"
            )

    def start_single_processing(self):
        """Запуск обработки одного файла"""
        if not self.single_file or not self.google_url_input.text().strip() or not self.single_config:
            return

        self.disable_ui()
        self.show_progress()
        self.log_text.clear()

        header = [
            f"🚀 Начало обработки: {datetime.now().strftime('%H:%M:%S')}",
            f"📄 Файл: {os.path.basename(self.single_file)}",
            f"🔗 Таблица: {self.google_url_input.text().strip()[:50]}..."
        ]
        self.open_log_file(header)

        self.worker_thread = WorkerThread(
            mode="single",
            excel_path=self.single_file,
            google_sheet_url=self.google_url_input.text().strip(),
            config=self.single_config
        )
        self.connect_worker_signals()
        self.worker_thread.start()

    def on_batch_files_dropped(self, files: List[str]):
        """Обработка выбранных файлов"""
        added_count = 0
        for file in files:
            if file not in self.batch_files:
                self.batch_files.append(file)
                item = QListWidgetItem(f"📄 {os.path.basename(file)}")
                item.setData(Qt.ItemDataRole.UserRole, file)
                self.files_list.addItem(item)
                added_count += 1

        self.batch_mappings = []
        self.check_ready_state()
        if added_count > 0:
            self.log_message(f"✅ Добавлено файлов: {added_count}")

    def clear_batch_files(self):
        """Очистка списка файлов"""
        self.batch_files = []
        self.batch_mappings = []
        self.files_list.clear()
        self.batch_drop_area.reset()
        self.check_ready_state()
        self.log_message("🗑️ Список файлов очищен")

    def remove_selected_files(self):
        """Удаление выбранных файлов"""
        selected_items = self.files_list.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Внимание", "Выберите файлы для удаления")
            return

        removed_count = 0
        for item in selected_items:
            file_path = item.data(Qt.ItemDataRole.UserRole)
            if file_path in self.batch_files:
                self.batch_files.remove(file_path)
                removed_count += 1
            self.files_list.takeItem(self.files_list.row(item))

        self.batch_mappings = []
        if not self.batch_files:
            self.batch_drop_area.reset()
        self.check_ready_state()
        if removed_count > 0:
            self.log_message(f"➖ Удалено файлов: {removed_count}")

    def configure_batch_mapping(self):
        """Настройка маппинга для пакетной обработки"""
        if not self.batch_files or not self.google_url_input.text().strip():
            return

        try:
            self.log_message("🔗 Подключение к Google Таблицам...")

            self.processor.connect_to_google_sheets(self.google_url_input.text().strip())
            google_sheets = self.processor.get_google_sheets()
            if not google_sheets:
                raise Exception("Не удалось получить листы Google")

            self.log_message("⚙️ Открытие настроек пакетного маппинга...")
            dialog = BatchMappingDialog(self.batch_files, google_sheets, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.batch_mappings = dialog.mappings
                self.log_message(f"✅ Настроено маппинг для {len(self.batch_mappings)} файлов")
                self.check_ready_state()

        except Exception as e:
            self.log_message(f"❌ Ошибка: {e}")
            QMessageBox.critical(
                self,
                "Ошибка настройки",
                f"Произошла ошибка при настройке маппинга:\n\n{str(e)}\n\nПроверьте:\n"
                "• Корректность ссылки на Google Таблицу\n"
                "• Наличие файла credentials.json\n"
                "• Доступ к интернету"
            )

    def start_batch_processing(self):
        """Запуск пакетной обработки"""
        if not self.batch_mappings or not self.google_url_input.text().strip():
            return

        self.disable_ui()
        self.show_progress()
        self.log_text.clear()

        header = [
            f"🚀 Начало пакетной обработки: {datetime.now().strftime('%H:%M:%S')}",
            f"📁 Файлов к обработке: {len(self.batch_mappings)}",
            f"🔗 Таблица: {self.google_url_input.text().strip()[:50]}..."
        ]
        self.open_log_file(header)

        self.worker_thread = WorkerThread(
            mode="batch",
            file_mappings=self.batch_mappings,
            google_sheet_url=self.google_url_input.text().strip()
        )
        self.connect_worker_signals()
        self.worker_thread.start()

    def connect_worker_signals(self):
        """Подключение сигналов потока"""
        self.worker_thread.progress_update.connect(self.update_progress)
        self.worker_thread.log_message.connect(self.log_message)
        self.worker_thread.finished_successfully.connect(self.on_processing_finished)
        self.worker_thread.error_occurred.connect(self.on_processing_error)

    def disable_ui(self):
        """Отключение UI"""
        self.tabs.setEnabled(False)
        self.google_url_input.setEnabled(False)
        self.single_mapping_btn.setEnabled(False)
        self.single_process_btn.setEnabled(False)
        self.batch_mapping_btn.setEnabled(False)
        self.batch_process_btn.setEnabled(False)

    def enable_ui(self):
        """Включение UI"""
        self.tabs.setEnabled(True)
        self.google_url_input.setEnabled(True)
        self.check_ready_state()

    def show_progress(self):
        """Показ прогресса"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.show()
        if not self.log_text.isVisible():
            self.toggle_log()

    def hide_progress(self):
        """Скрытие прогресса"""
        self.progress_bar.setVisible(False)
        self.status_label.hide()

    def open_log_file(self, header_lines):
        """Открытие файла логов"""
        logs_dir = BASE_DIR / "logs"
        logs_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file_path = logs_dir / f"log_{timestamp}.txt"
        self.log_file = open(self.log_file_path, "w", encoding="utf-8")
        for line in header_lines:
            self.log_file.write(line + "\n")
        self.log_file.write("\n")

    def close_log_file(self):
        """Закрытие файла логов"""
        if self.log_file:
            self.log_file.close()
            self.log_file = None

    def update_progress(self, current: int, total: int, item_name: str):
        """Обновление прогресса"""
        progress = int((current / total) * 100) if total > 0 else 0
        self.progress_bar.setValue(progress)

        if self.tabs.currentIndex() == 0:
            self.progress_bar.setFormat(f"{progress}% - {item_name}")
            self.status_label.setText(f"📋 Листов: {current}/{total}")
        else:
            self.progress_bar.setFormat(f"{progress}% - {item_name}")
            self.status_label.setText(f"📁 Файлов: {current}/{total}")

    def log_message(self, message: str):
        """Добавление сообщения в лог"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.log_text.append(formatted_message)

        if self.log_file:
            self.log_file.write(formatted_message + "\n")
            self.log_file.flush()

        # Автопрокрутка к концу
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def on_processing_finished(self):
        """Успешное завершение"""
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("✅ Готово!")
        self.status_label.setText("🎉 Завершено успешно")

        self.close_log_file()
        self.log_message("🎉 Обработка завершена успешно!")

        QTimer.singleShot(3000, self.hide_progress)
        self.enable_ui()

        # Красивое уведомление об успехе
        msg = QMessageBox(self)
        msg.setWindowTitle("Успешно!")
        msg.setText("🎉 Данные успешно скопированы!")
        msg.setInformativeText(
            f"Все данные были успешно перенесены в Google Таблицы.\n"
            f"Лог сохранен в: {self.log_file_path.name if self.log_file_path else 'неизвестно'}"
        )
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def on_processing_error(self, error_message: str):
        """Обработка ошибки"""
        self.log_message(f"💥 КРИТИЧЕСКАЯ ОШИБКА: {error_message}")
        self.hide_progress()
        self.close_log_file()
        self.enable_ui()

        # Подробное сообщение об ошибке
        msg = QMessageBox(self)
        msg.setWindowTitle("Ошибка обработки")
        msg.setText("💥 Произошла ошибка при обработке")
        msg.setInformativeText(
            f"Детали ошибки:\n{error_message}\n\n"
            "Возможные причины:\n"
            "• Нет доступа к интернету\n"
            "• Неверные права доступа к Google Таблице\n"
            "• Поврежденный Excel файл\n"
            "• Неверная настройка credentials.json"
        )
        msg.setIcon(QMessageBox.Icon.Critical)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()


def main():
    app = QApplication(sys.argv)

    # Современный стиль
    app.setStyle("Fusion")

    # Светлая палитра
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(33, 37, 41))
    palette.setColor(QPalette.ColorRole.Base, QColor(248, 249, 250))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(233, 236, 239))
    palette.setColor(QPalette.ColorRole.Text, QColor(33, 37, 41))
    palette.setColor(QPalette.ColorRole.Button, QColor(248, 249, 250))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(33, 37, 41))
    palette.setColor(QPalette.ColorRole.BrightText, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.Link, QColor(0, 102, 204))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(0, 102, 204))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()