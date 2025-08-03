import sys
import os
from datetime import datetime
from typing import List

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QProgressBar,
    QListWidget, QListWidgetItem, QTabWidget, QMessageBox, QFileDialog,
    QDialog, QFrame, QSpacerItem, QSizePolicy
)
from PySide6.QtCore import Qt, QTimer

from app_logic import AppLogic
from config import BASE_DIR, create_sample_config
from log_service import LogService
from state import AppState

from .dialogs import BatchMappingDialog, MappingDialog, DownloadDialog
from . import styles
from .widgets import ModernDropArea
from .sliding_log_widget import SlidingLogWidget
from .utils import handle_errors


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.logic = AppLogic()
        self.state = AppState()
        self.logger = LogService(BASE_DIR)

        self.init_ui()
        self.connect_signals()
        self.check_config()

    def init_ui(self):
        self.setWindowTitle("Excel to Google Sheets")
        self.setFixedSize(750, 650)  # Уменьшили размер окна

        # Главный контейнер с горизонтальным разделением
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Основная область (слева)
        self.content_widget = QWidget()
        self.content_widget.setStyleSheet("background: white;")
        main_layout.addWidget(self.content_widget)

        # Слайдер логов (справа)
        self.sliding_log = SlidingLogWidget(self)
        main_layout.addWidget(self.sliding_log)

        # Основной layout для контента
        content_layout = QVBoxLayout(self.content_widget)
        content_layout.setSpacing(20)
        content_layout.setContentsMargins(24, 24, 24, 24)

        # Создаем секции
        self.create_header_section(content_layout)
        self.create_url_section(content_layout)
        self.create_tabs_section(content_layout)
        self.create_progress_section(content_layout)
        self.create_log_toggle_section(content_layout)

    def create_log_toggle_section(self, parent_layout):
        """Создает секцию с кнопкой управления логом (изначально скрыта)"""
        self.log_toggle_container = QHBoxLayout()
        self.log_toggle_container.addStretch()

        self.log_toggle_btn = QPushButton("📋 Скрыть лог")
        self.log_toggle_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #666;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 6px 12px;
                font-size: 12px;
            }
            QPushButton:hover {
                background: #f8f9fa;
                color: #333;
            }
        """)
        self.log_toggle_btn.setFixedHeight(28)
        self.log_toggle_btn.clicked.connect(self.toggle_log_from_button)
        self.log_toggle_btn.hide()  # Изначально скрыта

        self.log_toggle_container.addWidget(self.log_toggle_btn)
        parent_layout.addLayout(self.log_toggle_container)

    def on_log_first_shown(self):
        """Вызывается при первом показе лога"""
        self.log_toggle_btn.show()  # Показать кнопку управления

    def toggle_log_from_button(self):
        """Переключение лога из кнопки в интерфейсе"""
        if self.sliding_log.is_expanded:
            self.sliding_log.slide_left()
            self.log_toggle_btn.setText("📋 Показать лог")
        else:
            self.sliding_log.slide_right()
            self.log_toggle_btn.setText("📋 Скрыть лог")

    def sync_toggle_button(self):
        """Синхронизация кнопки с состоянием слайдера"""
        if self.sliding_log.is_expanded:
            self.log_toggle_btn.setText("📋 Скрыть лог")
        else:
            self.log_toggle_btn.setText("📋 Показать лог")

    def create_header_section(self, parent_layout):
        """Создает заголовок приложения"""
        title = QLabel("Excel → Google Sheets")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            color: #333;
            margin: 0;
        """)

        subtitle = QLabel("Синхронизация данных")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("""
            font-size: 14px;
            color: #666;
            margin: 0;
        """)

        parent_layout.addWidget(title)
        parent_layout.addWidget(subtitle)

    def create_url_section(self, parent_layout):
        """Создает секцию для ввода URL Google Таблицы"""
        url_container = QFrame()
        url_container.setStyleSheet("""
            QFrame {
                background: #f8f9fa;
                border: 1px solid #e9ecef;
                border-radius: 8px;
                padding: 16px;
            }
        """)

        url_layout = QVBoxLayout(url_container)
        url_layout.setSpacing(8)

        # Заголовок секции
        url_label = QLabel("🔗 Ссылка на Google Таблицу")
        url_label.setStyleSheet("font-weight: bold; color: #333;")

        # Поле ввода
        self.google_url_input = QLineEdit()
        self.google_url_input.setPlaceholderText("https://docs.google.com/spreadsheets/d/...")
        self.google_url_input.setStyleSheet("""
            QLineEdit {
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 6px;
                font-size: 13px;
                background: white;
            }
            QLineEdit:focus {
                border-color: #007bff;
            }
        """)
        self.google_url_input.setFixedHeight(36)

        url_layout.addWidget(url_label)
        url_layout.addWidget(self.google_url_input)

        parent_layout.addWidget(url_container)

    def create_tabs_section(self, parent_layout):
        """Создает секцию с табами"""
        # Контейнер для табов и кнопки скачивания
        tabs_container = QVBoxLayout()
        tabs_container.setSpacing(8)

        # Заголовок и кнопка скачивания
        tabs_header = QHBoxLayout()
        tabs_header.addStretch()

        # Кнопка скачивания
        self.download_btn = QPushButton("💾")
        self.download_btn.setEnabled(False)
        self.download_btn.setStyleSheet("""
            QPushButton {
                background: #ffc107;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 16px;
                width: 32px;
                height: 32px;
            }
            QPushButton:hover {
                background: #e0a800;
            }
            QPushButton:disabled {
                background: #ccc;
            }
        """)
        self.download_btn.setFixedSize(32, 32)
        self.download_btn.setToolTip("Скачать Google таблицу")

        tabs_header.addWidget(self.download_btn)

        # Сами табы
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #ddd;
                border-radius: 6px;
                background: white;
                padding: 16px;
            }
            QTabBar::tab {
                background: #f8f9fa;
                border: 1px solid #ddd;
                border-bottom: none;
                padding: 8px 16px;
                margin-right: 2px;
                border-radius: 6px 6px 0 0;
            }
            QTabBar::tab:selected {
                background: white;
                color: #007bff;
                font-weight: bold;
            }
            QTabBar::tab:hover:!selected {
                background: #e9ecef;
            }
        """)
        self.tabs.setFixedHeight(300)

        # Создаем табы
        single_tab = self.create_single_tab()
        batch_tab = self.create_batch_tab()

        self.tabs.addTab(single_tab, "📄 Один файл")
        self.tabs.addTab(batch_tab, "📁 Пакетная обработка")

        tabs_container.addLayout(tabs_header)
        tabs_container.addWidget(self.tabs)

        parent_layout.addLayout(tabs_container)

    def create_single_tab(self):
        """Создает таб для обработки одного файла"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(16)
        layout.setContentsMargins(16, 16, 16, 16)

        # Область для перетаскивания
        drop_container = QHBoxLayout()
        drop_container.addStretch()

        self.single_drop_area = ModernDropArea(accept_multiple=False)
        self.single_drop_area.setFixedSize(300, 80)

        drop_container.addWidget(self.single_drop_area)
        drop_container.addStretch()

        layout.addLayout(drop_container)

        # Кнопки управления
        buttons_layout = QVBoxLayout()
        buttons_layout.setSpacing(8)

        self.single_mapping_btn = QPushButton("⚙️ Настроить маппинг")
        self.single_mapping_btn.setEnabled(False)
        self.single_mapping_btn.setStyleSheet("""
            QPushButton {
                background: #6c757d;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #5a6268;
            }
            QPushButton:disabled {
                background: #ccc;
            }
        """)
        self.single_mapping_btn.setFixedHeight(36)

        self.single_process_btn = QPushButton("🚀 Начать копирование")
        self.single_process_btn.setEnabled(False)
        self.single_process_btn.setStyleSheet("""
            QPushButton {
                background: #28a745;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #218838;
            }
            QPushButton:disabled {
                background: #ccc;
            }
        """)
        self.single_process_btn.setFixedHeight(36)

        buttons_layout.addWidget(self.single_mapping_btn)
        buttons_layout.addWidget(self.single_process_btn)

        layout.addLayout(buttons_layout)
        layout.addStretch()

        return tab

    def create_batch_tab(self):
        """Создает таб для пакетной обработки"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setSpacing(12)
        layout.setContentsMargins(16, 16, 16, 16)

        # Область для перетаскивания
        drop_container = QHBoxLayout()
        drop_container.addStretch()

        self.batch_drop_area = ModernDropArea(accept_multiple=True)
        self.batch_drop_area.setFixedSize(300, 60)

        drop_container.addWidget(self.batch_drop_area)
        drop_container.addStretch()

        layout.addLayout(drop_container)

        # Список файлов
        self.files_list = QListWidget()
        self.files_list.setFixedHeight(60)
        self.files_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 4px;
                background: white;
                padding: 4px;
                font-size: 12px;
            }
            QListWidget::item {
                padding: 4px;
                margin: 1px;
                border-radius: 3px;
            }
            QListWidget::item:selected {
                background: #e3f2fd;
                color: #1976d2;
            }
            QListWidget::item:hover {
                background: #f5f5f5;
            }
        """)

        # Кнопки управления списком
        list_buttons_layout = QHBoxLayout()
        list_buttons_layout.setSpacing(8)

        self.clear_btn = QPushButton("🗑️ Очистить")
        self.clear_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #666;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 6px 12px;
                font-size: 12px;
            }
            QPushButton:hover {
                background: #f8f9fa;
            }
        """)
        self.clear_btn.setFixedHeight(28)

        self.remove_btn = QPushButton("➖ Удалить выбранные")
        self.remove_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #666;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 6px 12px;
                font-size: 12px;
            }
            QPushButton:hover {
                background: #f8f9fa;
            }
        """)
        self.remove_btn.setFixedHeight(28)

        list_buttons_layout.addWidget(self.clear_btn)
        list_buttons_layout.addWidget(self.remove_btn)
        list_buttons_layout.addStretch()

        # Основные кнопки
        main_buttons_layout = QVBoxLayout()
        main_buttons_layout.setSpacing(8)

        self.batch_mapping_btn = QPushButton("⚙️ Настроить маппинг")
        self.batch_mapping_btn.setEnabled(False)
        self.batch_mapping_btn.setStyleSheet("""
            QPushButton {
                background: #6c757d;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #5a6268;
            }
            QPushButton:disabled {
                background: #ccc;
            }
        """)
        self.batch_mapping_btn.setFixedHeight(36)

        self.batch_process_btn = QPushButton("🚀 Начать копирование")
        self.batch_process_btn.setEnabled(False)
        self.batch_process_btn.setStyleSheet("""
            QPushButton {
                background: #28a745;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: #218838;
            }
            QPushButton:disabled {
                background: #ccc;
            }
        """)
        self.batch_process_btn.setFixedHeight(36)

        main_buttons_layout.addWidget(self.batch_mapping_btn)
        main_buttons_layout.addWidget(self.batch_process_btn)

        # Собираем все вместе
        layout.addWidget(self.files_list)
        layout.addLayout(list_buttons_layout)
        layout.addLayout(main_buttons_layout)
        layout.addStretch()

        return tab

    def create_progress_section(self, parent_layout):
        """Создает секцию прогресса"""
        progress_container = QVBoxLayout()
        progress_container.setSpacing(6)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 4px;
                background: #e9ecef;
                text-align: center;
                font-weight: bold;
                color: #333;
                height: 20px;
            }
            QProgressBar::chunk {
                background: #007bff;
                border-radius: 4px;
            }
        """)
        self.progress_bar.setFixedHeight(24)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #666; font-size: 12px;")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.hide()

        progress_container.addWidget(self.progress_bar)
        progress_container.addWidget(self.status_label)

        parent_layout.addLayout(progress_container)

    def connect_signals(self):
        """Подключает сигналы к слотам"""
        # URL и загрузка
        self.google_url_input.textChanged.connect(self.check_ready_state)
        self.download_btn.clicked.connect(self.download_google_sheet)

        # Одиночный файл
        self.single_drop_area.file_dropped.connect(self.on_single_file_dropped)
        self.single_mapping_btn.clicked.connect(self.configure_single_mapping)
        self.single_process_btn.clicked.connect(self.start_single_processing)

        # Пакетная обработка
        self.batch_drop_area.files_dropped.connect(self.on_batch_files_dropped)
        self.clear_btn.clicked.connect(self.clear_batch_files)
        self.remove_btn.clicked.connect(self.remove_selected_files)
        self.batch_mapping_btn.clicked.connect(self.configure_batch_mapping)
        self.batch_process_btn.clicked.connect(self.start_batch_processing)

        # Интерфейс
        self.tabs.currentChanged.connect(self.check_ready_state)

    @handle_errors
    def check_config(self):
        """Проверяет наличие конфигурационных файлов"""
        config_path = BASE_DIR / "config.yaml"
        if not config_path.exists():
            self.log_message("⚠️ Создаю config.yaml...", "warning")
            create_sample_config(config_path)
            self.log_message("✓ Создан config.yaml", "success")

        creds_path = BASE_DIR / "credentials.json"
        if not creds_path.exists():
            self.log_message("⚠️ Требуется файл credentials.json!", "warning")

    def check_ready_state(self):
        """Проверяет готовность интерфейса и активирует кнопки"""
        has_google_url = bool(self.google_url_input.text().strip())
        current_tab = self.tabs.currentIndex()

        # Кнопка скачивания
        self.download_btn.setEnabled(has_google_url)

        if current_tab == 0:  # Одиночный файл
            has_file = self.state.single_file is not None
            self.single_mapping_btn.setEnabled(has_google_url and has_file)
            self.single_process_btn.setEnabled(
                has_google_url and has_file and self.state.single_config is not None
            )
        else:  # Пакетная обработка
            has_files = len(self.state.batch_files) > 0
            self.batch_mapping_btn.setEnabled(has_google_url and has_files)
            self.batch_process_btn.setEnabled(
                has_google_url and has_files and len(self.state.batch_mappings) > 0
            )

    @handle_errors
    def download_google_sheet(self):
        """Скачивает Google таблицу"""
        google_url = self.google_url_input.text().strip()
        if not google_url:
            return

        self.log_message("🔍 Получение списка листов...")
        sheet_names = self.logic.get_google_sheets(google_url)
        if not sheet_names:
            raise Exception("Не удалось получить листы Google Таблицы")

        dialog = DownloadDialog(sheet_names, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            selected_sheets = dialog.get_selection()
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить таблицу как",
                f"{self.logic.get_google_sheet_title()}.xlsx",
                "Excel Files (*.xlsx)"
            )
            if file_path:
                self.disable_ui()
                self.show_progress()
                self.sliding_log.clear_log()

                # НЕ создаем файл журнала для скачивания
                self.logic.start_download(
                    google_url,
                    file_path,
                    selected_sheets,
                    self.update_progress,
                    self.log_message,
                    self.on_download_finished,  # Отдельный обработчик для скачивания
                    self.on_processing_error,
                )

    def on_download_finished(self):
        """Обработчик успешного завершения скачивания"""
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("✅ Скачивание завершено!")
        self.status_label.setText("🎉 Файл успешно скачан")
        self.log_message("🎉 Скачивание выполнено успешно!", "success")

        # Скрываем прогресс через 3 секунды
        QTimer.singleShot(3000, self.hide_progress)
        self.enable_ui()

        # Показываем уведомление БЕЗ упоминания журнала
        msg = QMessageBox(self)
        msg.setWindowTitle("Скачивание завершено")
        msg.setText("🎉 Файл успешно скачан!")
        msg.setInformativeText("Google таблица была успешно сохранена на ваш компьютер.")
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def on_single_file_dropped(self, file_path: str):
        """Обработчик добавления одиночного файла"""
        self.state.single_file = file_path
        self.state.single_config = None
        self.check_ready_state()
        self.log_message(f"✓ Выбран файл: {os.path.basename(file_path)}", "success")

    @handle_errors
    def configure_single_mapping(self):
        """Настройка маппинга для одиночного файла"""
        if not self.state.single_file or not self.google_url_input.text().strip():
            return

        self.log_message("🔍 Анализ структуры файла...")
        excel_sheets = self.logic.get_excel_sheets(self.state.single_file)
        if not excel_sheets:
            raise Exception("Не удалось получить листы Excel")

        self.log_message("🔗 Подключение к Google Таблицам...")
        google_sheets = self.logic.get_google_sheets(self.google_url_input.text().strip())
        if not google_sheets:
            raise Exception("Не удалось получить листы Google")

        self.log_message("⚙️ Открытие диалога настроек...")
        dialog = MappingDialog(excel_sheets, google_sheets, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.state.single_config = dialog.get_config()
            self.log_message("✅ Маппинг настроен успешно!", "success")
            self.check_ready_state()

    def start_single_processing(self):
        """Запуск обработки одиночного файла"""
        if not self.state.single_file or not self.google_url_input.text().strip() or not self.state.single_config:
            return

        self.disable_ui()
        self.show_progress()
        self.sliding_log.clear_log()

        # Создаем файл журнала ТОЛЬКО для копирования данных
        header = [
            f"🚀 Начало обработки: {datetime.now().strftime('%H:%M:%S')}",
            f"📄 Файл: {os.path.basename(self.state.single_file)}",
            f"🔗 Таблица: {self.google_url_input.text().strip()[:50]}..."
        ]
        self.logger.open(header)

        self.logic.start_single_processing(
            self.state.single_file,
            self.google_url_input.text().strip(),
            self.state.single_config,
            self.update_progress,
            self.log_message,
            self.on_processing_finished,
            self.on_processing_error,
        )

    def start_batch_processing(self):
        """Запуск пакетной обработки"""
        if not self.state.batch_mappings or not self.google_url_input.text().strip():
            return

        self.disable_ui()
        self.show_progress()
        self.sliding_log.clear_log()

        # Создаем файл журнала ТОЛЬКО для копирования данных
        header = [
            f"🚀 Начало пакетной обработки: {datetime.now().strftime('%H:%M:%S')}",
            f"📁 Файлов к обработке: {len(self.state.batch_mappings)}",
            f"🔗 Таблица: {self.google_url_input.text().strip()[:50]}..."
        ]
        self.logger.open(header)

        self.logic.start_batch_processing(
            self.state.batch_mappings,
            self.google_url_input.text().strip(),
            self.update_progress,
            self.log_message,
            self.on_processing_finished,
            self.on_processing_error,
        )

    def on_batch_files_dropped(self, files: List[str]):
        """Обработчик добавления пакета файлов"""
        added = 0
        for file in files:
            if file not in self.state.batch_files:
                self.state.batch_files.append(file)
                item = QListWidgetItem(f"📄 {os.path.basename(file)}")
                item.setData(Qt.ItemDataRole.UserRole, file)
                self.files_list.addItem(item)
                added += 1

        self.state.batch_mappings = []  # Сбрасываем маппинги при изменении списка
        self.check_ready_state()

        if added > 0:
            self.log_message(f"✅ Добавлено файлов: {added}", "success")

    def clear_batch_files(self):
        """Очищает список файлов для пакетной обработки"""
        self.state.batch_files = []
        self.state.batch_mappings = []
        self.files_list.clear()
        self.batch_drop_area.reset()
        self.check_ready_state()
        self.log_message("🗑️ Список файлов очищен")

    def remove_selected_files(self):
        """Удаляет выбранные файлы из списка"""
        selected_items = self.files_list.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "Внимание", "Выберите файлы для удаления")
            return

        removed = 0
        for item in selected_items:
            file_path = item.data(Qt.ItemDataRole.UserRole)
            if file_path in self.state.batch_files:
                self.state.batch_files.remove(file_path)
                removed += 1
            self.files_list.takeItem(self.files_list.row(item))

        self.state.batch_mappings = []  # Сбрасываем маппинги
        if not self.state.batch_files:
            self.batch_drop_area.reset()

        self.check_ready_state()
        if removed > 0:
            self.log_message(f"➖ Удалено файлов: {removed}")

    @handle_errors
    def configure_batch_mapping(self):
        """Настройка маппинга для пакетной обработки"""
        if not self.state.batch_files or not self.google_url_input.text().strip():
            return

        self.log_message("🔗 Подключение к Google Таблицам...")
        google_sheets = self.logic.get_google_sheets(self.google_url_input.text().strip())
        if not google_sheets:
            raise Exception("Не удалось получить листы Google")

        self.log_message("⚙️ Открытие диалога пакетных настроек...")
        dialog = BatchMappingDialog(self.state.batch_files, google_sheets, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.state.batch_mappings = dialog.mappings
            self.log_message(f"✅ Настроен маппинг для {len(self.state.batch_mappings)} файлов", "success")
            self.check_ready_state()

    def start_batch_processing(self):
        """Запуск пакетной обработки"""
        if not self.state.batch_mappings or not self.google_url_input.text().strip():
            return

        self.disable_ui()
        self.show_progress()
        self.sliding_log.clear_log()

        # Создаем файл журнала ТОЛЬКО для копирования данных
        header = [
            f"🚀 Начало пакетной обработки: {datetime.now().strftime('%H:%M:%S')}",
            f"📁 Файлов к обработке: {len(self.state.batch_mappings)}",
            f"🔗 Таблица: {self.google_url_input.text().strip()[:50]}..."
        ]
        self.logger.open(header)

        self.logic.start_batch_processing(
            self.state.batch_mappings,
            self.google_url_input.text().strip(),
            self.update_progress,
            self.log_message,
            self.on_processing_finished,
            self.on_processing_error,
        )

    def disable_ui(self):
        """Отключает элементы интерфейса во время обработки"""
        self.tabs.setEnabled(False)
        self.google_url_input.setEnabled(False)
        self.download_btn.setEnabled(False)

    def enable_ui(self):
        """Включает элементы интерфейса после обработки"""
        self.tabs.setEnabled(True)
        self.google_url_input.setEnabled(True)
        self.check_ready_state()

    def show_progress(self):
        """Показывает прогресс-бар и статус"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.show()

    def hide_progress(self):
        """Скрывает прогресс-бар и статус"""
        self.progress_bar.setVisible(False)
        self.status_label.hide()

    def update_progress(self, current: int, total: int, item_name: str):
        """Обновляет прогресс-бар"""
        progress = int((current / total) * 100) if total > 0 else 0
        self.progress_bar.setValue(progress)
        self.progress_bar.setFormat(f"{progress}% - {item_name}")

        # Добавляем сообщение в лог о прогрессе
        if progress > 0:
            self.log_message(f"📊 Прогресс: {progress}% - {item_name}")

        if self.tabs.currentIndex() == 0:
            self.status_label.setText(f"📋 Обработано листов: {current}/{total}")
        else:
            self.status_label.setText(f"📁 Обработано файлов: {current}/{total}")

    def log_message(self, message: str, message_type: str = "info"):
        """Добавляет сообщение в журнал"""
        formatted = self.logger.log(message) if hasattr(self.logger,
                                                        'log_file_path') and self.logger.log_file_path else f"[{datetime.now().strftime('%H:%M:%S')}] {message}"

        # Определяем тип сообщения для правильного отображения
        if "ОШИБКА" in message or "ERROR" in message:
            message_type = "error"
        elif "⚠️" in message or "WARNING" in message:
            message_type = "warning"
        elif "✅" in message or "✓" in message or "🎉" in message:
            message_type = "success"

        # Обработка специальных типов сообщений с ссылками
        if message.startswith("📋 Ссылка:"):
            url = message.split(": ", 1)[1]
            html_message = f'{formatted.split("] ", 1)[0]}] 📋 Ссылка: <a href="{url}" style="color: #007bff; text-decoration: underline;">{url}</a>'
            self.sliding_log.add_log_message(html_message, message_type)
        elif message.startswith("💾 Сохранено:"):
            path = message.split(": ", 1)[1]
            html_message = f'{formatted.split("] ", 1)[0]}] 💾 Сохранено: <a href="file://{path}" style="color: #007bff; text-decoration: underline;">{path}</a>'
            self.sliding_log.add_log_message(html_message, message_type)
        else:
            self.sliding_log.add_log_message(formatted, message_type)

    def on_processing_finished(self):
        """Обработчик успешного завершения операции"""
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("✅ Операция завершена!")
        self.status_label.setText("🎉 Успешно завершено")
        self.logger.close()
        self.log_message("🎉 Операция выполнена успешно!", "success")

        # Скрываем прогресс через 3 секунды
        QTimer.singleShot(3000, self.hide_progress)
        self.enable_ui()

        # Показываем уведомление
        msg = QMessageBox(self)
        msg.setWindowTitle("Операция завершена")
        msg.setText("🎉 Операция выполнена успешно!")
        msg.setInformativeText(
            f"Все данные были успешно обработаны.\n"
            f"Журнал сохранен в: {self.logger.log_file_path.name if self.logger.log_file_path else 'не определено'}"
        )
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def on_processing_error(self, error_message: str):
        """Обработчик ошибки операции"""
        self.log_message(f"💥 ОШИБКА: {error_message}", "error")
        self.hide_progress()
        self.logger.close()
        self.enable_ui()

        # Показываем подробное сообщение об ошибке
        msg = QMessageBox(self)
        msg.setWindowTitle("Ошибка операции")
        msg.setText("💥 Произошла ошибка при выполнении операции")
        msg.setInformativeText(
            f"Детали ошибки:\n{error_message}\n\n"
            "Возможные причины:\n"
            "• Отсутствует подключение к интернету\n"
            "• Недостаточно прав доступа к Google Таблице\n"
            "• Поврежден или заблокирован Excel файл\n"
            "• Неверная настройка credentials.json\n"
            "• Превышены лимиты Google API"
        )
        msg.setIcon(QMessageBox.Icon.Critical)
        msg.setDetailedText(error_message)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()


def main():
    """Главная функция приложения"""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Простая настройка палитры без лишних стилей
    app.setStyleSheet("""
        QMainWindow {
            background: white;
        }
        QWidget {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            font-size: 13px;
        }
    """)

    # Создание и отображение окна
    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()