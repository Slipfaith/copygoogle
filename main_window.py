import sys
import os
from datetime import datetime
from typing import List

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QProgressBar,
    QListWidget, QListWidgetItem, QTabWidget, QMessageBox, QFileDialog,
    QDialog
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QPalette, QColor

from app_logic import AppLogic
from config import BASE_DIR, create_sample_config
from dialogs import BatchMappingDialog, MappingDialog, DownloadDialog
import styles
from widgets import ClickableTextEdit, ModernDropArea
from log_service import LogService
from state import AppState
from utils import handle_errors


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
        self.setWindowTitle("Excel to Google Sheets - Улучшенная версия")
        self.setFixedSize(550, 700)

        central_widget = QWidget()
        central_widget.setStyleSheet(styles.WINDOW_STYLE)
        self.setCentralWidget(central_widget)

        self.layout = QVBoxLayout()
        self.layout.setSpacing(20)
        self.layout.setContentsMargins(30, 30, 30, 30)

        self.create_title_section()
        self.create_url_section()
        self.create_tabs()
        self.create_progress_section()
        self.create_log_section()

        central_widget.setLayout(self.layout)

    def create_title_section(self):
        title_layout = QVBoxLayout()
        title_layout.setSpacing(5)

        title = QLabel("Excel → Google Sheets")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet(styles.TITLE_LABEL_STYLE)

        subtitle = QLabel("Улучшенное копирование данных")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet(styles.SUBTITLE_LABEL_STYLE)

        title_layout.addWidget(title)
        title_layout.addWidget(subtitle)
        self.layout.addLayout(title_layout)

    def create_url_section(self):
        url_container = QWidget()
        url_container.setStyleSheet(styles.URL_CONTAINER_STYLE)
        url_layout = QVBoxLayout()
        url_layout.setContentsMargins(15, 15, 15, 15)

        url_label = QLabel("🔗 Google Таблица")
        url_label.setStyleSheet(styles.URL_LABEL_STYLE)

        self.google_url_input = QLineEdit()
        self.google_url_input.setPlaceholderText("Вставьте ссылку на Google Таблицу")
        self.google_url_input.setStyleSheet(styles.URL_INPUT_STYLE)

        url_btns_layout = QHBoxLayout()
        self.download_btn = QPushButton("💾 Скачать таблицу")
        self.download_btn.setEnabled(False)
        self.download_btn.setStyleSheet(styles.DOWNLOAD_BUTTON_STYLE)
        url_btns_layout.addWidget(self.download_btn)
        url_btns_layout.addStretch()

        url_layout.addWidget(url_label)
        url_layout.addWidget(self.google_url_input)
        url_layout.addLayout(url_btns_layout)
        url_container.setLayout(url_layout)
        self.layout.addWidget(url_container)

    def create_tabs(self):
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(styles.TAB_WIDGET_STYLE)
        single_tab = self.create_single_tab()
        batch_tab = self.create_batch_tab()
        self.tabs.addTab(single_tab, "📄 Один файл")
        self.tabs.addTab(batch_tab, "📁 Несколько файлов")
        self.layout.addWidget(self.tabs)

    def create_single_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)

        drop_container = QHBoxLayout()
        drop_container.addStretch()
        self.single_drop_area = ModernDropArea(accept_multiple=False)
        drop_container.addWidget(self.single_drop_area)
        drop_container.addStretch()
        layout.addLayout(drop_container)

        self.single_mapping_btn = self.create_button("⚙️ Настроить маппинг", "#6c757d", "#495057")
        self.single_mapping_btn.setEnabled(False)
        self.single_process_btn = self.create_button("🚀 Начать копирование", "#28a745", "#218838", primary=True)
        self.single_process_btn.setEnabled(False)

        layout.addWidget(self.single_mapping_btn)
        layout.addWidget(self.single_process_btn)
        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_batch_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setSpacing(15)

        drop_container = QHBoxLayout()
        drop_container.addStretch()
        self.batch_drop_area = ModernDropArea(accept_multiple=True)
        drop_container.addWidget(self.batch_drop_area)
        drop_container.addStretch()
        layout.addLayout(drop_container)

        self.files_list = QListWidget()
        self.files_list.setMaximumHeight(100)
        self.files_list.setStyleSheet(styles.FILES_LIST_STYLE)
        layout.addWidget(self.files_list)

        list_btns_layout = QHBoxLayout()
        self.clear_btn = self.create_small_button("🗑️ Очистить")
        self.remove_btn = self.create_small_button("➖ Удалить выбранные")
        list_btns_layout.addWidget(self.clear_btn)
        list_btns_layout.addWidget(self.remove_btn)
        list_btns_layout.addStretch()
        layout.addLayout(list_btns_layout)

        self.batch_mapping_btn = self.create_button("⚙️ Настроить маппинг", "#6c757d", "#495057")
        self.batch_mapping_btn.setEnabled(False)
        self.batch_process_btn = self.create_button("🚀 Начать копирование", "#28a745", "#218838", primary=True)
        self.batch_process_btn.setEnabled(False)

        layout.addWidget(self.batch_mapping_btn)
        layout.addWidget(self.batch_process_btn)
        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_progress_section(self):
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet(styles.PROGRESS_BAR_STYLE)
        self.layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet(styles.STATUS_LABEL_STYLE)
        self.status_label.hide()
        self.layout.addWidget(self.status_label)

    def create_log_section(self):
        self.log_text = ClickableTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet(styles.LOG_TEXT_STYLE)
        self.log_text.hide()
        self.layout.addWidget(self.log_text)

        self.toggle_log_btn = self.create_small_button("📋 Показать журнал")
        self.layout.addWidget(self.toggle_log_btn)

    def create_button(self, text: str, color: str, hover: str, primary: bool = False) -> QPushButton:
        btn = QPushButton(text)
        if primary:
            btn.setStyleSheet(styles.primary_button(color, hover))
        else:
            btn.setStyleSheet(styles.secondary_button(color))
        return btn

    def create_small_button(self, text: str) -> QPushButton:
        btn = QPushButton(text)
        btn.setStyleSheet(styles.SMALL_BUTTON_STYLE)
        return btn

    def connect_signals(self):
        self.download_btn.clicked.connect(self.download_google_sheet)
        self.single_drop_area.file_dropped.connect(self.on_single_file_dropped)
        self.single_mapping_btn.clicked.connect(self.configure_single_mapping)
        self.single_process_btn.clicked.connect(self.start_single_processing)
        self.batch_drop_area.files_dropped.connect(self.on_batch_files_dropped)
        self.clear_btn.clicked.connect(self.clear_batch_files)
        self.remove_btn.clicked.connect(self.remove_selected_files)
        self.batch_mapping_btn.clicked.connect(self.configure_batch_mapping)
        self.batch_process_btn.clicked.connect(self.start_batch_processing)
        self.toggle_log_btn.clicked.connect(self.toggle_log)
        self.google_url_input.textChanged.connect(self.check_ready_state)
        self.tabs.currentChanged.connect(self.check_ready_state)

    def toggle_log(self):
        if self.log_text.isVisible():
            self.log_text.hide()
            self.toggle_log_btn.setText("📋 Показать журнал")
        else:
            self.log_text.show()
            self.toggle_log_btn.setText("📋 Скрыть журнал")

    @handle_errors
    def check_config(self):
        config_path = BASE_DIR / "config.yaml"
        if not config_path.exists():
            self.log_message("⚠️ Создаю config.yaml...")
            create_sample_config(config_path)
            self.log_message("✓ Создан config.yaml")

        creds_path = BASE_DIR / "credentials.json"
        if not creds_path.exists():
            self.log_message("⚠️ Нужен credentials.json!")

    def check_ready_state(self):
        has_google_url = bool(self.google_url_input.text().strip())
        current_tab = self.tabs.currentIndex()

        self.download_btn.setEnabled(has_google_url)

        if current_tab == 0:
            has_file = self.state.single_file is not None
            self.single_mapping_btn.setEnabled(has_google_url and has_file)
            self.single_process_btn.setEnabled(
                has_google_url and has_file and self.state.single_config is not None
            )
        else:
            has_files = len(self.state.batch_files) > 0
            self.batch_mapping_btn.setEnabled(has_google_url and has_files)
            self.batch_process_btn.setEnabled(
                has_google_url and has_files and len(self.state.batch_mappings) > 0
            )

    @handle_errors
    def download_google_sheet(self):
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
                self.log_text.clear()
                self.logic.start_download(
                    google_url,
                    file_path,
                    selected_sheets,
                    self.update_progress,
                    self.log_message,
                    self.on_processing_finished,
                    self.on_processing_error,
                )

    def on_single_file_dropped(self, file_path: str):
        self.state.single_file = file_path
        self.state.single_config = None
        self.check_ready_state()
        self.log_message(f"✓ Файл: {os.path.basename(file_path)}")

    @handle_errors
    def configure_single_mapping(self):
        if not self.state.single_file or not self.google_url_input.text().strip():
            return

        self.log_message("🔍 Анализ файла...")
        excel_sheets = self.logic.get_excel_sheets(self.state.single_file)
        if not excel_sheets:
            raise Exception("Не удалось получить листы Excel")

        self.log_message("🔗 Подключение к Google Таблицам...")
        google_sheets = self.logic.get_google_sheets(self.google_url_input.text().strip())
        if not google_sheets:
            raise Exception("Не удалось получить листы Google")

        self.log_message("⚙️ Открытие настроек...")
        dialog = MappingDialog(excel_sheets, google_sheets, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.state.single_config = dialog.get_config()
            self.log_message("✅ Маппинг настроен успешно!")
            self.check_ready_state()

    def start_single_processing(self):
        if not self.state.single_file or not self.google_url_input.text().strip() or not self.state.single_config:
            return

        self.disable_ui()
        self.show_progress()
        self.log_text.clear()

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

    def on_batch_files_dropped(self, files: List[str]):
        added = 0
        for file in files:
            if file not in self.state.batch_files:
                self.state.batch_files.append(file)
                item = QListWidgetItem(f"📄 {os.path.basename(file)}")
                item.setData(Qt.ItemDataRole.UserRole, file)
                self.files_list.addItem(item)
                added += 1
        self.state.batch_mappings = []
        self.check_ready_state()
        if added > 0:
            self.log_message(f"✅ Добавлено файлов: {added}")

    def clear_batch_files(self):
        self.state.batch_files = []
        self.state.batch_mappings = []
        self.files_list.clear()
        self.batch_drop_area.reset()
        self.check_ready_state()
        self.log_message("🗑️ Список файлов очищен")

    def remove_selected_files(self):
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
        self.state.batch_mappings = []
        if not self.state.batch_files:
            self.batch_drop_area.reset()
        self.check_ready_state()
        if removed > 0:
            self.log_message(f"➖ Удалено файлов: {removed}")

    @handle_errors
    def configure_batch_mapping(self):
        if not self.state.batch_files or not self.google_url_input.text().strip():
            return

        self.log_message("🔗 Подключение к Google Таблицам...")
        google_sheets = self.logic.get_google_sheets(self.google_url_input.text().strip())
        if not google_sheets:
            raise Exception("Не удалось получить листы Google")

        self.log_message("⚙️ Открытие настроек пакетного маппинга...")
        dialog = BatchMappingDialog(self.state.batch_files, google_sheets, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.state.batch_mappings = dialog.mappings
            self.log_message(f"✅ Настроено маппинг для {len(self.state.batch_mappings)} файлов")
            self.check_ready_state()

    def start_batch_processing(self):
        if not self.state.batch_mappings or not self.google_url_input.text().strip():
            return

        self.disable_ui()
        self.show_progress()
        self.log_text.clear()

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
        self.tabs.setEnabled(False)
        self.google_url_input.setEnabled(False)
        self.download_btn.setEnabled(False)
        self.single_mapping_btn.setEnabled(False)
        self.single_process_btn.setEnabled(False)
        self.batch_mapping_btn.setEnabled(False)
        self.batch_process_btn.setEnabled(False)

    def enable_ui(self):
        self.tabs.setEnabled(True)
        self.google_url_input.setEnabled(True)
        self.check_ready_state()

    def show_progress(self):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.show()
        if not self.log_text.isVisible():
            self.toggle_log()

    def hide_progress(self):
        self.progress_bar.setVisible(False)
        self.status_label.hide()

    def update_progress(self, current: int, total: int, item_name: str):
        progress = int((current / total) * 100) if total > 0 else 0
        self.progress_bar.setValue(progress)
        self.progress_bar.setFormat(f"{progress}% - {item_name}")
        if self.tabs.currentIndex() == 0:
            self.status_label.setText(f"📋 Листов: {current}/{total}")
        else:
            self.status_label.setText(f"📁 Файлов: {current}/{total}")

    def log_message(self, message: str):
        formatted = self.logger.log(message)
        if message.startswith("📋 Ссылка:"):
            url = message.split(": ", 1)[1]
            html_message = f'{formatted.split("] ",1)[0]}] 📋 Ссылка: <a href="{url}">{url}</a>'
            self.log_text.append(html_message)
        elif message.startswith("💾 Сохранено:"):
            path = message.split(": ", 1)[1]
            html_message = f'{formatted.split("] ",1)[0]}] 💾 Сохранено: <a href="file://{path}">{path}</a>'
            self.log_text.append(html_message)
        else:
            self.log_text.append(formatted)
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def on_processing_finished(self):
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("✅ Готово!")
        self.status_label.setText("🎉 Завершено успешно")
        self.logger.close()
        self.log_message("🎉 Обработка завершена успешно!")
        QTimer.singleShot(3000, self.hide_progress)
        self.enable_ui()
        msg = QMessageBox(self)
        msg.setWindowTitle("Успешно!")
        msg.setText("🎉 Операция завершена успешно!")
        msg.setInformativeText(
            f"Все операции были успешно выполнены.\nЛог сохранен в: {self.logger.log_file_path.name if self.logger.log_file_path else 'неизвестно'}"
        )
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def on_processing_error(self, error_message: str):
        self.log_message(f"💥 КРИТИЧЕСКАЯ ОШИБКА: {error_message}")
        self.hide_progress()
        self.logger.close()
        self.enable_ui()
        msg = QMessageBox(self)
        msg.setWindowTitle("Ошибка обработки")
        msg.setText("💥 Произошла ошибка при обработке")
        msg.setInformativeText(
            f"Детали ошибки:\n{error_message}\n\nВозможные причины:\n"
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
    app.setStyle("Fusion")

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
