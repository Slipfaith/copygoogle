import sys
import os
import subprocess
import platform
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, List, Tuple

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QProgressBar, QLabel, QFrame,
    QMessageBox, QFileDialog, QLineEdit, QDialog, QDialogButtonBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox,
    QGroupBox, QSpinBox, QTabWidget, QListWidget, QListWidgetItem,
    QGraphicsDropShadowEffect, QSizePolicy, QCheckBox
)
from PySide6.QtCore import Qt, QThread, Signal, QMimeData, QTimer, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QPalette, QColor, QFont, QIcon, QTextCursor, QTextCharFormat

from business.processor import ExcelToGoogleSheets
from config import BASE_DIR, create_sample_config
from dialogs import BatchMappingDialog, MappingDialog, DropArea, DownloadDialog

BASE_DIR = Path(__file__).resolve().parent


class ClickableTextEdit(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setOpenLinks(False)
        self.anchorClicked.connect(self.handle_click)
        
    def handle_click(self, url):
        if url.startswith("file://"):
            path = url.replace("file://", "")
            if os.path.exists(path):
                if platform.system() == 'Windows':
                    subprocess.run(['explorer', '/select,', path])
                elif platform.system() == 'Darwin':
                    subprocess.run(['open', '-R', path])
                else:
                    subprocess.run(['xdg-open', os.path.dirname(path)])


class ModernDropArea(QWidget):
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
                create_backup = self.kwargs.get('create_backup', False)

                self.processor.update_config(
                    sheet_mapping=config['sheet_mapping'],
                    column_mapping=config['column_mapping'],
                    start_row=config['start_row']
                )

                self.log_message.emit("Подключение к Google Таблицам...")
                self.processor.connect_to_google_sheets(google_sheet_url)

                if create_backup:
                    self.processor.backup_google_sheet(log_callback=self.log_message.emit)

                self.processor.process_excel_file(
                    excel_path,
                    progress_callback=self.progress_update.emit,
                    log_callback=self.log_message.emit
                )

            elif self.mode == "batch":
                file_mappings = self.kwargs['file_mappings']
                google_sheet_url = self.kwargs['google_sheet_url']
                create_backup = self.kwargs.get('create_backup', False)

                if create_backup:
                    self.processor.connect_to_google_sheets(google_sheet_url)
                    self.processor.backup_google_sheet(log_callback=self.log_message.emit)

                self.processor.process_multiple_excel_files(
                    file_mappings,
                    google_sheet_url,
                    progress_callback=self.progress_update.emit,
                    log_callback=self.log_message.emit
                )

            elif self.mode == "download":
                google_sheet_url = self.kwargs['google_sheet_url']
                save_path = self.kwargs['save_path']
                sheet_names = self.kwargs.get('sheet_names')

                self.log_message.emit("Подключение к Google Таблицам...")
                self.processor.connect_to_google_sheets(google_sheet_url)

                self.processor.download_google_sheet(
                    save_path,
                    sheet_names=sheet_names,
                    log_callback=self.log_message.emit
                )

            self.finished_successfully.emit()

        except Exception as e:
            self.error_occurred.emit(str(e))


class MainWindow(QMainWindow):
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
        self.setFixedSize(550, 700)

        central_widget = QWidget()
        central_widget.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            }
        """)
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

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

        url_btns_layout = QHBoxLayout()
        
        self.download_btn = QPushButton("💾 Скачать таблицу")
        self.download_btn.setEnabled(False)
        self.download_btn.clicked.connect(self.download_google_sheet)
        self.download_btn.setStyleSheet("""
            QPushButton {
                background-color: #17a2b8;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                background-color: #138496;
            }
            QPushButton:disabled {
                background-color: #e9ecef;
                color: #6c757d;
            }
        """)
        
        url_btns_layout.addWidget(self.download_btn)
        url_btns_layout.addStretch()

        url_layout.addWidget(url_label)
        url_layout.addWidget(self.google_url_input)
        url_layout.addLayout(url_btns_layout)
        url_container.setLayout(url_layout)
        layout.addWidget(url_container)

        backup_container = QWidget()
        backup_layout = QHBoxLayout()
        backup_layout.setContentsMargins(0, 0, 0, 0)
        
        self.backup_checkbox = QCheckBox("🔒 Создать резервную копию перед вставкой данных")
        self.backup_checkbox.setStyleSheet("""
            QCheckBox {
                color: #495057;
                font-size: 13px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
        """)
        self.backup_checkbox.setChecked(True)
        
        backup_layout.addWidget(self.backup_checkbox)
        backup_layout.addStretch()
        backup_container.setLayout(backup_layout)
        layout.addWidget(backup_container)

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

        single_tab = QWidget()
        single_layout = QVBoxLayout()
        single_layout.setSpacing(15)

        drop_container = QHBoxLayout()
        drop_container.addStretch()
        self.single_drop_area = ModernDropArea(accept_multiple=False)
        self.single_drop_area.file_dropped.connect(self.on_single_file_dropped)
        drop_container.addWidget(self.single_drop_area)
        drop_container.addStretch()
        single_layout.addLayout(drop_container)

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

        batch_tab = QWidget()
        batch_layout = QVBoxLayout()
        batch_layout.setSpacing(15)

        batch_drop_container = QHBoxLayout()
        batch_drop_container.addStretch()
        self.batch_drop_area = ModernDropArea(accept_multiple=True)
        self.batch_drop_area.files_dropped.connect(self.on_batch_files_dropped)
        batch_drop_container.addWidget(self.batch_drop_area)
        batch_drop_container.addStretch()
        batch_layout.addLayout(batch_drop_container)

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

        list_btns_layout = QHBoxLayout()

        clear_btn = self.create_small_button("🗑️ Очистить")
        clear_btn.clicked.connect(self.clear_batch_files)

        remove_btn = self.create_small_button("➖ Удалить выбранные")
        remove_btn.clicked.connect(self.remove_selected_files)

        list_btns_layout.addWidget(clear_btn)
        list_btns_layout.addWidget(remove_btn)
        list_btns_layout.addStretch()
        batch_layout.addLayout(list_btns_layout)

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

        self.tabs.addTab(single_tab, "📄 Один файл")
        self.tabs.addTab(batch_tab, "📁 Несколько файлов")
        layout.addWidget(self.tabs)

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

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("""
            color: #6c757d;
            font-size: 12px;
        """)
        self.status_label.hide()
        layout.addWidget(self.status_label)

        self.log_text = ClickableTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
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

        self.toggle_log_btn = self.create_small_button("📋 Показать журнал")
        self.toggle_log_btn.clicked.connect(self.toggle_log)
        layout.addWidget(self.toggle_log_btn)

        central_widget.setLayout(layout)

        self.google_url_input.textChanged.connect(self.check_ready_state)
        self.tabs.currentChanged.connect(self.check_ready_state)

        self.check_config()

    def create_button(self, text: str, color: str, hover_color: str, primary: bool = False) -> QPushButton:
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
        if self.log_text.isVisible():
            self.log_text.hide()
            self.toggle_log_btn.setText("📋 Показать журнал")
        else:
            self.log_text.show()
            self.toggle_log_btn.setText("📋 Скрыть журнал")

    def check_config(self):
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
        has_google_url = bool(self.google_url_input.text().strip())
        current_tab = self.tabs.currentIndex()

        self.download_btn.setEnabled(has_google_url)

        if current_tab == 0:
            has_file = self.single_file is not None
            self.single_mapping_btn.setEnabled(has_google_url and has_file)
            self.single_process_btn.setEnabled(has_google_url and has_file and self.single_config is not None)
        else:
            has_files = len(self.batch_files) > 0
            self.batch_mapping_btn.setEnabled(has_google_url and has_files)
            self.batch_process_btn.setEnabled(has_google_url and has_files and len(self.batch_mappings) > 0)

    def download_google_sheet(self):
        google_url = self.google_url_input.text().strip()
        if not google_url:
            return

        try:
            self.log_message("🔍 Получение списка листов...")
            self.processor.connect_to_google_sheets(google_url)
            sheet_names = self.processor.get_google_sheets()
            
            if not sheet_names:
                raise Exception("Не удалось получить листы Google Таблицы")

            dialog = DownloadDialog(sheet_names, self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                selected_sheets = dialog.get_selection()
                
                file_path, _ = QFileDialog.getSaveFileName(
                    self,
                    "Сохранить таблицу как",
                    f"{self.processor.google_sheet.title}.xlsx",
                    "Excel Files (*.xlsx)"
                )
                
                if file_path:
                    self.disable_ui()
                    self.show_progress()
                    self.log_text.clear()
                    
                    self.worker_thread = WorkerThread(
                        mode="download",
                        google_sheet_url=google_url,
                        save_path=file_path,
                        sheet_names=selected_sheets
                    )
                    self.connect_worker_signals()
                    self.worker_thread.start()

        except Exception as e:
            self.log_message(f"❌ Ошибка: {e}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось подключиться к таблице:\n{str(e)}")

    def on_single_file_dropped(self, file_path: str):
        self.single_file = file_path
        self.single_config = None
        self.check_ready_state()
        self.log_message(f"✓ Файл: {os.path.basename(file_path)}")

    def configure_single_mapping(self):
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
            config=self.single_config,
            create_backup=self.backup_checkbox.isChecked()
        )
        self.connect_worker_signals()
        self.worker_thread.start()

    def on_batch_files_dropped(self, files: List[str]):
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
        self.batch_files = []
        self.batch_mappings = []
        self.files_list.clear()
        self.batch_drop_area.reset()
        self.check_ready_state()
        self.log_message("🗑️ Список файлов очищен")

    def remove_selected_files(self):
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
            google_sheet_url=self.google_url_input.text().strip(),
            create_backup=self.backup_checkbox.isChecked()
        )
        self.connect_worker_signals()
        self.worker_thread.start()

    def connect_worker_signals(self):
        self.worker_thread.progress_update.connect(self.update_progress)
        self.worker_thread.log_message.connect(self.log_message)
        self.worker_thread.finished_successfully.connect(self.on_processing_finished)
        self.worker_thread.error_occurred.connect(self.on_processing_error)

    def disable_ui(self):
        self.tabs.setEnabled(False)
        self.google_url_input.setEnabled(False)
        self.download_btn.setEnabled(False)
        self.backup_checkbox.setEnabled(False)
        self.single_mapping_btn.setEnabled(False)
        self.single_process_btn.setEnabled(False)
        self.batch_mapping_btn.setEnabled(False)
        self.batch_process_btn.setEnabled(False)

    def enable_ui(self):
        self.tabs.setEnabled(True)
        self.google_url_input.setEnabled(True)
        self.backup_checkbox.setEnabled(True)
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

    def open_log_file(self, header_lines):
        logs_dir = BASE_DIR / "logs"
        logs_dir.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file_path = logs_dir / f"log_{timestamp}.txt"
        self.log_file = open(self.log_file_path, "w", encoding="utf-8")
        for line in header_lines:
            self.log_file.write(line + "\n")
        self.log_file.write("\n")

    def close_log_file(self):
        if self.log_file:
            self.log_file.close()
            self.log_file = None

    def update_progress(self, current: int, total: int, item_name: str):
        progress = int((current / total) * 100) if total > 0 else 0
        self.progress_bar.setValue(progress)

        if self.tabs.currentIndex() == 0:
            self.progress_bar.setFormat(f"{progress}% - {item_name}")
            self.status_label.setText(f"📋 Листов: {current}/{total}")
        else:
            self.progress_bar.setFormat(f"{progress}% - {item_name}")
            self.status_label.setText(f"📁 Файлов: {current}/{total}")

    def log_message(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        
        if message.startswith("📋 Ссылка:"):
            url = message.split(": ", 1)[1]
            html_message = f'[{timestamp}] 📋 Ссылка: <a href="{url}">{url}</a>'
            self.log_text.append(html_message)
        elif message.startswith("💾 Сохранено:"):
            path = message.split(": ", 1)[1]
            html_message = f'[{timestamp}] 💾 Сохранено: <a href="file://{path}">{path}</a>'
            self.log_text.append(html_message)
        else:
            self.log_text.append(formatted_message)

        if self.log_file:
            self.log_file.write(formatted_message + "\n")
            self.log_file.flush()

        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def on_processing_finished(self):
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("✅ Готово!")
        self.status_label.setText("🎉 Завершено успешно")

        self.close_log_file()
        self.log_message("🎉 Обработка завершена успешно!")

        QTimer.singleShot(3000, self.hide_progress)
        self.enable_ui()

        msg = QMessageBox(self)
        msg.setWindowTitle("Успешно!")
        msg.setText("🎉 Операция завершена успешно!")
        msg.setInformativeText(
            f"Все операции были успешно выполнены.\n"
            f"Лог сохранен в: {self.log_file_path.name if self.log_file_path else 'неизвестно'}"
        )
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def on_processing_error(self, error_message: str):
        self.log_message(f"💥 КРИТИЧЕСКАЯ ОШИБКА: {error_message}")
        self.hide_progress()
        self.close_log_file()
        self.enable_ui()

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