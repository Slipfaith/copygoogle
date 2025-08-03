from typing import List
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QListWidget, QListWidgetItem, QCheckBox, QDialogButtonBox,
    QGroupBox, QRadioButton, QButtonGroup
)
from PySide6.QtCore import Qt


class DownloadDialog(QDialog):
    def __init__(self, sheet_names: List[str], parent=None):
        super().__init__(parent)
        self.sheet_names = sheet_names
        self.selected_sheets = []
        self.download_all = True

        self.setWindowTitle("Скачать Google таблицу")
        self.setModal(True)
        self.resize(400, 450)
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
            }
            QPushButton:hover {
                background-color: #0052a3;
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
            QListWidget {
                border: 1px solid #dee2e6;
                border-radius: 6px;
                background-color: #f8f9fa;
                padding: 5px;
            }
            QListWidget::item {
                padding: 5px;
                border-radius: 3px;
            }
            QListWidget::item:selected {
                background-color: #e7f3ff;
                color: #0066cc;
            }
        """)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(20)

        title = QLabel("💾 Выберите что скачать")
        title.setStyleSheet("font-size: 18px; font-weight: 600; color: #0066cc;")
        layout.addWidget(title)

        download_group = QGroupBox("Опции скачивания")
        download_layout = QVBoxLayout()

        self.radio_all = QRadioButton("📊 Скачать всю таблицу")
        self.radio_selected = QRadioButton("📋 Скачать выбранные листы")

        self.button_group = QButtonGroup()
        self.button_group.addButton(self.radio_all)
        self.button_group.addButton(self.radio_selected)

        self.radio_all.setChecked(True)
        self.radio_all.toggled.connect(self.on_radio_toggled)

        download_layout.addWidget(self.radio_all)
        download_layout.addWidget(self.radio_selected)

        self.sheets_list = QListWidget()
        self.sheets_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.sheets_list.setEnabled(False)

        for sheet_name in self.sheet_names:
            item = QListWidgetItem(f"📄 {sheet_name}")
            item.setData(Qt.ItemDataRole.UserRole, sheet_name)
            self.sheets_list.addItem(item)

        download_layout.addWidget(self.sheets_list)

        select_btns_layout = QHBoxLayout()

        select_all_btn = QPushButton("Выбрать все")
        select_all_btn.clicked.connect(self.select_all_sheets)
        select_all_btn.setStyleSheet("QPushButton { background-color: #28a745; }")

        clear_btn = QPushButton("Снять выделение")
        clear_btn.clicked.connect(self.clear_selection)
        clear_btn.setStyleSheet("QPushButton { background-color: #6c757d; }")

        select_btns_layout.addWidget(select_all_btn)
        select_btns_layout.addWidget(clear_btn)
        select_btns_layout.addStretch()

        download_layout.addLayout(select_btns_layout)
        download_group.setLayout(download_layout)
        layout.addWidget(download_group)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        ok_btn = buttons.button(QDialogButtonBox.StandardButton.Ok)
        ok_btn.setText("💾 Скачать")
        ok_btn.setStyleSheet("QPushButton { background-color: #0066cc; min-width: 100px; }")

        cancel_btn = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        cancel_btn.setText("✕ Отмена")
        cancel_btn.setStyleSheet("QPushButton { background-color: #dc3545; }")

        layout.addWidget(buttons)
        self.setLayout(layout)

    def on_radio_toggled(self):
        self.download_all = self.radio_all.isChecked()
        self.sheets_list.setEnabled(not self.download_all)

    def select_all_sheets(self):
        for i in range(self.sheets_list.count()):
            self.sheets_list.item(i).setSelected(True)

    def clear_selection(self):
        self.sheets_list.clearSelection()

    def get_selection(self):
        if self.download_all:
            return None
        else:
            selected_items = self.sheets_list.selectedItems()
            return [item.data(Qt.ItemDataRole.UserRole) for item in selected_items]