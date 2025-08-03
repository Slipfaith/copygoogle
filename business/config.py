from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List
import yaml
import os

# Базовая директория проекта
BASE_DIR = Path(__file__).resolve().parent.parent


@dataclass
class Config:
    """Конфигурация приложения."""
    excel_path: str
    google_sheet_id: str
    credentials_path: str
    sheet_mapping: Dict[str, str]
    column_mapping: Dict[str, List[str]]
    start_row: int = 1


def load_config(config_path: str) -> Config:
    """Загрузка конфигурации из YAML файла."""
    if not os.path.exists(config_path):
        return Config(
            excel_path='',
            google_sheet_id='',
            credentials_path='',
            sheet_mapping={},
            column_mapping={'source': ['A'], 'target': ['A']},
            start_row=1,
        )

    with open(config_path, 'r', encoding='utf-8') as f:
        data = yaml.safe_load(f) or {}

    return Config(
        excel_path=data.get('excel_path', ''),
        google_sheet_id=data.get('google_sheet_id', ''),
        credentials_path=data.get('credentials_path', ''),
        sheet_mapping=data.get('sheet_mapping', {}),
        column_mapping=data.get('column_mapping', {'source': ['A'], 'target': ['A']}),
        start_row=data.get('start_row', 1),
    )


def create_sample_config(path: str | None = None) -> None:
    """Создание примера конфигурационного файла."""
    if path is None:
        path = BASE_DIR / "config.yaml"
    else:
        path = Path(path)
        if not path.is_absolute():
            path = BASE_DIR / path

    sample_config = {
        'credentials_path': 'credentials.json',
        'sheet_mapping': {
            'Sheet1': 'Лист1',
            'Sheet2': 'Лист2'
        },
        'column_mapping': {
            'source': ['A', 'C', 'E'],
            'target': ['B', 'D', 'F']
        },
        'start_row': 2
    }

    with open(str(path), 'w', encoding='utf-8') as f:
        yaml.dump(sample_config, f, allow_unicode=True, default_flow_style=False)

    print(f"Создан пример конфигурационного файла: {path}")
