from typing import List, Dict, Optional, Callable
import logging

import gspread
from openpyxl.utils import get_column_letter, column_index_from_string


def resolve_excel_columns(sheet, columns: List[str]) -> List[str]:
    """Преобразование номеров, заголовков или диапазонов Excel в буквы столбцов."""
    # Кэшируем заголовки только один раз
    if not hasattr(resolve_excel_columns, '_header_cache'):
        resolve_excel_columns._header_cache = {}

    sheet_id = id(sheet)
    if sheet_id not in resolve_excel_columns._header_cache:
        header_map = {}
        for cell in sheet[1]:
            if cell.value is not None:
                header_map[str(cell.value).strip().lower()] = cell.column_letter
        resolve_excel_columns._header_cache[sheet_id] = header_map

    header_map = resolve_excel_columns._header_cache[sheet_id]
    result: List[str] = []

    for col in columns:
        col_str = str(col).strip()
        if not col_str:
            continue

        # Поддержка диапазонов вида "A-D" или "1-4"
        range_found = False
        for delim in ("-", ":"):
            if delim in col_str:
                start, end = col_str.split(delim, 1)
                start_letter = resolve_excel_columns(sheet, [start])[0]
                end_letter = resolve_excel_columns(sheet, [end])[0]
                start_idx = column_index_from_string(start_letter)
                end_idx = column_index_from_string(end_letter)
                step = 1 if end_idx >= start_idx else -1
                for idx in range(start_idx, end_idx + step, step):
                    result.append(get_column_letter(idx))
                range_found = True
                break

        if not range_found:
            if col_str.isdigit():
                result.append(get_column_letter(int(col_str)))
            elif col_str.isalpha():
                result.append(col_str.upper())
            else:
                key = col_str.lower()
                if key in header_map:
                    result.append(header_map[key])
                else:
                    raise ValueError(f"Заголовок '{col}' не найден в Excel листе")
    return result


def resolve_google_columns(worksheet, columns: List[str]) -> List[str]:
    """Преобразование номеров, заголовков или диапазонов Google в буквы столбцов."""
    # Кэшируем заголовки Google Sheets
    if not hasattr(resolve_google_columns, '_header_cache'):
        resolve_google_columns._header_cache = {}

    sheet_id = id(worksheet)
    if sheet_id not in resolve_google_columns._header_cache:
        headers = worksheet.row_values(1)
        header_map = {}
        for i, val in enumerate(headers):
            if val:
                header_map[str(val).strip().lower()] = get_column_letter(i + 1)
        resolve_google_columns._header_cache[sheet_id] = header_map

    header_map = resolve_google_columns._header_cache[sheet_id]
    result: List[str] = []

    for col in columns:
        col_str = str(col).strip()
        if not col_str:
            continue

        # Поддержка диапазонов вида "A-D" или "1-4"
        range_found = False
        for delim in ("-", ":"):
            if delim in col_str:
                start, end = col_str.split(delim, 1)
                start_letter = resolve_google_columns(worksheet, [start])[0]
                end_letter = resolve_google_columns(worksheet, [end])[0]
                start_idx = column_index_from_string(start_letter)
                end_idx = column_index_from_string(end_letter)
                step = 1 if end_idx >= start_idx else -1
                for idx in range(start_idx, end_idx + step, step):
                    result.append(get_column_letter(idx))
                range_found = True
                break

        if not range_found:
            if col_str.isdigit():
                result.append(get_column_letter(int(col_str)))
            elif col_str.isalpha():
                result.append(col_str.upper())
            else:
                key = col_str.lower()
                if key in header_map:
                    result.append(header_map[key])
                else:
                    raise ValueError(f"Заголовок '{col}' не найден в Google листе")
    return result


def copy_sheet_data(
        excel_sheet,
        google_worksheet,
        column_mapping: Dict[str, List[str]],
        start_row: int,
        log_callback: Optional[Callable[[str], None]] = None
) -> int:
    """Оптимизированное копирование данных из Excel листа в Google Worksheet."""

    if log_callback:
        log_callback("Подготовка данных...")

    # Получаем колонки
    source_cols = resolve_excel_columns(excel_sheet, column_mapping['source'])
    target_cols = resolve_google_columns(google_worksheet, column_mapping['target'])

    if len(source_cols) != len(target_cols):
        raise ValueError("Количество исходных и целевых колонок должно совпадать")

    # ОПТИМИЗАЦИЯ 1: Читаем все данные за один раз пакетом
    if log_callback:
        log_callback("Чтение данных из Excel...")

    max_row = excel_sheet.max_row
    if max_row < start_row:
        if log_callback:
            log_callback("Нет данных для копирования")
        return 0

    # Читаем данные батчами для оптимизации
    excel_data = []

    # Определяем диапазон для чтения
    min_col_idx = min(column_index_from_string(col) for col in source_cols)
    max_col_idx = max(column_index_from_string(col) for col in source_cols)
    min_col_letter = get_column_letter(min_col_idx)
    max_col_letter = get_column_letter(max_col_idx)

    # Читаем весь блок данных за раз
    data_range = f"{min_col_letter}{start_row}:{max_col_letter}{max_row}"

    try:
        # Получаем все ячейки в диапазоне за одну операцию
        cell_range = excel_sheet[data_range]

        # Создаем индексы для нужных колонок
        source_col_indices = [column_index_from_string(col) - min_col_idx for col in source_cols]

        rows_processed = 0
        for row_cells in cell_range:
            if not isinstance(row_cells, tuple):
                row_cells = (row_cells,)

            row_data = []
            has_data = False

            # Извлекаем только нужные колонки
            for col_idx in source_col_indices:
                if col_idx < len(row_cells):
                    cell_value = row_cells[col_idx].value
                    if cell_value is not None:
                        has_data = True
                    row_data.append(cell_value if cell_value is not None else '')
                else:
                    row_data.append('')

            if has_data:
                excel_data.append(row_data)

            rows_processed += 1

            # Логирование прогресса каждые 50 строк
            if rows_processed % 50 == 0 and log_callback:
                log_callback(f"Обработано строк Excel: {rows_processed}")

    except Exception as e:
        if log_callback:
            log_callback(f"Ошибка при чтении Excel: {e}")
        raise

    if not excel_data:
        if log_callback:
            log_callback("Нет данных для копирования")
        return 0

    if log_callback:
        log_callback(f"Подготовлено {len(excel_data)} строк данных")

    # ОПТИМИЗАЦИЯ 2: Записываем в Google Sheets одной операцией
    if log_callback:
        log_callback("Запись данных в Google Sheets...")

    try:
        # Определяем диапазон обновления в Google Sheet
        col_numbers = [column_index_from_string(col) for col in target_cols]
        min_col = min(col_numbers)
        max_col = max(col_numbers)
        start_col_letter = get_column_letter(min_col)
        end_col_letter = get_column_letter(max_col)
        target_range = f"{start_col_letter}{start_row}:{end_col_letter}{start_row + len(excel_data) - 1}"

        # Подготавливаем данные для Google Sheets
        num_cols = max_col - min_col + 1
        index_map = [column_index_from_string(col) - min_col for col in target_cols]

        values = []
        for row in excel_data:
            row_values = [''] * num_cols
            for value, idx in zip(row, index_map):
                row_values[idx] = value if value is not None else ''
            values.append(row_values)

        # КРИТИЧЕСКАЯ ОПТИМИЗАЦИЯ: Одна операция записи вместо множества
        if log_callback:
            log_callback(f"Обновление диапазона {target_range}...")

        google_worksheet.update(
            target_range,
            values,
            value_input_option='USER_ENTERED'
        )

        if log_callback:
            log_callback(f"✓ Успешно обновлено {len(excel_data)} строк, {len(target_cols)} колонок")

        return len(excel_data)

    except Exception as e:
        if log_callback:
            log_callback(f"Ошибка при записи в Google Sheets: {e}")
        raise


# Функция для очистки кэша (вызывать между обработкой разных файлов)
def clear_column_cache():
    """Очистка кэша заголовков колонок."""
    if hasattr(resolve_excel_columns, '_header_cache'):
        resolve_excel_columns._header_cache.clear()
    if hasattr(resolve_google_columns, '_header_cache'):
        resolve_google_columns._header_cache.clear()