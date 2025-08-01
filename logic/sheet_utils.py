from typing import List, Dict, Optional, Callable

import gspread
from openpyxl.utils import get_column_letter


def resolve_excel_columns(sheet, columns: List[str]) -> List[str]:
    """Преобразование номеров или заголовков Excel в буквы столбцов."""
    header_map = {str(cell.value).strip().lower(): cell.column_letter for cell in sheet[1] if cell.value is not None}
    result: List[str] = []
    for col in columns:
        col_str = str(col).strip()
        if not col_str:
            continue
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
    """Преобразование номеров или заголовков Google в буквы столбцов."""
    headers = worksheet.row_values(1)
    header_map = {str(val).strip().lower(): get_column_letter(i + 1) for i, val in enumerate(headers) if val}
    result: List[str] = []
    for col in columns:
        col_str = str(col).strip()
        if not col_str:
            continue
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
    """Копирование данных из Excel листа в Google Worksheet."""
    source_cols = resolve_excel_columns(excel_sheet, column_mapping['source'])
    target_cols = resolve_google_columns(google_worksheet, column_mapping['target'])

    if len(source_cols) != len(target_cols):
        raise ValueError("Количество исходных и целевых колонок должно совпадать")

    excel_data = []
    max_row = excel_sheet.max_row

    for row_idx in range(start_row, max_row + 1):
        row_data = []
        has_data = False

        for source_col in source_cols:
            cell_value = excel_sheet[f"{source_col}{row_idx}"].value
            if cell_value is not None:
                has_data = True
            row_data.append(cell_value if cell_value is not None else '')

        if has_data:
            excel_data.append(row_data)

    if not excel_data:
        if log_callback:
            log_callback("Нет данных для копирования")
        return 0

    updates = []
    for row_offset, row_data in enumerate(excel_data):
        google_row = start_row + row_offset
        for value, target_col in zip(row_data, target_cols):
            cell_address = f"{target_col}{google_row}"
            updates.append({'range': cell_address, 'values': [[value]]})

    if updates:
        batch_size = 1000
        for i in range(0, len(updates), batch_size):
            batch = updates[i:i + batch_size]
            google_worksheet.batch_update(batch, value_input_option='USER_ENTERED')
        if log_callback:
            log_callback(f"Обновлено ячеек: {len(updates)}")

    return len(excel_data)
