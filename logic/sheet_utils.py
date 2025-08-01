from typing import List, Dict, Optional, Callable
import logging

import gspread
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import PatternFill, Font


def resolve_excel_columns(sheet, columns: List[str]) -> List[str]:
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


def rgb_to_hex(rgb_color):
    if not rgb_color or rgb_color == "00000000":
        return None

    if len(rgb_color) == 8:
        rgb_color = rgb_color[2:]

    if not rgb_color.startswith('#'):
        rgb_color = '#' + rgb_color

    return rgb_color


def get_cell_formatting(cell):
    formatting = {}

    try:
        if cell.fill and cell.fill.fgColor and hasattr(cell.fill.fgColor, 'rgb'):
            bg_color = rgb_to_hex(str(cell.fill.fgColor.rgb))
            if bg_color and bg_color != '#FFFFFF':
                formatting['backgroundColor'] = {
                    'red': int(bg_color[1:3], 16) / 255,
                    'green': int(bg_color[3:5], 16) / 255,
                    'blue': int(bg_color[5:7], 16) / 255
                }

        if cell.font:
            font_format = {}

            if cell.font.color and hasattr(cell.font.color, 'rgb'):
                text_color = rgb_to_hex(str(cell.font.color.rgb))
                if text_color and text_color != '#000000':
                    font_format['foregroundColor'] = {
                        'red': int(text_color[1:3], 16) / 255,
                        'green': int(text_color[3:5], 16) / 255,
                        'blue': int(text_color[5:7], 16) / 255
                    }

            if cell.font.bold:
                font_format['bold'] = True

            if cell.font.italic:
                font_format['italic'] = True

            if cell.font.size:
                font_format['fontSize'] = int(cell.font.size)

            if font_format:
                formatting['textFormat'] = font_format

        if cell.alignment:
            text_format = formatting.get('textFormat', {})

            if cell.alignment.horizontal:
                alignment_map = {
                    'left': 'LEFT',
                    'center': 'CENTER',
                    'right': 'RIGHT'
                }
                h_align = alignment_map.get(cell.alignment.horizontal)
                if h_align:
                    text_format['horizontalAlignment'] = h_align

            if cell.alignment.vertical:
                v_alignment_map = {
                    'top': 'TOP',
                    'center': 'MIDDLE',
                    'bottom': 'BOTTOM'
                }
                v_align = v_alignment_map.get(cell.alignment.vertical)
                if v_align:
                    text_format['verticalAlignment'] = v_align

            if text_format:
                formatting['textFormat'] = text_format

    except Exception as e:
        pass

    return formatting if formatting else None


def convert_excel_formula_to_google(formula: str) -> str:
    if not formula or not formula.startswith('='):
        return formula

    replacements = {
        'ПРОПИСН': 'UPPER',
        'СТРОЧН': 'LOWER',
        'ПРОПНАЧ': 'PROPER',
        'СЦЕПИТЬ': 'CONCATENATE',
        'ДЛСТР': 'LEN',
        'ЛЕВСИМВ': 'LEFT',
        'ПРАВСИМВ': 'RIGHT',
        'ПСТР': 'MID',
        'НАЙТИ': 'FIND',
        'ЗАМЕНИТЬ': 'SUBSTITUTE',
        'ОКРУГЛ': 'ROUND',
        'ЦЕЛОЕ': 'INT',
        'СУММ': 'SUM',
        'СРЗНАЧ': 'AVERAGE',
        'СЧЁТ': 'COUNT',
        'СЧЁТЗ': 'COUNTA',
        'МАКС': 'MAX',
        'МИН': 'MIN',
        'ЕСЛИ': 'IF',
        'И': 'AND',
        'ИЛИ': 'OR',
        'НЕ': 'NOT',
        'ВПЕР': 'VLOOKUP',
        'ГПР': 'HLOOKUP',
        'ИНДЕКС': 'INDEX',
        'ПОИСКПOZ': 'MATCH'
    }

    original_formula = formula
    converted_formula = formula
    for excel_func, google_func in replacements.items():
        converted_formula = converted_formula.replace(excel_func, google_func)

    print(f"КОНВЕРТАЦИЯ ФОРМУЛЫ: '{original_formula}' → '{converted_formula}'")
    return converted_formula


def copy_sheet_data(
        excel_sheet,
        google_worksheet,
        column_mapping: Dict[str, List[str]],
        start_row: int,
        log_callback: Optional[Callable[[str], None]] = None
) -> int:
    if log_callback:
        log_callback("Подготовка данных...")

    source_cols = resolve_excel_columns(excel_sheet, column_mapping['source'])
    target_cols = resolve_google_columns(google_worksheet, column_mapping['target'])

    if len(source_cols) != len(target_cols):
        raise ValueError("Количество исходных и целевых колонок должно совпадать")

    if log_callback:
        log_callback("Анализ видимых строк и данных Excel...")

    max_row = excel_sheet.max_row
    if max_row < start_row:
        if log_callback:
            log_callback("Нет данных для копирования")
        return 0

    visible_rows_with_data = []

    for row_num in range(start_row, max_row + 1):
        row_dimension = excel_sheet.row_dimensions.get(row_num)
        if row_dimension and row_dimension.hidden:
            continue

        has_data = False
        row_cells = []

        for col_letter in source_cols:
            cell = excel_sheet[f"{col_letter}{row_num}"]
            cell_value = cell.value
            cell_formula = None

            # ОТЛАДКА - выводим все что есть в ячейке
            if log_callback and cell_value is not None:
                debug_info = f"Ячейка {col_letter}{row_num}: value='{cell_value}', data_type='{getattr(cell, 'data_type', 'None')}'"
                if hasattr(cell, 'internal_value'):
                    debug_info += f", internal_value='{cell.internal_value}'"
                if hasattr(cell, 'f'):
                    debug_info += f", f='{cell.f}'"
                log_callback(debug_info)

            # Проверяем формулу через internal_value
            if hasattr(cell, 'internal_value') and str(cell.internal_value).startswith('='):
                cell_formula = str(cell.internal_value)
                has_data = True
                if log_callback:
                    log_callback(f"ФОРМУЛА найдена через internal_value: {cell_formula}")
            # Если нет internal_value, проверяем обычное значение
            elif cell_value and str(cell_value).startswith('='):
                cell_formula = str(cell_value)
                has_data = True
                if log_callback:
                    log_callback(f"ФОРМУЛА найдена через value: {cell_formula}")
            # Проверяем data_type
            elif hasattr(cell, 'data_type') and cell.data_type == 'f':
                if hasattr(cell, 'f'):
                    cell_formula = '=' + str(cell.f)
                else:
                    cell_formula = str(cell_value) if cell_value else None
                if cell_formula:
                    has_data = True
                    if log_callback:
                        log_callback(f"ФОРМУЛА найдена через data_type: {cell_formula}")
            elif cell_value is not None:
                has_data = True

            row_cells.append({
                'cell': cell,
                'value': cell_value,
                'formula': cell_formula,
                'formatting': get_cell_formatting(cell)
            })

        if has_data:
            visible_rows_with_data.append({
                'row_num': row_num,
                'cells': row_cells
            })

    if not visible_rows_with_data:
        if log_callback:
            log_callback("Нет видимых строк с данными для копирования")
        return 0

    if log_callback:
        log_callback(f"Найдено {len(visible_rows_with_data)} видимых строк с данными")

    values_to_update = []
    formats_to_apply = []

    google_start_row = start_row

    for i, row_data in enumerate(visible_rows_with_data):
        current_google_row = google_start_row + i
        row_values = []
        row_formats = []

        for j, cell_data in enumerate(row_data['cells']):
            cell_value = cell_data['value']
            cell_formula = cell_data['formula']
            cell_formatting = cell_data['formatting']

            if cell_formula:
                converted_formula = convert_excel_formula_to_google(cell_formula)
                if log_callback:
                    log_callback(
                        f"→ ФОРМУЛА ОБРАБОТАНА: '{cell_formula}' → '{converted_formula}' (строка {current_google_row}, колонка {target_cols[j]})")
                row_values.append(converted_formula)
            else:
                if log_callback and cell_value is not None:
                    log_callback(f"→ ЗНАЧЕНИЕ: '{cell_value}' (строка {current_google_row}, колонка {target_cols[j]})")
                row_values.append(cell_value if cell_value is not None else '')

            if cell_formatting:
                target_col_index = column_index_from_string(target_cols[j])
                row_formats.append({
                    'row': current_google_row,
                    'col': target_col_index,
                    'format': cell_formatting
                })

        values_to_update.append(row_values)
        formats_to_apply.extend(row_formats)

    if log_callback:
        log_callback(f"Подготовлено {len(values_to_update)} строк для записи")

    if log_callback:
        log_callback("Запись данных в Google Sheets...")

    try:
        col_numbers = [column_index_from_string(col) for col in target_cols]
        min_col = min(col_numbers)
        max_col = max(col_numbers)
        start_col_letter = get_column_letter(min_col)
        end_col_letter = get_column_letter(max_col)
        end_row = google_start_row + len(values_to_update) - 1
        target_range = f"{start_col_letter}{google_start_row}:{end_col_letter}{end_row}"

        num_cols = max_col - min_col + 1
        index_map = [column_index_from_string(col) - min_col for col in target_cols]

        formatted_values = []
        for row in values_to_update:
            row_values = [''] * num_cols
            for value, idx in zip(row, index_map):
                row_values[idx] = value if value is not None else ''
            formatted_values.append(row_values)

        if log_callback:
            log_callback(f"Обновление диапазона {target_range}...")

        google_worksheet.update(
            target_range,
            formatted_values,
            value_input_option='USER_ENTERED'
        )

        if log_callback:
            log_callback(f"✓ ДАННЫЕ ЗАПИСАНЫ в диапазон {target_range}")
            for i, row in enumerate(formatted_values):
                row_num = google_start_row + i
                for j, val in enumerate(row):
                    if val and str(val).startswith('='):
                        col_letter = get_column_letter(min_col + j)
                        log_callback(f"  ✓ ФОРМУЛА в {col_letter}{row_num}: {val}")

        if formats_to_apply:
            if log_callback:
                log_callback(f"Применение форматирования к {len(formats_to_apply)} ячейкам...")

            try:
                format_requests = []

                for format_data in formats_to_apply:
                    row = format_data['row']
                    col = format_data['col']
                    cell_format = format_data['format']

                    format_request = {
                        'updateCells': {
                            'range': {
                                'sheetId': google_worksheet.id,
                                'startRowIndex': row - 1,
                                'endRowIndex': row,
                                'startColumnIndex': col - 1,
                                'endColumnIndex': col
                            },
                            'rows': [{
                                'values': [{
                                    'userEnteredFormat': cell_format
                                }]
                            }],
                            'fields': 'userEnteredFormat'
                        }
                    }
                    format_requests.append(format_request)

                batch_size = 100
                for i in range(0, len(format_requests), batch_size):
                    batch = format_requests[i:i + batch_size]
                    if batch:
                        google_worksheet.spreadsheet.batch_update({
                            'requests': batch
                        })

                        if log_callback:
                            log_callback(
                                f"Применено форматирование к {min(len(batch), len(format_requests) - i)} ячейкам")

            except Exception as e:
                if log_callback:
                    log_callback(f"Предупреждение: не удалось применить все форматирование - {str(e)}")

        if log_callback:
            log_callback(f"✓ Успешно обновлено {len(values_to_update)} строк, {len(target_cols)} колонок")
            if formats_to_apply:
                log_callback(f"✓ Применено форматирование к {len(formats_to_apply)} ячейкам")

        return len(values_to_update)

    except Exception as e:
        if log_callback:
            log_callback(f"Ошибка при записи в Google Sheets: {e}")
        raise


def clear_column_cache():
    if hasattr(resolve_excel_columns, '_header_cache'):
        resolve_excel_columns._header_cache.clear()
    if hasattr(resolve_google_columns, '_header_cache'):
        resolve_google_columns._header_cache.clear()