import pandas as pd
from datetime import datetime, timedelta
import calendar
import re

def normalize_arenda_columns(df):
    """Приводит заголовки в стандартные имена: 'Номер гаража', 'Дата оплаты', 'Сумма оплаты'."""
    rename_map = {}
    for col in df.columns:
        col_lower = str(col).lower()
        if 'гараж' in col_lower:
            rename_map[col] = 'Номер гаража'
        elif 'дата' in col_lower:
            rename_map[col] = 'Дата оплаты'
        elif 'сумм' in col_lower:
            rename_map[col] = 'Сумма оплаты'
    df = df.rename(columns=rename_map)
    return df

def load_excel(path):
    df = pd.read_excel(path)
    df = normalize_arenda_columns(df)
    return df

def short_header(col):
    s = str(col).lower()
    s = re.sub(r'[^а-яa-z0-9 ]', '', s)
    if "дата" in s: return "дата"
    if "сумм" in s: return "сумма"
    if "категор" in s or "описан" in s: return "категория"
    return s

def load_vypiska_operations(path):
    df = pd.read_excel(path, header=None)
    op_blocks = []
    search_titles = ["дата операции", "сумма в валюте"]
    n_cols = df.shape[1]

    i = 0
    while i < len(df):
        row = df.iloc[i]
        row_str = [str(cell).lower() for cell in row[:n_cols]]
        if search_titles[0] in "".join(row_str) and search_titles[1] in "".join(row_str):
            header_row = i
            block_rows = []
            j = i + 1
            while j < len(df):
                next_row_str = "".join([str(x).lower() for x in df.iloc[j][:n_cols]])
                if search_titles[0] in next_row_str and search_titles[1] in next_row_str:
                    break
                if 'выписка по платёжному счёту' in next_row_str or 'продолжение на следующей странице' in next_row_str:
                    break
                if all((str(v).strip() == '' or pd.isna(v)) for v in df.iloc[j]):
                    break
                block_rows.append(j)
                j += 1
            # Обрезаем nan-колонки в заголовке
            cols = df.iloc[header_row, :n_cols].tolist()
            valid_cols_idx = [ii for ii, c in enumerate(cols) if str(c).strip() and str(c).lower() != "nan"]
            cols = [cols[ii] for ii in valid_cols_idx]
            block_df = df.iloc[block_rows, valid_cols_idx].copy()
            block_df.columns = cols
            op_blocks.append(block_df)
            i = j
        else:
            i += 1
    if not op_blocks:
        raise ValueError("Не найдено ни одной таблицы операций!")
    all_ops_df = pd.concat(op_blocks, ignore_index=True)
    all_ops_df.columns = [short_header(c) for c in all_ops_df.columns]
    return all_ops_df

def normalize_sum(val):
    """Преобразует сумму из выписки Сбера к float"""
    try:
        sval = str(val).strip().replace(' ', '').replace('+','').replace(',','.')
        return float(sval)
    except:
        return None

def extract_first_date(val):
    """Достаёт только первую дату формата ДД.ММ.ГГГГ из строки (работает для столбца дата выписки)"""
    val = str(val) if not isinstance(val, str) else val
    m = re.search(r'(\d{2}\.\d{2}\.\d{4})', val)
    return m.group(1) if m else None

def analyze_payments(arenda_df, vypiska_df):
    today = datetime.today().date()
    result = []

    vypiska_df.columns = [str(c).strip().lower() for c in vypiska_df.columns]
    # print("Vypiska DF columns:", list(vypiska_df.columns))

    sum_cols = [c for c in vypiska_df.columns if 'сумм' in c or 'сумма' in c]
    date_cols = [c for c in vypiska_df.columns if 'дата' in c]
    if not sum_cols or not date_cols:
        raise ValueError("Не найдены колонки с суммой или датой в выписке. Проверь формат файла.")

    for _, row in arenda_df.iterrows():
        garage = row["Номер гаража"]
        try:
            raw_date = pd.to_datetime(row["Дата оплаты"]).date()
        except Exception:
            raw_date = today

        try:
            sum_arenda = float(row["Сумма оплаты"])
        except Exception:
            sum_arenda = None

        # Корректировка даты
        year = today.year
        target_month = today.month
        last_day = calendar.monthrange(year, target_month)[1]
        corrected_day = min(raw_date.day, last_day)
        corrected_date = raw_date.replace(year=year, month=target_month, day=corrected_day)
        deadline = corrected_date + timedelta(days=3)

        sums_vypiska = vypiska_df[sum_cols[0]].apply(normalize_sum)
        matched = vypiska_df[sums_vypiska.round(2) == round(sum_arenda, 2)]

        # Главный фикс для Сбера: достать дату из вида "02.06.2025 02.06.2025"
        date_strings = matched[date_cols[0]].astype(str).map(extract_first_date)
        matched_dates = pd.to_datetime(date_strings, format='%d.%m.%Y', errors='coerce').dropna()

        if not matched.empty and not matched_dates.empty:
            last_payment = matched_dates.max().date()
            status = "Получен" if last_payment <= deadline else "Просрочен"
        else:
            status = "Срок не наступил" if today <= deadline else "Просрочен"

        result.append({
            "Номер гаража": garage,
            "Дата оплаты": corrected_date,
            "Сумма оплаты": sum_arenda,
            "Статус": status
        })

    return pd.DataFrame(result)



def apply_status_styling(status):
    """
    Возвращает строку с CSS-стилями для ячейки 'Статус'
    на основании её значения.
    """
    mapping = {
        'Оплачено':        'background-color: #2ecc71; color: white;',  # зелёный
        'Просрочено':      'background-color: #e74c3c; color: white;',  # красный
        'Частично оплачено':'background-color: #f1c40f; color: black;',  # жёлтый
        'Не оплачено':     'background-color: #95a5a6; color: white;',  # серый
    }
    # если встречается непрописанный статус — ничего не красим
    return mapping.get(status, '')

def save_to_excel(df, filename="result.xlsx"):
    df.to_excel(filename, index=False)