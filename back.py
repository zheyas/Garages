import calendar
import re
from datetime import date, timedelta

import pandas as pd

# Константы для имён колонок и статусов
RENT_GARAGE_COL = "Номер гаража"
RENT_DATE_COL = "Дата оплаты"
RENT_SUM_COL = "Сумма оплаты"
STATUS_COL = "Статус"

STATUS_PAID = "Оплачено"
STATUS_OVERDUE = "Просрочено"
STATUS_PENDING = "Не оплачено"


def normalize_rent_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Переименовывает колонки с арендой в стандартизированные:
      * гараж → Номер гаража
      * дата  → Дата оплаты
      * сумм  → Сумма оплаты
    """
    rename_map = {}
    for col in df.columns:
        low = str(col).lower()
        if "гараж" in low:
            rename_map[col] = RENT_GARAGE_COL
        elif "дата" in low:
            rename_map[col] = RENT_DATE_COL
        elif "сумм" in low:
            rename_map[col] = RENT_SUM_COL
    return df.rename(columns=rename_map)


def load_rent_excel(path: str) -> pd.DataFrame:
    """
    Подгружает xlsx с арендой и сразу нормализует колонки.
    """
    df = pd.read_excel(path)
    return normalize_rent_columns(df)


def short_header(col) -> str:
    """
    Превращает заголовок столбца выписки
    в низкоуровневое имя: 'дата', 'сумма', 'категория' или чистый текст.
    """
    s = re.sub(r"[^а-яa-z0-9 ]", "", str(col).lower())
    if "дата" in s:
        return "дата"
    if "сумм" in s:
        return "сумма"
    if "категор" in s or "описан" in s:
        return "категория"
    return s.strip()


def load_statement(path: str) -> pd.DataFrame:
    """
    Извлекает из банковской выписки все блоки таблиц операций,
    склеивает их в один DataFrame и нормализует заголовки.
    """
    df = pd.read_excel(path, header=None)
    n_cols = df.shape[1]
    blocks = []
    titles = ["дата операции", "сумма в валюте"]

    i = 0
    while i < len(df):
        row_text = " ".join(df.iloc[i, :n_cols].astype(str).str.lower())
        if all(t in row_text for t in titles):
            header_idx = i
            data_idxs = []
            j = i + 1
            while j < len(df):
                txt = " ".join(df.iloc[j, :n_cols].astype(str).str.lower())
                if (
                    any(t in txt for t in titles)
                    or "выписка по платёжному счёту" in txt
                    or "продолжение на следующей странице" in txt
                    or df.iloc[j].isna().all()
                    or df.iloc[j]
                    .astype(str)
                    .str.strip()
                    .eq("")
                    .all()
                ):
                    break
                data_idxs.append(j)
                j += 1

            raw_cols = df.iloc[header_idx, :n_cols].tolist()
            valid = [
                idx
                for idx, c in enumerate(raw_cols)
                if str(c).strip() and str(c).lower() != "nan"
            ]
            cols = [raw_cols[idx] for idx in valid]

            block = df.iloc[data_idxs, valid].copy()
            block.columns = cols
            blocks.append(block)

            i = j
        else:
            i += 1

    if not blocks:
        raise ValueError("Не найдено ни одной таблицы операций в выписке!")

    all_ops = pd.concat(blocks, ignore_index=True)
    all_ops.columns = [short_header(c) for c in all_ops.columns]
    return all_ops


def normalize_sum(val) -> float:
    """
    Приводит сумму из выписки к float, убирая пробелы, '+' и меняя ','→'.'
    """
    try:
        s = str(val).replace(" ", "").replace("+", "").replace(",", ".")
        return float(s)
    except (ValueError, TypeError):
        return None


def extract_first_date(val) -> str:
    """
    Извлекает первую дату формата dd.mm.yyyy из строки.
    """
    m = re.search(r"\d{2}\.\d{2}\.\d{4}", str(val))
    return m.group(0) if m else None


def analyze_payments(rent_df: pd.DataFrame,
                     stmt_df: pd.DataFrame) -> pd.DataFrame:
    """
    Сверяет суммы и даты из арендных данных и выписки,
    возвращает DataFrame с колонками:
      * Номер гаража
      * Дата оплаты (исправленная)
      * Сумма оплаты
      * Статус
    """
    today = date.today()

    stmt = stmt_df.copy()
    stmt.columns = [str(c).strip().lower() for c in stmt.columns]
    sum_cols = [c for c in stmt.columns if "сумм" in c or "сумма" in c]
    date_cols = [c for c in stmt.columns if "дата" in c]
    if not sum_cols or not date_cols:
        raise ValueError(
            "В выписке не найдены колонки с суммой или датой!"
        )

    sum_col = sum_cols[0]
    date_col = date_cols[0]
    stmt[sum_col + "_num"] = stmt[sum_col].apply(normalize_sum)

    rows = []
    for _, row in rent_df.iterrows():
        garage = row.get(RENT_GARAGE_COL)
        try:
            raw_date = pd.to_datetime(row.get(RENT_DATE_COL)).date()
        except Exception:
            raw_date = today

        try:
            rent_sum = float(row.get(RENT_SUM_COL))
        except Exception:
            rent_sum = None

        year, mon = today.year, today.month
        last_day = calendar.monthrange(year, mon)[1]
        day = min(raw_date.day, last_day)
        pay_date = raw_date.replace(year=year, month=mon, day=day)
        deadline = pay_date + timedelta(days=3)

        matched = stmt.loc[
            stmt[sum_col + "_num"].round(2) == round(rent_sum, 2)
            ].copy()  # <- здесь .copy()
        matched.loc[:, "dt_str"] = (
            matched[date_col]
            .astype(str)
            .map(extract_first_date)
        )

        dates = (
            pd.to_datetime(
                matched["dt_str"], format="%d.%m.%Y", errors="coerce"
            )
            .dt.date.dropna()
        )

        if not matched.empty and not dates.empty:
            last_pay = dates.max()
            status = (
                STATUS_PAID
                if last_pay <= deadline
                else STATUS_OVERDUE
            )
        else:
            status = (
                STATUS_PENDING
                if today <= deadline
                else STATUS_OVERDUE
            )

        rows.append(
            {
                RENT_GARAGE_COL: garage,
                RENT_DATE_COL: pay_date,
                RENT_SUM_COL: rent_sum,
                STATUS_COL: status,
            }
        )

    return pd.DataFrame(rows)


def apply_status_styling(status: str) -> str:
    """
    Для pandas.Style: возвращает inline-css для ячейки 'Статус'.
    """
    mapping = {
        STATUS_PAID: "background-color: #2ecc71; color: white;",
        STATUS_OVERDUE: "background-color: #e74c3c; color: white;",
        STATUS_PENDING: "background-color: #95a5a6; color: white;",
    }
    return mapping.get(status, "")


def save_to_excel(df: pd.DataFrame, filename: str) -> None:
    """
    Сохраняет DataFrame в Excel без индекса.
    """
    df.to_excel(filename, index=False)
