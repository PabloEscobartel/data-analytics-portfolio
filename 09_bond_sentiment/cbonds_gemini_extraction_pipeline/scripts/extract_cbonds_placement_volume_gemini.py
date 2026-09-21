#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Извлечение объема размещения из новостей Cbonds с помощью Gemini.

Логика для каждого выпуска из bonds_offer_before2018_all_orientirs_gemini_cbonds.xlsx:
1. Берём время начала букбилдинга из колонки `cbonds_book_start_datetime`.
2. Берём все новости, уже привязанные к выпуску в `issue_articles`, в окне:
   [cbonds_book_start_datetime минус 1 месяц; Окончание размещения плюс 1 неделя].
3. Полностью пропускаем эмитента ВЭБ.РФ.
4. Gemini ищет:
   - последний доступный ориентир/планируемое значение объема размещения ДО
     начала букбилдинга;
   - финальное значение объема размещения, установленное/объявленное эмитентом
     по итогам букбилдинга.

Числовые значения возвращаются в миллиардах валюты выпуска. Для рублевых
выпусков это млрд рублей. Валютная конвертация не выполняется.

Результат пишется в копию Excel и в SQLite-таблицу issue_placement_volume_gemini.
Промежуточный кэш сохраняется в JSON, поэтому скрипт можно перезапускать.
"""

from __future__ import annotations

import argparse
import calendar
import json
import math
import os
import re
import sqlite3
import time
from collections import deque
from dataclasses import dataclass
from datetime import date, datetime, time as dt_time, timedelta
from pathlib import Path
from typing import Any, Optional

import openpyxl

try:
    from google import genai
    from google.genai import types
except ImportError:  # pragma: no cover
    genai = None
    types = None


DEFAULT_MODEL = os.getenv("GEMINI_MODEL", "gemini-3.1-flash-lite")
DEFAULT_DELAY = 15.0
DEFAULT_MAX_RETRIES = 5
DEFAULT_MAX_OUTPUT_TOKENS = 8192
DEFAULT_SAVE_EVERY = 10
DEFAULT_MAX_BATCH_CHARS = 25000
DEFAULT_MAX_ARTICLE_CHARS = 0
DEFAULT_MAX_ISSUE_NEWS = 0
DEFAULT_TPM_LIMIT = int(os.getenv("GEMINI_TPM_LIMIT", "0") or 0)
DEFAULT_TPM_SAFETY = float(os.getenv("GEMINI_TPM_SAFETY", "0.9") or 0.9)
DEFAULT_CHARS_PER_TOKEN = float(os.getenv("GEMINI_CHARS_PER_TOKEN", "2.0") or 2.0)

INPUT_FILE = "bonds_offer_before2018_all_orientirs_gemini_cbonds.xlsx"
OUTPUT_FILE = "bonds_offer_before2018_placement_volume_gemini.xlsx"
DB_FILE = "cbonds_news.sqlite3"
CACHE_FILE = "gemini_placement_volume_before2018_cache.json"
RESULTS_TABLE = "issue_placement_volume_before2018_gemini"

DEFAULT_SHEET = "place_before_2018"
DEFAULT_PAPER_COLUMN = "Бумага"
DEFAULT_ISIN_COLUMN = "ISIN"
DEFAULT_ISSUER_COLUMN = "Эмитент"
DEFAULT_BOOK_DATE_COLUMN = "cbonds_book_start_datetime"
DEFAULT_PLACEMENT_END_COLUMN = "Окончание размещения"
OPTIONAL_CURRENCY_HEADERS = ["Валюта", "currency"]

OUTPUT_COLUMNS = [
    "volume_gemini_pre_value_bn",
    "volume_gemini_pre_value_text",
    "volume_gemini_pre_source_title",
    "volume_gemini_pre_source_url",
    "volume_gemini_pre_source_article_id",
    "volume_gemini_pre_source_datetime",
    "volume_gemini_pre_evidence_text",
    "volume_gemini_final_value_bn",
    "volume_gemini_final_value_text",
    "volume_gemini_final_source_title",
    "volume_gemini_final_source_url",
    "volume_gemini_final_source_article_id",
    "volume_gemini_final_source_datetime",
    "volume_gemini_final_evidence_text",
    "volume_gemini_currency",
    "volume_gemini_needs_attention",
    "volume_gemini_attention_reason",
    "volume_gemini_reason_short",
    "volume_gemini_confidence",
    "volume_gemini_model_name",
    "volume_gemini_model_version",
    "volume_gemini_status",
]

SYSTEM_PROMPT = """Ты аналитик российского первичного долгового рынка.

Тебе передают ОДИН выпуск облигаций и СПИСОК новостей Cbonds, собранных за окно от месяца до начала букбилдинга до недели после окончания размещения.

Задача:
1. Определи, какие новости относятся именно к указанному выпуску/эмиссии.
2. Найди ПОСЛЕДНЕЕ доступное значение объема размещения/объема выпуска/планируемого объема ДО начала букбилдинга.
3. Найди ФИНАЛЬНОЕ значение объема размещения, которое эмитент установил или объявил по итогам букбилдинга/сбора книги заявок/размещения.

КРИТИЧЕСКИ ВАЖНО:
- Нужен именно ОБЪЕМ РАЗМЕЩЕНИЯ / объем выпуска / планируемый объем займа, а не купон, ставка, доходность, spread, срок, рейтинг или объем спроса.
- Для pre_value выбирай только новости с published_at < bookbuilding_start. Если таких значений несколько, выбери самое позднее по времени публикации.
- Для final_value выбирай значение, явно являющееся итогом букбилдинга/сбора заявок/размещения: "установил объем", "объем размещения составил", "по итогам букбилдинга объем выпуска установлен", "разместил облигации на сумму".
- Не считай финальным значением объем спроса, книгу заявок, переподписку, лимит программы, общий объем программы облигаций или объем другого выпуска.
- Если финальная новость говорит, что объем увеличен/снижен до X, верни X.
- Если выпуск размещен частично и в новости указан фактически размещенный объем, это финальное значение.
- Если в новости несколько выпусков, выбери только значение, относящееся к ISSUE.paper / ISSUE.isin.
- parsed_emissions у новости — это отдельное поле эмиссии, извлеченное со страницы новости Cbonds.
- Если parsed_emissions совпадает с ISSUE.paper / ISSUE.issue_number / ISSUE.isin, считай такую новость относящейся к целевому выпуску, даже если номер выпуска есть только в parsed_emissions и отсутствует в заголовке или основном тексте.
- Поле parsed_emission_matches_target: yes означает, что скрипт уже нашел совпадение parsed_emissions с целевой эмиссией; используй это как сильный сигнал релевантности новости.
- parsed_emissions может отсутствовать или содержать неполный список. Если совпадения нет, обязательно читай заголовок и текст.
- Если значение вероятно относится к целевой эмиссии, но связь неоднозначна, верни значение и поставь needs_attention=true.

Нормализация числа:
- Верни pre_value_bn и final_value_bn в МИЛЛИАРДАХ валюты выпуска.
- 10 млрд = 10.
- 750 млн = 0.75.
- 100 млн = 0.1.
- 1 500 000 000 = 1.5.
- Не выполняй валютную конвертацию. Если валюта новости отличается от валюты выпуска или непонятна, поставь needs_attention=true.
- Если значения нет, верни 0.

source_article_id:
- Верни точный integer article_id выбранной новости из списка ARTICLES.
- Если значение не найдено, верни 0.

status:
- found: найдены оба значения.
- partial: найдено только одно из двух значений.
- none: не найдено ни одного значения.

Верни только JSON-массив объектов без markdown. Один объект на каждый выпуск из ISSUES.
"""

RESULT_SCHEMA = {
    "type": "array",
    "items": {
        "type": "object",
        "properties": {
            "excel_row": {"type": "integer", "minimum": 2},
            "status": {"type": "string", "enum": ["found", "partial", "none"]},
            "pre_value_bn": {"type": "number", "minimum": 0},
            "pre_value_text": {"type": "string"},
            "pre_evidence_text": {"type": "string"},
            "pre_source_article_id": {"type": "integer"},
            "final_value_bn": {"type": "number", "minimum": 0},
            "final_value_text": {"type": "string"},
            "final_evidence_text": {"type": "string"},
            "final_source_article_id": {"type": "integer"},
            "needs_attention": {"type": "boolean"},
            "attention_reason": {"type": "string"},
            "confidence": {"type": "number", "minimum": 0, "maximum": 1},
            "reason_short": {"type": "string"},
        },
        "required": [
            "excel_row",
            "status",
            "pre_value_bn",
            "pre_value_text",
            "pre_evidence_text",
            "pre_source_article_id",
            "final_value_bn",
            "final_value_text",
            "final_evidence_text",
            "final_source_article_id",
            "needs_attention",
            "attention_reason",
            "confidence",
            "reason_short",
        ],
        "additionalProperties": False,
    },
}

LAST_REQUEST_TS = 0.0
TOKEN_WINDOW: deque[tuple[float, int]] = deque()


@dataclass
class IssueInfo:
    excel_row: int
    paper: str
    isin: str
    emitent: str
    bookbuilding_start: datetime
    placement_date_end: date
    currency: str


@dataclass
class CandidateArticle:
    article_row_id: int
    article_id: int | None
    url: str
    title: str
    date_time: str
    published_at: datetime
    text: str
    emission_cbonds_id: int | None
    emission_name: str
    emission_url: str
    timing: str


@dataclass
class BatchItem:
    excel_row: int
    prompt_text: str
    estimated_chars: int


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def normalize_key(value: Any) -> str:
    text = normalize_text(value).casefold().replace("ё", "е")
    text = re.sub(r"\b(пао|ао|ооо|зао|оао)\b", " ", text)
    text = re.sub(r"[^0-9a-zа-я]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


CYR_TO_LAT = str.maketrans(
    {
        "А": "A",
        "В": "B",
        "Е": "E",
        "К": "K",
        "М": "M",
        "Н": "H",
        "О": "O",
        "Р": "P",
        "С": "C",
        "Т": "T",
        "У": "Y",
        "Х": "X",
        "а": "A",
        "в": "B",
        "е": "E",
        "к": "K",
        "м": "M",
        "н": "H",
        "о": "O",
        "р": "P",
        "с": "C",
        "т": "T",
        "у": "Y",
        "х": "X",
    }
)


def normalize_issue_code(value: Any) -> str:
    text = normalize_text(value).translate(CYR_TO_LAT).upper()
    return re.sub(r"[^0-9A-ZА-Я]+", "", text)


def extract_issue_number(paper: Any) -> str:
    text = normalize_text(paper)
    if "," not in text:
        return ""
    return text.split(",", 1)[1].strip()


def parsed_emission_matches_target(issue: "IssueInfo", candidate: "CandidateArticle") -> bool:
    emission_text = " ".join([candidate.emission_name, candidate.emission_url])
    emission_key = normalize_issue_code(emission_text)
    if not emission_key:
        return False

    issue_number = normalize_issue_code(extract_issue_number(issue.paper))
    isin = normalize_issue_code(issue.isin)
    paper = normalize_issue_code(issue.paper)

    if issue_number and issue_number in emission_key:
        return True
    if isin and isin in emission_key:
        return True
    return bool(paper and (paper in emission_key or emission_key in paper))


def is_veb_rf(issuer: str) -> bool:
    return normalize_key(issuer) in {"вэб рф", "веб рф"}


def safe_float_01(value: Any, default: float = 0.5) -> float:
    try:
        x = float(value)
    except (TypeError, ValueError):
        return default
    if math.isnan(x):
        return default
    return max(0.0, min(1.0, x))


def safe_nonneg_number(value: Any, default: float = 0.0) -> float:
    try:
        x = float(value)
    except (TypeError, ValueError):
        return default
    if math.isnan(x):
        return default
    return max(0.0, x)


def parse_date(value: Any) -> Optional[date]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = normalize_text(value)
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text).date()
    except Exception:
        return None


def parse_datetime(value: Any) -> Optional[datetime]:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.replace(microsecond=0)
    if isinstance(value, date):
        return datetime.combine(value, dt_time(0, 0, 0))
    text = normalize_text(value)
    if not text:
        return None
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%Y-%m-%d",
        "%d.%m.%Y",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text)
    except Exception:
        return None


def parse_article_datetime(value: Any) -> Optional[datetime]:
    text = normalize_text(value)
    if not text:
        return None
    for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text)
    except Exception:
        return None


def subtract_calendar_month(value: datetime) -> datetime:
    year = value.year
    month = value.month - 1
    if month == 0:
        month = 12
        year -= 1
    day = min(value.day, calendar.monthrange(year, month)[1])
    return value.replace(year=year, month=month, day=day)


def window_end_for_placement(placement_end: date) -> datetime:
    return datetime.combine(placement_end + timedelta(days=7), dt_time(23, 59, 59))


def load_json(path: Path) -> dict[str, Any]:
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_json(path: Path, payload: dict[str, Any]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def build_header_index(ws) -> dict[str, int]:
    return {normalize_key(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1)}


def find_optional_column(headers: dict[str, int], variants: list[str]) -> int | None:
    for variant in variants:
        idx = headers.get(normalize_key(variant))
        if idx:
            return idx
    return None


def required_headers(
    paper_column: str,
    isin_column: str,
    issuer_column: str,
    book_date_column: str,
    placement_end_column: str,
) -> list[str]:
    return [paper_column, isin_column, issuer_column, book_date_column, placement_end_column]


def find_target_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str | None,
    paper_column: str,
    isin_column: str,
    issuer_column: str,
    book_date_column: str,
    placement_end_column: str,
):
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise RuntimeError(f"Лист {sheet_name!r} не найден. Доступные листы: {', '.join(wb.sheetnames)}")
        ws = wb[sheet_name]
        headers = build_header_index(ws)
        missing = [
            name
            for name in required_headers(paper_column, isin_column, issuer_column, book_date_column, placement_end_column)
            if normalize_key(name) not in headers
        ]
        if missing:
            raise RuntimeError(f"Не найдены колонки на листе {ws.title!r}: {missing}")
        return ws

    required = {
        normalize_key(x)
        for x in required_headers(paper_column, isin_column, issuer_column, book_date_column, placement_end_column)
    }
    for ws in wb.worksheets:
        headers = build_header_index(ws)
        if required.issubset(headers):
            return ws
    raise RuntimeError(f"Не найден лист с колонками: {sorted(required)}")


def strip_workbook_images(wb: openpyxl.Workbook) -> None:
    for ws in wb.worksheets:
        if hasattr(ws, "_images"):
            ws._images = []


def load_excel_issues(
    input_path: Path,
    sheet_name: str | None,
    paper_column: str = DEFAULT_PAPER_COLUMN,
    isin_column: str = DEFAULT_ISIN_COLUMN,
    issuer_column: str = DEFAULT_ISSUER_COLUMN,
    book_date_column: str = DEFAULT_BOOK_DATE_COLUMN,
    placement_end_column: str = DEFAULT_PLACEMENT_END_COLUMN,
) -> tuple[openpyxl.Workbook, Any, dict[int, IssueInfo], int]:
    wb = openpyxl.load_workbook(input_path)
    strip_workbook_images(wb)
    ws = find_target_sheet(
        wb,
        sheet_name,
        paper_column=paper_column,
        isin_column=isin_column,
        issuer_column=issuer_column,
        book_date_column=book_date_column,
        placement_end_column=placement_end_column,
    )
    headers = build_header_index(ws)
    missing = [
        name
        for name in required_headers(paper_column, isin_column, issuer_column, book_date_column, placement_end_column)
        if normalize_key(name) not in headers
    ]
    if missing:
        raise RuntimeError(f"Не найдены колонки на листе {ws.title!r}: {missing}")

    currency_col = find_optional_column(headers, OPTIONAL_CURRENCY_HEADERS)
    skipped_veb = 0
    issues: dict[int, IssueInfo] = {}
    for row_idx in range(2, ws.max_row + 1):
        paper = normalize_text(ws.cell(row_idx, headers[normalize_key(paper_column)]).value)
        isin = normalize_text(ws.cell(row_idx, headers[normalize_key(isin_column)]).value)
        emitent = normalize_text(ws.cell(row_idx, headers[normalize_key(issuer_column)]).value)
        bookbuilding_start = parse_datetime(ws.cell(row_idx, headers[normalize_key(book_date_column)]).value)
        placement_end = parse_date(ws.cell(row_idx, headers[normalize_key(placement_end_column)]).value)
        currency = normalize_text(ws.cell(row_idx, currency_col).value) if currency_col else ""

        if is_veb_rf(emitent):
            skipped_veb += 1
            continue
        if not paper or not emitent or bookbuilding_start is None or placement_end is None:
            continue

        issues[row_idx] = IssueInfo(
            excel_row=row_idx,
            paper=paper,
            isin=isin,
            emitent=emitent,
            bookbuilding_start=bookbuilding_start,
            placement_date_end=placement_end,
            currency=currency,
        )

    return wb, ws, issues, skipped_veb


def numeric_cell(value: Any, default: float = 0.0) -> float:
    if value is None or normalize_text(value) == "":
        return default
    text = normalize_text(value).replace(",", ".")
    try:
        return float(text)
    except (TypeError, ValueError):
        return default


def retry_problem_rows_from_workbook(ws, issues: dict[int, IssueInfo]) -> set[int]:
    headers = build_header_index(ws)
    required = [
        "volume_gemini_pre_value_bn",
        "volume_gemini_final_value_bn",
        "volume_issue_number_mismatch",
    ]
    missing = [column for column in required if normalize_key(column) not in headers]
    if missing:
        raise RuntimeError(
            "Для --retry-volume-problems не найдены колонки: "
            f"{missing}. Сначала запусти flag_volume_issue_number_mismatch.py."
        )

    pre_col = headers[normalize_key("volume_gemini_pre_value_bn")]
    final_col = headers[normalize_key("volume_gemini_final_value_bn")]
    mismatch_col = headers[normalize_key("volume_issue_number_mismatch")]

    result: set[int] = set()
    for row_idx in issues:
        pre_value = numeric_cell(ws.cell(row_idx, pre_col).value)
        final_value = numeric_cell(ws.cell(row_idx, final_col).value)
        mismatch = numeric_cell(ws.cell(row_idx, mismatch_col).value)
        if pre_value <= 0 or final_value <= 0 or mismatch == 1:
            result.add(row_idx)
    return result


def ensure_output_columns(ws) -> dict[str, int]:
    headers = build_header_index(ws)
    for col_name in OUTPUT_COLUMNS:
        key = normalize_key(col_name)
        if key not in headers:
            col_idx = ws.max_column + 1
            ws.cell(1, col_idx, col_name)
            headers[key] = col_idx
    return headers


def load_article_candidates(db_path: Path, issues: dict[int, IssueInfo]) -> dict[int, list[CandidateArticle]]:
    if not issues:
        return {}
    issues_by_isin = {
        normalize_issue_code(issue.isin): row_idx
        for row_idx, issue in issues.items()
        if normalize_issue_code(issue.isin)
    }
    issues_by_emitent_paper = {
        (normalize_key(issue.emitent), normalize_key(issue.paper)): row_idx
        for row_idx, issue in issues.items()
        if normalize_key(issue.emitent) and normalize_key(issue.paper)
    }
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    rows = conn.execute(
        """
        SELECT
            ia.excel_row,
            ia.paper AS linked_paper,
            ia.isin AS linked_isin,
            ia.emitent_name AS linked_emitent_name,
            a.id AS article_row_id,
            a.article_id,
            a.url,
            a.title,
            a.date_time,
            a.content_text,
            a.introtext,
            GROUP_CONCAT(DISTINCT ae.emission_cbonds_id) AS emission_cbonds_ids,
            GROUP_CONCAT(DISTINCT ae.emission_name) AS emission_names,
            GROUP_CONCAT(DISTINCT ae.emission_url) AS emission_urls
        FROM issue_articles ia
        JOIN articles a ON a.article_id = ia.article_id
        LEFT JOIN article_emissions ae ON ae.article_row_id = a.id
        GROUP BY
            ia.excel_row,
            ia.paper,
            ia.isin,
            ia.emitent_name,
            a.id,
            a.article_id,
            a.url,
            a.title,
            a.date_time,
            a.content_text,
            a.introtext
        ORDER BY ia.excel_row, a.date_time, a.id
        """
    ).fetchall()
    conn.close()

    by_issue: dict[int, dict[int, CandidateArticle]] = {}
    for row in rows:
        linked_isin = normalize_issue_code(row["linked_isin"])
        linked_emitent = normalize_key(row["linked_emitent_name"])
        linked_paper = normalize_key(row["linked_paper"])

        excel_row = issues_by_isin.get(linked_isin)
        if excel_row is None:
            excel_row = issues_by_emitent_paper.get((linked_emitent, linked_paper))
        if excel_row is None:
            raw_excel_row = int(row["excel_row"])
            if not linked_isin and not linked_paper and raw_excel_row in issues:
                excel_row = raw_excel_row
        if excel_row is None:
            continue

        issue = issues.get(excel_row)
        if issue is None:
            continue

        published_at = parse_article_datetime(row["date_time"])
        if published_at is None:
            continue

        window_start = subtract_calendar_month(issue.bookbuilding_start)
        window_end = window_end_for_placement(issue.placement_date_end)
        if not (window_start <= published_at <= window_end):
            continue
        timing = "pre_book" if published_at < issue.bookbuilding_start else "post_book"

        emission_names = normalize_text(row["emission_names"])
        emission_urls = normalize_text(row["emission_urls"])
        emission_ids_text = normalize_text(row["emission_cbonds_ids"])
        emission_cbonds_id = None
        if emission_ids_text and "," not in emission_ids_text:
            try:
                emission_cbonds_id = int(emission_ids_text)
            except ValueError:
                emission_cbonds_id = None

        article_row_id = int(row["article_row_id"])
        text = normalize_text(" ".join([normalize_text(row["introtext"]), normalize_text(row["content_text"])]))
        candidate = CandidateArticle(
            article_row_id=article_row_id,
            article_id=None if row["article_id"] is None else int(row["article_id"]),
            url=normalize_text(row["url"]),
            title=normalize_text(row["title"]),
            date_time=normalize_text(row["date_time"]),
            published_at=published_at,
            text=text,
            emission_cbonds_id=emission_cbonds_id,
            emission_name=emission_names,
            emission_url=emission_urls,
            timing=timing,
        )

        by_issue.setdefault(excel_row, {})[article_row_id] = candidate

    return {row_id: sorted(items.values(), key=lambda x: x.published_at) for row_id, items in by_issue.items()}


def validate_sql_identifier(value: str) -> str:
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", value or ""):
        raise ValueError(f"Некорректное имя SQL-таблицы: {value!r}")
    return value


def ensure_results_table(conn: sqlite3.Connection, table_name: str = RESULTS_TABLE) -> None:
    table_name = validate_sql_identifier(table_name)
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {table_name} (
            excel_row INTEGER PRIMARY KEY,
            paper TEXT,
            isin TEXT,
            emitent TEXT,
            bookbuilding_start TEXT,
            placement_date_end TEXT,
            currency TEXT,
            pre_value_bn REAL,
            pre_value_text TEXT,
            pre_source_title TEXT,
            pre_source_url TEXT,
            pre_source_article_id INTEGER,
            pre_source_datetime TEXT,
            pre_evidence_text TEXT,
            final_value_bn REAL,
            final_value_text TEXT,
            final_source_title TEXT,
            final_source_url TEXT,
            final_source_article_id INTEGER,
            final_source_datetime TEXT,
            final_evidence_text TEXT,
            needs_attention INTEGER NOT NULL DEFAULT 0,
            attention_reason TEXT,
            reason_short TEXT,
            confidence REAL,
            model_name TEXT,
            model_version TEXT,
            status TEXT,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.commit()


def source_by_article_id(candidates: list[CandidateArticle], source_article_id: int | None) -> CandidateArticle | None:
    if not source_article_id:
        return None
    for candidate in candidates:
        if candidate.article_id == source_article_id:
            return candidate
    return None


def write_cache_to_db(
    db_path: Path,
    issues: dict[int, IssueInfo],
    cache: dict[str, Any],
    table_name: str = RESULTS_TABLE,
) -> None:
    conn = sqlite3.connect(db_path)
    table_name = validate_sql_identifier(table_name)
    ensure_results_table(conn, table_name)
    now = datetime.now().isoformat(timespec="seconds")

    rows = []
    for key, result in cache.items():
        try:
            row_idx = int(key)
        except ValueError:
            continue
        issue = issues.get(row_idx)
        if issue is None or not isinstance(result, dict):
            continue
        rows.append(
            (
                row_idx,
                issue.paper,
                issue.isin,
                issue.emitent,
                issue.bookbuilding_start.isoformat(sep=" "),
                issue.placement_date_end.isoformat(),
                issue.currency,
                result.get("pre_value_bn"),
                result.get("pre_value_text"),
                result.get("pre_source_title"),
                result.get("pre_source_url"),
                result.get("pre_source_article_id"),
                result.get("pre_source_datetime"),
                result.get("pre_evidence_text"),
                result.get("final_value_bn"),
                result.get("final_value_text"),
                result.get("final_source_title"),
                result.get("final_source_url"),
                result.get("final_source_article_id"),
                result.get("final_source_datetime"),
                result.get("final_evidence_text"),
                int(result.get("needs_attention") or 0),
                result.get("attention_reason"),
                result.get("reason_short"),
                result.get("confidence"),
                result.get("model_name"),
                result.get("model_version"),
                result.get("status"),
                now,
            )
        )

    conn.executemany(
        f"""
        INSERT INTO {table_name} (
            excel_row, paper, isin, emitent, bookbuilding_start, placement_date_end, currency,
            pre_value_bn, pre_value_text, pre_source_title, pre_source_url,
            pre_source_article_id, pre_source_datetime, pre_evidence_text,
            final_value_bn, final_value_text, final_source_title, final_source_url,
            final_source_article_id, final_source_datetime, final_evidence_text,
            needs_attention, attention_reason, reason_short, confidence,
            model_name, model_version, status, updated_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(excel_row) DO UPDATE SET
            paper = excluded.paper,
            isin = excluded.isin,
            emitent = excluded.emitent,
            bookbuilding_start = excluded.bookbuilding_start,
            placement_date_end = excluded.placement_date_end,
            currency = excluded.currency,
            pre_value_bn = excluded.pre_value_bn,
            pre_value_text = excluded.pre_value_text,
            pre_source_title = excluded.pre_source_title,
            pre_source_url = excluded.pre_source_url,
            pre_source_article_id = excluded.pre_source_article_id,
            pre_source_datetime = excluded.pre_source_datetime,
            pre_evidence_text = excluded.pre_evidence_text,
            final_value_bn = excluded.final_value_bn,
            final_value_text = excluded.final_value_text,
            final_source_title = excluded.final_source_title,
            final_source_url = excluded.final_source_url,
            final_source_article_id = excluded.final_source_article_id,
            final_source_datetime = excluded.final_source_datetime,
            final_evidence_text = excluded.final_evidence_text,
            needs_attention = excluded.needs_attention,
            attention_reason = excluded.attention_reason,
            reason_short = excluded.reason_short,
            confidence = excluded.confidence,
            model_name = excluded.model_name,
            model_version = excluded.model_version,
            status = excluded.status,
            updated_at = excluded.updated_at
        """,
        rows,
    )
    conn.commit()
    conn.close()


def build_issue_prompt_block(
    issue: IssueInfo,
    candidates: list[CandidateArticle],
    max_article_chars: int = DEFAULT_MAX_ARTICLE_CHARS,
    max_issue_news: int = DEFAULT_MAX_ISSUE_NEWS,
) -> str:
    if max_issue_news and max_issue_news > 0:
        pre = sorted([c for c in candidates if c.timing == "pre_book"], key=lambda x: x.published_at, reverse=True)
        post = sorted([c for c in candidates if c.timing == "post_book"], key=lambda x: x.published_at)
        selected_candidates = pre[: max_issue_news // 2] + post[: max_issue_news - len(pre[: max_issue_news // 2])]
        selected_candidates = sorted({c.article_row_id: c for c in selected_candidates}.values(), key=lambda x: x.published_at)
    else:
        selected_candidates = sorted(candidates, key=lambda x: x.published_at)

    article_blocks: list[str] = []
    issue_number = extract_issue_number(issue.paper)
    for idx, candidate in enumerate(selected_candidates, start=1):
        text = candidate.text
        if max_article_chars and len(text) > max_article_chars:
            text = text[:max_article_chars]
        emission_matches_target = parsed_emission_matches_target(issue, candidate)
        article_blocks.append(
            "\n".join(
                [
                    f"[A{idx}]",
                    f"article_id: {candidate.article_id or 'NA'}",
                    f"published_at: {candidate.date_time}",
                    f"timing_relative_to_book_start: {candidate.timing}",
                    f"url: {candidate.url}",
                    f"title: {candidate.title}",
                    f"parsed_emissions: {candidate.emission_name or 'NA'}",
                    f"parsed_emission_urls: {candidate.emission_url or 'NA'}",
                    f"parsed_emission_matches_target: {'yes' if emission_matches_target else 'no'}",
                    "text:",
                    text or "NA",
                ]
            )
        )

    return "\n".join(
        [
            "ISSUE:",
            f"excel_row: {issue.excel_row}",
            f"paper: {issue.paper}",
            f"issue_number: {issue_number or 'NA'}",
            f"isin: {issue.isin or 'NA'}",
            f"emitent: {issue.emitent}",
            f"bookbuilding_start: {issue.bookbuilding_start.isoformat(sep=' ')}",
            f"placement_date_end: {issue.placement_date_end.isoformat()}",
            f"currency: {issue.currency or 'NA'}",
            "",
            "SELECTION_RULE:",
            "pre_value: последняя по времени релевантная новость с объемом размещения до bookbuilding_start.",
            "final_value: финальное значение объема размещения по итогам букбилдинга/размещения, обычно после bookbuilding_start.",
            "Возвращай значения в млрд валюты выпуска, без валютной конвертации.",
            "",
            "ARTICLES:",
            "\n\n".join(article_blocks) if article_blocks else "NA",
        ]
    )


def build_batch_prompt(batch: list[BatchItem]) -> str:
    return "\n\n".join(
        [
            "Ниже несколько выпусков. Для каждого блока ISSUE верни ровно один объект в JSON-массиве.",
            "Ключ сопоставления результата: excel_row.",
            "Для каждого выпуска анализируй только новости из его блока.",
            *(item.prompt_text for item in batch),
        ]
    )


def make_batches(items: list[BatchItem], max_batch_chars: int) -> list[list[BatchItem]]:
    batches: list[list[BatchItem]] = []
    batch_chars: list[int] = []
    for item in sorted(items, key=lambda x: x.estimated_chars, reverse=True):
        placed = False
        for idx, current_chars in enumerate(batch_chars):
            if current_chars + item.estimated_chars <= max_batch_chars:
                batches[idx].append(item)
                batch_chars[idx] += item.estimated_chars
                placed = True
                break
        if not placed:
            batches.append([item])
            batch_chars.append(item.estimated_chars)
    return batches


def extract_json_array(raw_text: str) -> str:
    raw_text = (raw_text or "").strip()
    if raw_text.startswith("```"):
        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text, flags=re.IGNORECASE)
        raw_text = re.sub(r"\s*```$", "", raw_text).strip()
    start = raw_text.find("[")
    end = raw_text.rfind("]")
    if start != -1 and end != -1 and end > start:
        return raw_text[start : end + 1]
    raise ValueError("JSON array not found in model output")


def parse_model_items(payload: Any) -> list[dict[str, Any]]:
    if isinstance(payload, list):
        parsed = payload
    elif isinstance(payload, str):
        parsed = json.loads(extract_json_array(payload))
    else:
        raise ValueError("Unsupported model payload type")
    if not isinstance(parsed, list):
        raise ValueError("Model output is not a list")
    return [item for item in parsed if isinstance(item, dict)]


def status_from_values(pre_value: float, final_value: float) -> str:
    if pre_value > 0 and final_value > 0:
        return "found"
    if pre_value > 0 or final_value > 0:
        return "partial"
    return "none"


def normalize_model_item(
    item: dict[str, Any],
    candidates: list[CandidateArticle],
    model_name: str,
    model_version: str | None,
) -> dict[str, Any]:
    pre_value = safe_nonneg_number(item.get("pre_value_bn"), 0.0)
    final_value = safe_nonneg_number(item.get("final_value_bn"), 0.0)

    def get_source(prefix: str) -> tuple[int | None, CandidateArticle | None]:
        try:
            source_id = int(item.get(f"{prefix}_source_article_id") or 0) or None
        except (TypeError, ValueError):
            source_id = None
        return source_id, source_by_article_id(candidates, source_id)

    pre_source_id, pre_source = get_source("pre")
    final_source_id, final_source = get_source("final")

    needs_attention = bool(item.get("needs_attention"))
    attention_reason = normalize_text(item.get("attention_reason")) if needs_attention else ""
    for label, source_id, source_candidate, value in [
        ("pre", pre_source_id, pre_source, pre_value),
        ("final", final_source_id, final_source, final_value),
    ]:
        if value > 0 and source_id is None:
            needs_attention = True
            attention_reason = " ".join(
                x
                for x in [
                    attention_reason,
                    f"Gemini вернул {label}-значение без source_article_id.",
                ]
                if x
            ).strip()
        elif value > 0 and source_candidate is None:
            needs_attention = True
            attention_reason = " ".join(
                x
                for x in [
                    attention_reason,
                    f"Gemini вернул {label}_source_article_id, которого нет в переданном списке новостей.",
                ]
                if x
            ).strip()

    if pre_value > 0 and pre_source is not None and pre_source.timing != "pre_book":
        needs_attention = True
        attention_reason = " ".join(
            x
            for x in [
                attention_reason,
                "pre-значение взято из новости не до начала букбилдинга.",
            ]
            if x
        ).strip()
    if final_value > 0 and final_source is not None and final_source.timing != "post_book":
        needs_attention = True
        attention_reason = " ".join(
            x
            for x in [
                attention_reason,
                "final-значение взято из новости до начала букбилдинга.",
            ]
            if x
        ).strip()

    status = normalize_text(item.get("status")).lower()
    if status not in {"found", "partial", "none"}:
        status = status_from_values(pre_value, final_value)
    else:
        status = status_from_values(pre_value, final_value)

    return {
        "pre_value_bn": pre_value,
        "pre_value_text": normalize_text(item.get("pre_value_text"))[:240],
        "pre_source_title": pre_source.title if pre_source else "",
        "pre_source_url": pre_source.url if pre_source else "",
        "pre_source_article_id": pre_source_id,
        "pre_source_datetime": pre_source.date_time if pre_source else "",
        "pre_evidence_text": normalize_text(item.get("pre_evidence_text"))[:900],
        "final_value_bn": final_value,
        "final_value_text": normalize_text(item.get("final_value_text"))[:240],
        "final_source_title": final_source.title if final_source else "",
        "final_source_url": final_source.url if final_source else "",
        "final_source_article_id": final_source_id,
        "final_source_datetime": final_source.date_time if final_source else "",
        "final_evidence_text": normalize_text(item.get("final_evidence_text"))[:900],
        "needs_attention": int(needs_attention),
        "attention_reason": attention_reason,
        "reason_short": normalize_text(item.get("reason_short"))[:300],
        "confidence": safe_float_01(item.get("confidence"), 0.5),
        "model_name": model_name,
        "model_version": model_version,
        "status": status,
    }


def has_extracted_volume(result: dict[str, Any]) -> bool:
    return safe_nonneg_number(result.get("pre_value_bn"), 0.0) > 0 or safe_nonneg_number(result.get("final_value_bn"), 0.0) > 0


def sleep_to_respect_min_interval(min_interval: float) -> None:
    global LAST_REQUEST_TS
    elapsed = time.time() - LAST_REQUEST_TS
    if LAST_REQUEST_TS and elapsed < min_interval:
        time.sleep(min_interval - elapsed)


def mark_request_sent() -> None:
    global LAST_REQUEST_TS
    LAST_REQUEST_TS = time.time()


def estimate_request_tokens(prompt: str, max_output_tokens: int, chars_per_token: float) -> int:
    chars_per_token = max(1.0, chars_per_token)
    schema_chars = len(json.dumps(RESULT_SCHEMA, ensure_ascii=False))
    input_chars = len(prompt) + len(SYSTEM_PROMPT) + schema_chars
    return math.ceil(input_chars / chars_per_token) + max_output_tokens


def prune_token_window(now: float) -> None:
    while TOKEN_WINDOW and now - TOKEN_WINDOW[0][0] >= 60.0:
        TOKEN_WINDOW.popleft()


def wait_for_tpm_capacity(estimated_tokens: int, tpm_limit: int, tpm_safety: float) -> None:
    if tpm_limit <= 0:
        return

    allowed_tokens = max(1, math.floor(tpm_limit * min(max(tpm_safety, 0.01), 1.0)))
    if estimated_tokens > allowed_tokens:
        raise ValueError(
            "Один запрос оценивается выше разрешенного TPM-окна: "
            f"{estimated_tokens:,} токенов > {allowed_tokens:,}. "
            "Уменьши --max-batch-chars или подними --tpm-limit/--tpm-safety."
        )

    while True:
        now = time.time()
        prune_token_window(now)
        used_tokens = sum(tokens for _, tokens in TOKEN_WINDOW)
        if used_tokens + estimated_tokens <= allowed_tokens:
            TOKEN_WINDOW.append((now, estimated_tokens))
            return

        wait_seconds = max(0.1, 60.0 - (now - TOKEN_WINDOW[0][0]) + 0.1)
        print(
            "  TPM limiter: "
            f"used={used_tokens:,}, next={estimated_tokens:,}, allowed={allowed_tokens:,}; "
            f"sleep {wait_seconds:.1f}s"
        )
        time.sleep(wait_seconds)


def extract_retry_delay_seconds(exc: Exception) -> float | None:
    text = str(exc)
    for pattern in (r"retry in\s+([0-9]+(?:\.[0-9]+)?)s", r"Retry in\s+([0-9]+(?:\.[0-9]+)?)s"):
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return float(match.group(1))
    return None


def call_gemini(
    client: Any,
    model_name: str,
    prompt: str,
    max_retries: int,
    max_output_tokens: int,
    delay: float,
    tpm_limit: int,
    tpm_safety: float,
    chars_per_token: float,
) -> tuple[Any, str | None]:
    last_exc: Exception | None = None
    estimated_tokens = estimate_request_tokens(prompt, max_output_tokens, chars_per_token)
    for attempt in range(1, max_retries + 1):
        try:
            sleep_to_respect_min_interval(delay)
            wait_for_tpm_capacity(estimated_tokens, tpm_limit, tpm_safety)
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    temperature=0,
                    max_output_tokens=max_output_tokens,
                    response_mime_type="application/json",
                    response_json_schema=RESULT_SCHEMA,
                ),
            )
            mark_request_sent()
            parsed = getattr(response, "parsed", None)
            raw_text = getattr(response, "text", "") or ""
            version = getattr(response, "model_version", None)
            return (parsed if parsed is not None else raw_text), version
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
            wait = extract_retry_delay_seconds(exc)
            if wait is None:
                wait = min(5 * (2 ** (attempt - 1)), 60)
            if attempt < max_retries:
                print(f"  retry {attempt}/{max_retries} after {wait:.1f}s: {exc}")
                time.sleep(wait)
    raise RuntimeError(str(last_exc) if last_exc else "Unknown API error")


def empty_result(status: str, reason: str) -> dict[str, Any]:
    return {
        "pre_value_bn": 0.0,
        "pre_value_text": "",
        "pre_source_title": "",
        "pre_source_url": "",
        "pre_source_article_id": None,
        "pre_source_datetime": "",
        "pre_evidence_text": "",
        "final_value_bn": 0.0,
        "final_value_text": "",
        "final_source_title": "",
        "final_source_url": "",
        "final_source_article_id": None,
        "final_source_datetime": "",
        "final_evidence_text": "",
        "needs_attention": 1 if status not in {"no_candidate", "excluded_veb"} else 0,
        "attention_reason": reason,
        "reason_short": reason,
        "confidence": 0.0,
        "model_name": "",
        "model_version": "",
        "status": status,
    }


def apply_cache_to_workbook(
    ws,
    headers: dict[str, int],
    cache: dict[str, Any],
    issues: dict[int, IssueInfo],
    row_filter: set[int] | None = None,
) -> None:
    mapping = {
        "volume_gemini_pre_value_bn": "pre_value_bn",
        "volume_gemini_pre_value_text": "pre_value_text",
        "volume_gemini_pre_source_title": "pre_source_title",
        "volume_gemini_pre_source_url": "pre_source_url",
        "volume_gemini_pre_source_article_id": "pre_source_article_id",
        "volume_gemini_pre_source_datetime": "pre_source_datetime",
        "volume_gemini_pre_evidence_text": "pre_evidence_text",
        "volume_gemini_final_value_bn": "final_value_bn",
        "volume_gemini_final_value_text": "final_value_text",
        "volume_gemini_final_source_title": "final_source_title",
        "volume_gemini_final_source_url": "final_source_url",
        "volume_gemini_final_source_article_id": "final_source_article_id",
        "volume_gemini_final_source_datetime": "final_source_datetime",
        "volume_gemini_final_evidence_text": "final_evidence_text",
        "volume_gemini_needs_attention": "needs_attention",
        "volume_gemini_attention_reason": "attention_reason",
        "volume_gemini_reason_short": "reason_short",
        "volume_gemini_confidence": "confidence",
        "volume_gemini_model_name": "model_name",
        "volume_gemini_model_version": "model_version",
        "volume_gemini_status": "status",
    }
    for key, result in cache.items():
        try:
            row_idx = int(key)
        except ValueError:
            continue
        if row_filter is not None and row_idx not in row_filter:
            continue
        if not isinstance(result, dict) or row_idx < 2 or row_idx > ws.max_row:
            continue
        for col_name, result_key in mapping.items():
            ws.cell(row_idx, headers[normalize_key(col_name)], result.get(result_key))
        issue = issues.get(row_idx)
        if issue is not None:
            ws.cell(row_idx, headers[normalize_key("volume_gemini_currency")], issue.currency)


def save_workbook(
    wb: openpyxl.Workbook,
    ws,
    cache: dict[str, Any],
    issues: dict[int, IssueInfo],
    output_path: Path,
    row_filter: set[int] | None = None,
) -> None:
    headers = ensure_output_columns(ws)
    apply_cache_to_workbook(ws, headers, cache, issues, row_filter=row_filter)
    try:
        wb.save(output_path)
    except ValueError as exc:
        if "closed file" not in str(exc).lower():
            raise
        strip_workbook_images(wb)
        wb.save(output_path)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Extract placement volume guidance/final values from Cbonds news with Gemini.")
    parser.add_argument("--input", type=Path, default=Path(INPUT_FILE))
    parser.add_argument("--output", type=Path, default=Path(OUTPUT_FILE))
    parser.add_argument("--db", type=Path, default=Path(DB_FILE))
    parser.add_argument("--results-table", type=str, default=RESULTS_TABLE)
    parser.add_argument("--cache", type=Path, default=Path(CACHE_FILE))
    parser.add_argument("--sheet", type=str, default=DEFAULT_SHEET)
    parser.add_argument("--paper-column", default=DEFAULT_PAPER_COLUMN)
    parser.add_argument("--isin-column", default=DEFAULT_ISIN_COLUMN)
    parser.add_argument("--issuer-column", default=DEFAULT_ISSUER_COLUMN)
    parser.add_argument("--book-date-column", default=DEFAULT_BOOK_DATE_COLUMN)
    parser.add_argument("--placement-end-column", default=DEFAULT_PLACEMENT_END_COLUMN)
    parser.add_argument("--model", type=str, default=DEFAULT_MODEL)
    parser.add_argument("--api-key", type=str, default=os.getenv("GEMINI_API_KEY", ""))
    parser.add_argument("--delay", type=float, default=DEFAULT_DELAY)
    parser.add_argument(
        "--tpm-limit",
        type=int,
        default=DEFAULT_TPM_LIMIT,
        help="Лимит токенов в минуту. 0 отключает TPM-лимитер. Можно задать GEMINI_TPM_LIMIT.",
    )
    parser.add_argument(
        "--tpm-safety",
        type=float,
        default=DEFAULT_TPM_SAFETY,
        help="Доля TPM-лимита, которую разрешено использовать, по умолчанию 0.9.",
    )
    parser.add_argument(
        "--chars-per-token",
        type=float,
        default=DEFAULT_CHARS_PER_TOKEN,
        help="Консервативная оценка знаков на токен, по умолчанию 2.0.",
    )
    parser.add_argument("--max-retries", type=int, default=DEFAULT_MAX_RETRIES)
    parser.add_argument("--max-output-tokens", type=int, default=DEFAULT_MAX_OUTPUT_TOKENS)
    parser.add_argument("--max-requests", type=int, default=None)
    parser.add_argument("--max-batch-chars", type=int, default=DEFAULT_MAX_BATCH_CHARS)
    parser.add_argument("--max-article-chars", type=int, default=DEFAULT_MAX_ARTICLE_CHARS)
    parser.add_argument("--max-issue-news", type=int, default=DEFAULT_MAX_ISSUE_NEWS)
    parser.add_argument("--issue-char-threshold", type=int, default=DEFAULT_MAX_BATCH_CHARS)
    parser.add_argument(
        "--issue-size-mode",
        choices=["all", "under", "over"],
        default="all",
        help="all: все выпуски; under: только выпуски <= threshold; over: только выпуски > threshold.",
    )
    parser.add_argument("--one-issue-per-batch", action="store_true", help="Не объединять выпуски в батчи.")
    parser.add_argument("--limit-issues", type=int, default=None, help="Ограничить число выпусков после фильтров.")
    parser.add_argument("--save-every", type=int, default=DEFAULT_SAVE_EVERY)
    parser.add_argument("--refresh", action="store_true")
    parser.add_argument(
        "--retry-volume-problems",
        action="store_true",
        help=(
            "Заново обработать только строки, где volume_gemini_pre_value_bn=0 "
            "или volume_gemini_final_value_bn=0 или volume_issue_number_mismatch=1. "
            "Старый кэш для этих строк игнорируется; батчи принудительно по одному выпуску."
        ),
    )
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--no-confirm", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    wb, ws, issues, skipped_veb = load_excel_issues(
        args.input,
        sheet_name=args.sheet,
        paper_column=args.paper_column,
        isin_column=args.isin_column,
        issuer_column=args.issuer_column,
        book_date_column=args.book_date_column,
        placement_end_column=args.placement_end_column,
    )
    loaded_issue_count = len(issues)
    retry_problem_rows: set[int] = set()
    if args.retry_volume_problems:
        retry_problem_rows = retry_problem_rows_from_workbook(ws, issues)
        issues = {row_idx: issue for row_idx, issue in issues.items() if row_idx in retry_problem_rows}
    if args.limit_issues is not None:
        issues = dict(list(sorted(issues.items()))[: args.limit_issues])
    workbook_row_filter = set(issues) if args.retry_volume_problems else None

    candidates_by_issue = load_article_candidates(args.db, issues)
    cache = load_json(args.cache)

    no_candidate_count = 0
    for row_idx in issues:
        if not candidates_by_issue.get(row_idx):
            no_candidate_count += 1
            no_candidate_result = empty_result("no_candidate", "Не найдены привязанные новости в нужном окне.")
            if args.retry_volume_problems or args.refresh:
                cache[str(row_idx)] = no_candidate_result
            else:
                cache.setdefault(str(row_idx), no_candidate_result)

    rows_with_candidates = [row_idx for row_idx in sorted(issues) if candidates_by_issue.get(row_idx)]
    if args.retry_volume_problems:
        pending_rows = rows_with_candidates
    else:
        pending_rows = [row_idx for row_idx in rows_with_candidates if args.refresh or str(row_idx) not in cache]
    pending_items: list[BatchItem] = []
    for row_idx in pending_rows:
        prompt_text = build_issue_prompt_block(
            issues[row_idx],
            candidates_by_issue[row_idx],
            max_article_chars=args.max_article_chars,
            max_issue_news=args.max_issue_news,
        )
        item = BatchItem(
            excel_row=row_idx,
            prompt_text=prompt_text,
            estimated_chars=len(prompt_text) + 450,
        )
        if args.issue_size_mode == "under" and item.estimated_chars > args.issue_char_threshold:
            continue
        if args.issue_size_mode == "over" and item.estimated_chars <= args.issue_char_threshold:
            continue
        pending_items.append(item)

    one_issue_per_batch = args.one_issue_per_batch or args.retry_volume_problems
    batches = [[item] for item in pending_items] if one_issue_per_batch else make_batches(pending_items, args.max_batch_chars)
    if args.max_requests is not None:
        batches = batches[: args.max_requests]

    print("=" * 72)
    print("Placement volume extraction from Cbonds news")
    print("=" * 72)
    print(f"Input: {args.input}")
    print(f"Output: {args.output}")
    print(f"DB: {args.db}")
    print(f"Results table: {args.results_table}")
    print(f"Sheet: {ws.title}")
    print(f"Paper column: {args.paper_column}")
    print(f"Issuer column: {args.issuer_column}")
    print(f"Book start column: {args.book_date_column}")
    print(f"Placement end column: {args.placement_end_column}")
    print(f"Issues loaded excluding VEB.RF: {loaded_issue_count:,}")
    print(f"VEB.RF skipped: {skipped_veb:,}")
    if args.retry_volume_problems:
        print(f"Retry volume problems: yes | problem rows before candidate filter: {len(retry_problem_rows):,}")
    print(f"Issues eligible after limit: {len(issues):,}")
    print(f"Issues with candidate news: {len(rows_with_candidates):,}")
    print(f"No candidate: {no_candidate_count:,}")
    print(f"Cache rows: {len(cache):,}")
    print(f"Pending issues: {len(pending_rows):,}")
    print(f"Pending issues after size filter: {len(pending_items):,}")
    print(f"Planned batches: {len(batches):,}")
    print(f"Max batch chars: {args.max_batch_chars:,}")
    print(f"Max article chars: {args.max_article_chars:,}")
    print(f"Max issue news: {args.max_issue_news:,}")
    print(f"Issue size mode: {args.issue_size_mode} | threshold: {args.issue_char_threshold:,}")
    print(f"One issue per batch: {one_issue_per_batch}")
    print(f"TPM limit: {args.tpm_limit:,}" if args.tpm_limit else "TPM limit: disabled")
    if args.tpm_limit:
        allowed_tokens = math.floor(args.tpm_limit * min(max(args.tpm_safety, 0.01), 1.0))
        print(f"TPM safety: {args.tpm_safety:g} | allowed window: {allowed_tokens:,}")
        print(f"Chars per token estimate: {args.chars_per_token:g}")

    total_candidate_news = sum(len(items) for items in candidates_by_issue.values())
    pre_news_count = sum(1 for items in candidates_by_issue.values() for c in items if c.timing == "pre_book")
    post_news_count = sum(1 for items in candidates_by_issue.values() for c in items if c.timing == "post_book")
    print(f"Candidate news total: {total_candidate_news:,}")
    print(f"Candidate news pre-book: {pre_news_count:,}")
    print(f"Candidate news post-book: {post_news_count:,}")
    if batches:
        batch_prompts = [build_batch_prompt(batch) for batch in batches]
        batch_chars = [len(prompt) for prompt in batch_prompts]
        print(f"Batch chars avg: {sum(batch_chars) / len(batch_chars):,.1f}")
        print(f"Batch chars max: {max(batch_chars):,}")
        if args.tpm_limit:
            allowed_tokens = math.floor(args.tpm_limit * min(max(args.tpm_safety, 0.01), 1.0))
            token_estimates = [
                estimate_request_tokens(prompt, args.max_output_tokens, args.chars_per_token)
                for prompt in batch_prompts
            ]
            too_large = sum(1 for tokens in token_estimates if tokens > allowed_tokens)
            print(f"Batch token estimate avg: {sum(token_estimates) / len(token_estimates):,.1f}")
            print(f"Batch token estimate max: {max(token_estimates):,}")
            print(f"Batches above TPM allowed window: {too_large:,}")
            if too_large:
                print("WARNING: at least one batch cannot fit into one TPM window and will fail before sending.")

    if args.dry_run:
        save_workbook(wb, ws, cache, issues, args.output, row_filter=workbook_row_filter)
        write_cache_to_db(args.db, issues, cache, table_name=args.results_table)
        print("\nDry run complete. Workbook/DB written from existing cache only.")
        return

    if genai is None or types is None:
        raise ImportError("Не найден пакет google-genai. Установи его: pip install google-genai")
    if not args.api_key:
        raise ValueError("API key not found. Pass --api-key or set GEMINI_API_KEY.")

    if batches and not args.no_confirm:
        answer = input(f"Send {len(batches)} batched requests to Gemini? (y/n): ").strip().lower()
        if answer != "y":
            print("Cancelled by user.")
            save_workbook(wb, ws, cache, issues, args.output, row_filter=workbook_row_filter)
            write_cache_to_db(args.db, issues, cache, table_name=args.results_table)
            return

    client = genai.Client(api_key=args.api_key)
    try:
        for idx, batch in enumerate(batches, start=1):
            batch_rows = [item.excel_row for item in batch]
            print(f"\nBatch {idx}/{len(batches)} | issues={len(batch)} | rows={batch_rows[:8]}{'...' if len(batch_rows) > 8 else ''}")
            try:
                payload, model_version = call_gemini(
                    client=client,
                    model_name=args.model,
                    prompt=build_batch_prompt(batch),
                    max_retries=args.max_retries,
                    max_output_tokens=args.max_output_tokens,
                    delay=args.delay,
                    tpm_limit=args.tpm_limit,
                    tpm_safety=args.tpm_safety,
                    chars_per_token=args.chars_per_token,
                )
                parsed_items = parse_model_items(payload)
                results_by_row: dict[int, dict[str, Any]] = {}
                for item in parsed_items:
                    try:
                        item_row = int(item.get("excel_row"))
                    except (TypeError, ValueError):
                        continue
                    if item_row not in batch_rows:
                        continue
                    results_by_row[item_row] = normalize_model_item(
                        item,
                        candidates_by_issue[item_row],
                        args.model,
                        model_version,
                    )

                for row_idx in batch_rows:
                    result = results_by_row.get(row_idx)
                    if result is None:
                        result = empty_result("missing_item_in_response", "Gemini не вернул объект для этой строки в батче.")
                    cache[str(row_idx)] = result

                extracted_now = sum(1 for r in results_by_row.values() if has_extracted_volume(r))
                print(f"  -> returned={len(results_by_row)}/{len(batch_rows)} extracted_any={extracted_now}")
            except Exception as exc:  # noqa: BLE001
                for row_idx in batch_rows:
                    cache[str(row_idx)] = empty_result("api_error", f"{type(exc).__name__}: {exc}")
                print(f"  [error] {exc}")

            if idx % args.save_every == 0:
                save_json(args.cache, cache)
                save_workbook(wb, ws, cache, issues, args.output, row_filter=workbook_row_filter)
                write_cache_to_db(args.db, issues, cache, table_name=args.results_table)
                print(f"  saved: {args.cache}, {args.output}")

    except KeyboardInterrupt:
        print("\nInterrupted. Saving progress...")
    finally:
        save_json(args.cache, cache)
        save_workbook(wb, ws, cache, issues, args.output, row_filter=workbook_row_filter)
        write_cache_to_db(args.db, issues, cache, table_name=args.results_table)
        print(f"\nSaved cache: {args.cache}")
        print(f"Saved workbook: {args.output}")
        print(f"Saved DB table: {args.results_table}")

    extracted_pre = sum(1 for item in cache.values() if isinstance(item, dict) and safe_nonneg_number(item.get("pre_value_bn")) > 0)
    extracted_final = sum(1 for item in cache.values() if isinstance(item, dict) and safe_nonneg_number(item.get("final_value_bn")) > 0)
    attention = sum(1 for item in cache.values() if isinstance(item, dict) and int(item.get("needs_attention") or 0) == 1)
    print(f"Extracted pre: {extracted_pre:,}")
    print(f"Extracted final: {extracted_final:,}")
    print(f"Needs attention: {attention:,}")


if __name__ == "__main__":
    main()
