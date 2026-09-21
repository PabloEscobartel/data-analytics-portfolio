#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сбор периода книги заявок с cbonds.ru.

Сделано по образцу la_finale/cbonds_scrape.py, но вместо способа/типа/формата
размещения парсится поле:

    Размещение -> Книга заявок
    li#cb_bond_page_placement_book

Вход:
    bonds_offer_updated.xlsx, лист main_no_na

Выход:
    bonds_offer_updated_book_period.xlsx

Кэш:
    cbonds_book_period_cache.json

Установка зависимостей:
    pip install requests openpyxl beautifulsoup4 lxml

Запуск:
    python scrape_cbonds_book_period.py

Тест на первых 10 ISIN:
    python scrape_cbonds_book_period.py --limit 10
"""

from __future__ import annotations

import argparse
import json
import os
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl
import requests
from bs4 import BeautifulSoup


INPUT_FILE = "bonds_offer_updated_cbonds_orientir_gemini_retry.xlsx"
OUTPUT_FILE = "bonds_offer_updated_cbonds_orientir_gemini_retry_book.xlsx"
SHEET_NAME = "place_before_2018"
CACHE_FILE = "cbonds_book_period_cache.json"
COOKIE_FILE = "cbonds_cookie.txt"

DELAY = 0.7
SUGGEST_URL = "https://cbonds.ru/api/suggest/bonds/"
BOND_PAGE_URL = "https://cbonds.ru/bonds/{bond_id}/"
BOOK_PERIOD_FIELD_ID = "cb_bond_page_placement_book"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/147.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.7,en;q=0.6",
}

OUTPUT_COLUMNS = [
    "cbonds_id",
    "cbonds_book_period",
    "cbonds_book_start_date",
    "cbonds_book_start_time",
    "cbonds_book_start_datetime",
    "cbonds_book_end_date",
    "cbonds_book_end_time",
    "cbonds_book_end_datetime",
    "cbonds_book_status",
]


session = requests.Session()
session.headers.update(HEADERS)


def load_cookie_header() -> str | None:
    env_cookie = os.getenv("CBONDS_COOKIE", "").strip()
    if env_cookie:
        return env_cookie

    cookie_path = Path(COOKIE_FILE)
    if cookie_path.exists():
        text = cookie_path.read_text(encoding="utf-8").strip()
        return text or None

    return None


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def load_cache(cache_path: Path) -> dict[str, dict[str, Any]]:
    if cache_path.exists():
        with open(cache_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache_path: Path, cache: dict[str, dict[str, Any]]) -> None:
    tmp = cache_path.with_suffix(cache_path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    os.replace(tmp, cache_path)


def build_header_index(ws) -> dict[str, int]:
    return {normalize_text(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1)}


def ensure_output_columns(ws) -> dict[str, int]:
    headers = build_header_index(ws)
    for name in OUTPUT_COLUMNS:
        if name not in headers:
            col = ws.max_column + 1
            ws.cell(1, col, name)
            headers[name] = col
    return headers


def get_cbonds_id(isin: str) -> int | None:
    """ISIN -> cbonds ID через suggest API."""
    try:
        resp = session.post(
            SUGGEST_URL,
            data=isin,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=20,
        )
        resp.raise_for_status()
        items = resp.json().get("items", [])
        if not items:
            return None
        return int(items[0]["id"])
    except Exception as exc:  # noqa: BLE001
        print(f"    [suggest error] {isin}: {exc}")
        return None


def parse_book_period_parts(value: str) -> dict[str, str | None]:
    text = normalize_text(value)
    if not text:
        return {
            "cbonds_book_start_date": None,
            "cbonds_book_start_time": None,
            "cbonds_book_start_datetime": None,
            "cbonds_book_end_date": None,
            "cbonds_book_end_time": None,
            "cbonds_book_end_datetime": None,
        }

    parts = re.findall(
        r"(\d{2}\.\d{2}\.\d{4})(?:\s*\((\d{1,2}:\d{2})\))?",
        text,
    )

    parsed: list[tuple[str | None, str | None, str | None]] = []
    for date_text, time_text in parts[:2]:
        try:
            parsed_date = datetime.strptime(date_text, "%d.%m.%Y").date().isoformat()
        except ValueError:
            parsed_date = None

        parsed_time = None
        if time_text:
            try:
                parsed_time = datetime.strptime(time_text, "%H:%M").time().strftime("%H:%M")
            except ValueError:
                parsed_time = None

        parsed_datetime = f"{parsed_date} {parsed_time}" if parsed_date and parsed_time else None
        parsed.append((parsed_date, parsed_time, parsed_datetime))

    start_date, start_time, start_datetime = parsed[0] if len(parsed) >= 1 else (None, None, None)
    end_date, end_time, end_datetime = parsed[1] if len(parsed) >= 2 else (start_date, None, None)

    return {
        "cbonds_book_start_date": start_date,
        "cbonds_book_start_time": start_time,
        "cbonds_book_start_datetime": start_datetime,
        "cbonds_book_end_date": end_date,
        "cbonds_book_end_time": end_time,
        "cbonds_book_end_datetime": end_datetime,
    }


def parse_book_period_dates(value: str) -> tuple[str | None, str | None]:
    """Backward-compatible helper: returns only start/end dates."""
    parts = parse_book_period_parts(value)
    return parts["cbonds_book_start_date"], parts["cbonds_book_end_date"]


def parse_bond_page(bond_id: int) -> dict[str, Any]:
    """Загружает страницу бумаги и парсит период книги заявок."""
    result: dict[str, Any] = {
        "cbonds_id": bond_id,
        "cbonds_book_period": None,
        "cbonds_book_start_date": None,
        "cbonds_book_start_time": None,
        "cbonds_book_start_datetime": None,
        "cbonds_book_end_date": None,
        "cbonds_book_end_time": None,
        "cbonds_book_end_datetime": None,
        "cbonds_book_status": "not_found",
    }

    try:
        resp = session.get(BOND_PAGE_URL.format(bond_id=bond_id), timeout=25)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")

        li = soup.find("li", id=BOOK_PERIOD_FIELD_ID)
        if not li:
            return result

        value_div = li.find("div", class_="value")
        if not value_div:
            return result

        span = value_div.find("span")
        raw_period = normalize_text(span.get_text(" ", strip=True) if span else value_div.get_text(" ", strip=True))
        parsed_period = parse_book_period_parts(raw_period)

        result.update(
            {
                "cbonds_book_period": raw_period or None,
                "cbonds_book_status": "ok" if raw_period else "empty",
                **parsed_period,
            }
        )
        return result

    except Exception as exc:  # noqa: BLE001
        print(f"    [page error] bond_id={bond_id}: {exc}")
        result["cbonds_book_status"] = f"page_error: {exc}"
        return result


def process_isin(isin: str) -> dict[str, Any]:
    bond_id = get_cbonds_id(isin)
    if bond_id is None:
        return {
            "cbonds_id": None,
            "cbonds_book_period": None,
            "cbonds_book_start_date": None,
            "cbonds_book_start_time": None,
            "cbonds_book_start_datetime": None,
            "cbonds_book_end_date": None,
            "cbonds_book_end_time": None,
            "cbonds_book_end_datetime": None,
            "cbonds_book_status": "suggest_not_found",
        }

    time.sleep(DELAY)
    return parse_bond_page(bond_id)


def collect_unique_isins(ws, isin_col: int) -> list[str]:
    unique = set()
    for row in range(2, ws.max_row + 1):
        isin = normalize_text(ws.cell(row, isin_col).value)
        if isin:
            unique.add(isin)
    return sorted(unique)


def write_results(ws, headers: dict[str, int], cache: dict[str, dict[str, Any]], isin_col: int) -> None:
    for row in range(2, ws.max_row + 1):
        isin = normalize_text(ws.cell(row, isin_col).value)
        data = cache.get(isin, {})
        for col_name in OUTPUT_COLUMNS:
            ws.cell(row, headers[col_name], data.get(col_name))


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Собрать период книги заявок с Cbonds по ISIN.")
    parser.add_argument("--input", type=Path, default=Path(INPUT_FILE))
    parser.add_argument("--output", type=Path, default=Path(OUTPUT_FILE))
    parser.add_argument("--sheet", default=SHEET_NAME)
    parser.add_argument("--cache", type=Path, default=Path(CACHE_FILE))
    parser.add_argument("--limit", type=int, default=None, help="Ограничить число новых ISIN для тестового запуска.")
    parser.add_argument("--refresh", action="store_true", help="Перепарсить даже те ISIN, которые уже есть в кэше.")
    parser.add_argument("--delay", type=float, default=DELAY)
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    global DELAY
    DELAY = args.delay

    cookie_header = load_cookie_header()
    if cookie_header:
        session.headers["Cookie"] = cookie_header

    if not args.input.exists():
        raise FileNotFoundError(f"Не найден входной файл: {args.input.resolve()}")

    wb = openpyxl.load_workbook(args.input)
    if args.sheet not in wb.sheetnames:
        raise RuntimeError(f"Лист {args.sheet!r} не найден. Доступные листы: {', '.join(wb.sheetnames)}")
    ws = wb[args.sheet]

    headers = build_header_index(ws)
    if "ISIN" not in headers:
        raise RuntimeError(f"На листе {args.sheet!r} не найдена колонка ISIN")
    isin_col = headers["ISIN"]

    unique_isins = collect_unique_isins(ws, isin_col)
    cache = load_cache(args.cache)

    if args.refresh:
        todo = unique_isins
    else:
        todo = [isin for isin in unique_isins if isin not in cache]
    if args.limit is not None:
        todo = todo[: args.limit]

    print(f"Файл: {args.input}")
    print(f"Лист: {args.sheet}")
    print(f"Уникальных ISIN: {len(unique_isins)}")
    print(f"Уже в кэше: {sum(1 for isin in unique_isins if isin in cache)}")
    print(f"К обработке сейчас: {len(todo)}")
    print(f"Cookie: {'загружены' if 'Cookie' in session.headers else 'не заданы'}")
    print()

    for idx, isin in enumerate(todo, start=1):
        print(f"[{idx}/{len(todo)}] {isin}", end=" -> ")
        data = process_isin(isin)
        cache[isin] = data
        print(
            f"id={data.get('cbonds_id')} | "
            f"book={data.get('cbonds_book_period') or '-'} | "
            f"status={data.get('cbonds_book_status')}"
        )

        if idx % 20 == 0:
            save_cache(args.cache, cache)
            print(f"    [кэш сохранён: {len(cache)} записей]")

        if idx < len(todo):
            time.sleep(DELAY)

    save_cache(args.cache, cache)
    print(f"\nКэш финализирован: {len(cache)} записей")

    headers = ensure_output_columns(ws)
    write_results(ws, headers, cache, isin_col)
    wb.save(args.output)
    print(f"Готово: {args.output}")


if __name__ == "__main__":
    main()
