#!/usr/bin/env python3
"""
Тест Cbonds news API для одного эмитента и одной даты.

Что делает:
1. Через /api/suggest/companies/ получает emitent_id по названию эмитента.
2. Через /api/news/ получает новости по этому emitent_id и дате.
3. Печатает краткий результат или сырой JSON.

Cookie:
- Можно положить заголовок Cookie целиком в файл cbonds_cookie.txt рядом со скриптом,
  либо передать его через переменную окружения CBONDS_COOKIE.

Примеры:
    python cbonds_news_api_test.py "АВТОБАН-Финанс" 2025-11-17
    python cbonds_news_api_test.py "А101" 2025-12-24 --raw
    python cbonds_news_api_test.py "ВТБ-Лизинг" 2025-12-24 --limit 10
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path
from typing import Any
from urllib.parse import quote

import requests

SITE_BASE_URL = "https://cbonds.ru"
SUGGEST_COMPANIES_URL = f"{SITE_BASE_URL}/api/suggest/companies/"
NEWS_API_URL = f"{SITE_BASE_URL}/api/news/"
COOKIE_FILE = "cbonds_cookie.txt"
TIMEOUT = 25

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/147.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-GB,en;q=0.9,ru-RU;q=0.8,ru;q=0.7,en-US;q=0.6",
}


def normalize_text(value: str | None) -> str:
    if not value:
        return ""
    value = str(value).strip().casefold().replace("ё", "е")
    return re.sub(r"\s+", " ", value)


def load_cookie_header() -> str | None:
    env_cookie = os.getenv("CBONDS_COOKIE", "").strip()
    if env_cookie:
        return env_cookie

    cookie_path = Path(COOKIE_FILE)
    if cookie_path.exists():
        text = cookie_path.read_text(encoding="utf-8").strip()
        return text or None

    return None


def parse_json_response(resp: requests.Response) -> dict[str, Any]:
    try:
        return resp.json()
    except Exception:
        text = resp.text.lstrip("\ufeff\n\r\t ")
        return json.loads(text)


def build_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(HEADERS)
    cookie = load_cookie_header()
    if cookie:
        session.headers["Cookie"] = cookie
    return session


def choose_best_emitent_match(query: str, items: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not items:
        return None

    q_norm = normalize_text(query)
    exact_matches: list[dict[str, Any]] = []
    startswith_matches: list[dict[str, Any]] = []

    for item in items:
        ttl = str(item.get("ttl") or "").strip()
        ttl_norm = normalize_text(ttl)
        if ttl_norm == q_norm:
            exact_matches.append(item)
        elif q_norm and ttl_norm.startswith(q_norm):
            startswith_matches.append(item)

    if exact_matches:
        return exact_matches[0]
    if startswith_matches:
        return startswith_matches[0]
    return items[0]


def get_emitent_match(session: requests.Session, emitent_name: str) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    payload = {
        "term": quote(emitent_name, safe=""),
        "emitent_stop_statuses_ids": [],
    }

    resp = session.post(
        SUGGEST_COMPANIES_URL,
        json=payload,
        timeout=TIMEOUT,
        headers={
            "Accept": "application/json, text/plain, */*",
            "Content-Type": "application/json",
            "Origin": SITE_BASE_URL,
            "Referer": f"{SITE_BASE_URL}/",
        },
    )
    resp.raise_for_status()
    data = parse_json_response(resp)
    items = list(data.get("items") or [])
    chosen = choose_best_emitent_match(emitent_name, items)
    if not chosen:
        raise RuntimeError(f"Эмитент не найден: {emitent_name}")
    return chosen, items


def build_news_payload(emitent_id: int, date_str: str, limit: int, page: int) -> dict[str, Any]:
    return {
        "filters": [
            {"field": "show_only_loans", "operator": "eq", "value": 0},
            {"field": "emitent_id", "operator": "in", "value": [str(emitent_id)]},
            {"field": "date", "operator": "ge", "value": date_str},
            {"field": "date", "operator": "le", "value": date_str},
        ],
        "lang": "rus",
        "quantity": {
            "offset": (page - 1) * limit,
            "limit": limit,
            "page": page,
        },
        "sorting": [],
    }


def fetch_news(session: requests.Session, emitent_id: int, date_str: str, limit: int, page: int) -> dict[str, Any]:
    payload = build_news_payload(emitent_id, date_str, limit, page)
    resp = session.post(
        NEWS_API_URL,
        json=payload,
        timeout=TIMEOUT,
        headers={
            "Accept": "application/json, text/plain, */*",
            "Content-Type": "application/json",
            "Origin": SITE_BASE_URL,
            "Referer": (
                f"{SITE_BASE_URL}/news/?emitent_id[]={emitent_id}"
                f"&date_min={date_str}&date_max={date_str}&page={page}"
            ),
        },
    )
    resp.raise_for_status()
    return parse_json_response(resp)


def print_match_info(query: str, chosen: dict[str, Any], items: list[dict[str, Any]]) -> None:
    print(f"Запрос: {query}")
    print(f"Endpoint suggest: {SUGGEST_COMPANIES_URL}")
    print(f"Найдено совпадений: {len(items)}")
    print()
    print("Первые совпадения:")
    for i, item in enumerate(items[:10], start=1):
        print(
            f"  {i}. id={item.get('id')} | ttl={item.get('ttl')} "
            f"| emitent_statuses_id={item.get('emitent_statuses_id')}"
        )
    print()
    print("Выбранное совпадение:")
    print(f"  id  = {chosen.get('id')}")
    print(f"  ttl = {chosen.get('ttl')}")
    print()


def print_news_summary(data: dict[str, Any]) -> None:
    response = data.get("response") or {}
    items = list(response.get("items") or [])

    print(f"Endpoint news: {NEWS_API_URL}")
    print("Ответ API:")
    print(f"  count: {response.get('count')}")
    print(f"  total: {response.get('total')}")
    print(f"  limit: {response.get('limit')}")
    print(f"  page: {response.get('page')}")
    print(f"  items_count: {len(items)}")
    print()

    if not items:
        print("Новостей нет.")
        return

    print("Новости:")
    for i, item in enumerate(items, start=1):
        text = str(item.get("text") or "").strip()
        preview = re.sub(r"\s+", " ", text)
        if len(preview) > 180:
            preview = preview[:177] + "..."
        print(f"  {i}. id={item.get('id')} | {item.get('date_time')} | {item.get('caption')}")
        print(f"     cb_link={item.get('cb_link')}")
        print(f"     preview={preview}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Проверка Cbonds /api/news/ по эмитенту и дате")
    parser.add_argument("emitent_name", help="Название эмитента, например: АВТОБАН-Финанс")
    parser.add_argument("date", help="Дата в формате YYYY-MM-DD")
    parser.add_argument("--limit", type=int, default=30, help="Лимит на страницу для /api/news/")
    parser.add_argument("--page", type=int, default=1, help="Номер страницы для /api/news/")
    parser.add_argument("--raw", action="store_true", help="Показать сырой JSON ответа /api/news/")
    parser.add_argument(
        "--show-text",
        type=int,
        metavar="N",
        help="Показать полный text для N-й новости из ответа (нумерация с 1)",
    )
    args = parser.parse_args()

    session = build_session()

    try:
        chosen, items = get_emitent_match(session, args.emitent_name)
        print_match_info(args.emitent_name, chosen, items)

        emitent_id = int(chosen["id"])
        payload = build_news_payload(emitent_id, args.date, args.limit, args.page)
        print("Payload /api/news/:")
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        print()

        data = fetch_news(session, emitent_id, args.date, args.limit, args.page)

        if args.raw:
            print(json.dumps(data, ensure_ascii=False, indent=2))
            return 0

        print_news_summary(data)

        if args.show_text is not None:
            response = data.get("response") or {}
            news_items = list(response.get("items") or [])
            index = args.show_text - 1
            if index < 0 or index >= len(news_items):
                raise IndexError(
                    f"Нельзя показать text для новости #{args.show_text}: всего новостей {len(news_items)}"
                )
            item = news_items[index]
            print()
            print(f"Полный text новости #{args.show_text} (id={item.get('id')}):")
            print(str(item.get("text") or ""))

        if data.get("error", {}).get("err_no") not in (0, None):
            print()
            print("В ответе есть ошибка API:")
            print(json.dumps(data.get("error"), ensure_ascii=False, indent=2))

        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
