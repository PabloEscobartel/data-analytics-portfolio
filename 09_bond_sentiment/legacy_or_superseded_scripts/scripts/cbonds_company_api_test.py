import argparse
import json
import sys
from typing import Any
from urllib.parse import quote

import requests

SUGGEST_URL = "https://cbonds.ru/api/suggest/companies/"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/146.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "Origin": "https://cbonds.ru",
    "Referer": "https://cbonds.ru/",
}


def normalize_name(value: str) -> str:
    return " ".join(value.strip().lower().replace("ё", "е").split())


def fetch_company_suggest(session: requests.Session, name: str) -> dict[str, Any]:
    payload = {
        "term": quote(name, safe=""),
        "emitent_stop_statuses_ids": [],
    }

    response = session.post(SUGGEST_URL, json=payload, timeout=20)
    response.raise_for_status()
    return response.json()


def pick_best_item(items: list[dict[str, Any]], query_name: str) -> dict[str, Any] | None:
    if not items:
        return None

    q = normalize_name(query_name)

    exact = [x for x in items if normalize_name(str(x.get("ttl", ""))) == q]
    if exact:
        return exact[0]

    startswith = [x for x in items if normalize_name(str(x.get("ttl", ""))).startswith(q)]
    if startswith:
        return startswith[0]

    return items[0]


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Проверка suggest API Cbonds для получения emitent_id по названию эмитента"
    )
    parser.add_argument(
        "name",
        nargs="+",
        help="Название эмитента, например: А101 или ВТБ-Лизинг",
    )
    parser.add_argument(
        "--raw",
        action="store_true",
        help="Печатать полный JSON-ответ API",
    )
    args = parser.parse_args()

    query_name = " ".join(args.name)

    session = requests.Session()
    session.headers.update(HEADERS)

    print(f"Запрос: {query_name}")
    print(f"Endpoint: {SUGGEST_URL}")
    print()

    try:
        data = fetch_company_suggest(session, query_name)
    except requests.HTTPError as e:
        resp = e.response
        print(f"HTTP error: {resp.status_code}", file=sys.stderr)
        print(resp.text[:2000], file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Request error: {e}", file=sys.stderr)
        return 1

    items = data.get("items", []) or []
    best = pick_best_item(items, query_name)

    print("Ответ API:")
    print(f"  search_error: {data.get('search_error', '')!r}")
    print(f"  counter_title: {data.get('counter_title', '')!r}")
    print(f"  items_count: {len(items)}")
    print()

    if items:
        print("Первые совпадения:")
        for i, item in enumerate(items[:10], start=1):
            print(
                f"  {i}. id={item.get('id')} | ttl={item.get('ttl')} | "
                f"emitent_statuses_id={item.get('emitent_statuses_id')}"
            )
        print()
    else:
        print("Совпадений не найдено.")
        print()

    if best:
        print("Лучшее совпадение:")
        print(f"  id  = {best.get('id')}")
        print(f"  ttl = {best.get('ttl')}")
        print()
        print("Только id:")
        print(best.get("id"))
    else:
        print("ID не найден.")

    if args.raw:
        print()
        print("Полный JSON:")
        print(json.dumps(data, ensure_ascii=False, indent=2))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
