"""
Скрипт для сбора данных о размещении с cbonds.ru

Что делает:
1. Читает ISIN из panel_data.xlsx
2. Для каждого ISIN получает cbonds ID через suggest API
3. Загружает страницу бумаги и парсит: Способ размещения, Тип размещения, Формат размещения
4. Сохраняет результат в panel_data_enriched.xlsx

Промежуточные результаты кэшируются в cbonds_cache.json — если прервать и запустить заново,
уже обработанные ISIN будут пропущены.

Установка зависимостей:
    pip install requests openpyxl beautifulsoup4 lxml
"""

import requests
import time
import json
import os
from bs4 import BeautifulSoup
import openpyxl

# ─── Настройки ────────────────────────────────────────────────────────────────

INPUT_FILE = "../bonds_offer_before2018_all_orientirs_gemini.xlsx"
OUTPUT_FILE = "../bonds_offer_before2018_all_orientirs_gemini_cbonds.xlsx"
CACHE_FILE = "cbonds_cache.json"

DELAY = 0.7          # секунд между запросами (не слишком агрессивно)
SUGGEST_URL = "https://cbonds.ru/api/suggest/bonds/"
BOND_PAGE_URL = "https://cbonds.ru/bonds/{bond_id}/"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/146.0.0.0 Safari/537.36"
    ),
}

# HTML id элементов, из которых берём значения
FIELD_IDS = {
    "placement_method":  "cb_bond_page_ginfo_resume_private_offering",       # Способ размещения
    "placement_type":    "cb_bond_page_ginfo_resume_placing_type_name",      # Тип размещения
    "placement_format":  "cb_bond_page_placement_auction_type_name",         # Формат размещения
}


# ─── Кэш ─────────────────────────────────────────────────────────────────────

def load_cache() -> dict:
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache: dict):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


# ─── API ──────────────────────────────────────────────────────────────────────

session = requests.Session()
session.headers.update(HEADERS)


def get_cbonds_id(isin: str) -> int | None:
    """ISIN → cbonds ID через suggest API."""
    try:
        resp = session.post(
            SUGGEST_URL,
            data=isin,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=15,
        )
        resp.raise_for_status()
        items = resp.json().get("items", [])
        return items[0]["id"] if items else None
    except Exception as e:
        print(f"    [suggest error] {isin}: {e}")
        return None


def parse_bond_page(bond_id: int) -> dict:
    """Загружает страницу бумаги и парсит три поля о размещении."""
    result = {k: None for k in FIELD_IDS}
    try:
        resp = session.get(BOND_PAGE_URL.format(bond_id=bond_id), timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")

        for key, html_id in FIELD_IDS.items():
            li = soup.find("li", id=html_id)
            if li:
                value_div = li.find("div", class_="value")
                if value_div:
                    span = value_div.find("span")
                    result[key] = span.get_text(strip=True) if span else value_div.get_text(strip=True)
    except Exception as e:
        print(f"    [page error] bond_id={bond_id}: {e}")

    return result


def process_isin(isin: str) -> dict:
    """Полный пайплайн: ISIN → cbonds ID → парсинг страницы."""
    bond_id = get_cbonds_id(isin)
    if bond_id is None:
        return {"cbonds_id": None, "placement_method": None, "placement_type": None, "placement_format": None}

    time.sleep(DELAY)
    fields = parse_bond_page(bond_id)
    fields["cbonds_id"] = bond_id
    return fields


# ─── Основной скрипт ─────────────────────────────────────────────────────────

def main():
    # Читаем Excel
    wb = openpyxl.load_workbook(INPUT_FILE)
    ws = wb["bonds_debut"]

    # Собираем уникальные ISIN
    isin_col = 2  # столбец B
    unique_isins = set()
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, isin_col).value
        if v:
            unique_isins.add(str(v).strip())

    print(f"Уникальных ISIN: {len(unique_isins)}")

    # Загружаем кэш
    cache = load_cache()
    already = sum(1 for isin in unique_isins if isin in cache)
    print(f"Уже в кэше: {already}")
    print(f"Осталось: {len(unique_isins) - already}")
    print()

    # Обрабатываем
    todo = [isin for isin in sorted(unique_isins) if isin not in cache]
    for i, isin in enumerate(todo):
        print(f"[{already + i + 1}/{len(unique_isins)}] {isin}", end=" → ")
        data = process_isin(isin)
        cache[isin] = data

        bid = data.get("cbonds_id", "?")
        pm = data.get("placement_method", "-")
        pt = data.get("placement_type", "-")
        pf = data.get("placement_format", "-")
        print(f"id={bid}  |  {pm}  |  {pt}  |  {pf}")

        # Сохраняем кэш каждые 20 записей
        if (i + 1) % 20 == 0:
            save_cache(cache)
            print(f"    [кэш сохранён: {len(cache)} записей]")

        if i < len(todo) - 1:
            time.sleep(DELAY)

    save_cache(cache)
    print(f"\nКэш финализирован: {len(cache)} записей")

    # Записываем результат в Excel
    new_cols = {
        ws.max_column + 1: "placement_method",
        ws.max_column + 2: "placement_type",
        ws.max_column + 3: "placement_format",
    }

    for col, name in new_cols.items():
        ws.cell(1, col, name)

    for r in range(2, ws.max_row + 1):
        isin = str(ws.cell(r, isin_col).value or "").strip()
        data = cache.get(isin, {})
        for col, key in new_cols.items():
            ws.cell(r, col, data.get(key))

    wb.save(OUTPUT_FILE)
    print(f"Готово: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
