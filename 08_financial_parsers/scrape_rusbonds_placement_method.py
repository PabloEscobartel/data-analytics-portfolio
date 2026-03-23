"""
scrape_rusbonds_placement_method.py
Парсер способа размещения с RusBonds по ISIN
"""

import os
import time
import random
import json
from typing import Optional

import pandas as pd
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# ====================== НАСТРОЙКИ ======================

INPUT_FILE = "panel_data.xlsx"
OUTPUT_FILE = "panel_with_placement_method.xlsx"
CACHE_FILE = "rusbonds_placement_method_cache.json"

# ====================== API ======================

API_URL = "https://rusbonds.ru/api/v2/search/hint"

BEARER_TOKEN = os.getenv("RUSBONDS_BEARER_TOKEN", "")

COOKIES = {
    # Заполните после авторизации на rusbonds.ru
    # См. README для инструкций
}

HEADERS = {
    "accept": "application/json, text/plain, */*",
    "authorization": f"Bearer {BEARER_TOKEN}",
    "content-type": "application/json",
    "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"
}

# ====================== КЭШ ======================

def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

cache = load_cache()

# ====================== API ПОИСК ID ======================

def get_bond_id_via_api(isin: str) -> Optional[int]:
    session = requests.Session()
    session.headers.update(HEADERS)
    session.cookies.update(COOKIES)

    payload = {
        "query": isin,
        "page": 1,
        "size": 50,
        "types": []
    }

    try:
        r = session.post(API_URL, json=payload, timeout=30)
        if r.status_code == 200:
            data = r.json()
            if data.get("entities"):
                return data["entities"][0]["id"]
        print(f"  API: {isin} не найден")
    except Exception as e:
        print(f"  API ошибка {isin}: {e}")

    return None

# ====================== PLAYWRIGHT ======================

def get_placement_method(page, bond_id: int, isin: str):
    try:
        page.goto(
            f"https://rusbonds.ru/bonds/{bond_id}/placement",
            wait_until="domcontentloaded",
            timeout=40000
        )

        if "xpvnsulc" in page.url:
            print(f"  ⚠ {isin}: проверка сайта, повтор...")
            time.sleep(2)
            page.goto(
                f"https://rusbonds.ru/bonds/{bond_id}/placement",
                wait_until="domcontentloaded",
                timeout=40000
            )

        locator = page.locator(
            'div.data-item:has(div.data-name:has-text("Способ размещения")) div.data-value'
        ).first

        value_text = locator.text_content(timeout=20000)

        if not value_text:
            print(f"  ❌ {isin}: способ размещения не найден")
            return None

        value_text = value_text.strip()
        print(f"  ✅ {isin}: {value_text}")
        return value_text

    except PlaywrightTimeoutError:
        print(f"  ❌ {isin}: таймаут")
        return None
    except Exception as e:
        print(f"  ❌ {isin}: ошибка {e}")
        return None

# ====================== КОМБИНИРОВАННАЯ ФУНКЦИЯ ======================

def get_method(isin: str, page):
    if isin in cache:
        print(f"  {isin} → кэш: {cache[isin]}")
        return cache[isin]

    bond_id = get_bond_id_via_api(isin)
    if not bond_id:
        cache[isin] = None
        save_cache(cache)
        return None

    value = get_placement_method(page, bond_id, isin)
    cache[isin] = value
    save_cache(cache)
    return value

# ====================== MAIN ======================

def main():
    print("=" * 70)
    print("RUSBONDS — СПОСОБ РАЗМЕЩЕНИЯ")
    print("=" * 70)

    try:
        df = pd.read_excel(INPUT_FILE)
    except:
        print(f"Файл {INPUT_FILE} не найден")
        return

    if "ISIN" not in df.columns:
        print("Нет колонки ISIN")
        return

    if "placement_method" not in df.columns:
        df["placement_method"] = ""

    isins = df["ISIN"].dropna().tolist()
    indices = df[df["ISIN"].notna()].index.tolist()

    # Пропускаем уже собранные
    to_parse = [(isin, idx) for isin, idx in zip(isins, indices)
                if isin not in cache]

    print(f"Всего ISIN: {len(isins)}")
    print(f"В кэше: {len(isins) - len(to_parse)}")
    print(f"К парсингу: {len(to_parse)}")

    if len(to_parse) > 0:
        if input("Запустить? y/n: ").lower() != "y":
            return

    # Заполняем из кэша
    for isin, idx in zip(isins, indices):
        if isin in cache and cache[isin]:
            df.loc[idx, "placement_method"] = cache[isin]

    if len(to_parse) == 0:
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\nВсё в кэше. Сохранено: {OUTPUT_FILE}")
        return

    # Парсинг
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=["--disable-blink-features=AutomationControlled"]
        )

        context = browser.new_context(
            viewport={"width": 1200, "height": 800},
            locale="ru-RU",
            user_agent="Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"
        )

        context.add_cookies([
            {"name": k, "value": v, "domain": "rusbonds.ru", "path": "/"}
            for k, v in COOKIES.items()
        ])

        page = context.new_page()
        updated = 0

        try:
            for i, (isin, idx) in enumerate(to_parse, 1):
                print(f"\n{i}/{len(to_parse)}  {isin}")

                value = get_method(isin, page)

                if value is not None:
                    df.loc[idx, "placement_method"] = value
                    updated += 1

                if i % 5 == 0:
                    df.to_excel(OUTPUT_FILE, index=False)
                    print("💾 промежуточное сохранение")

                if i < len(to_parse):
                    time.sleep(random.uniform(0.2, 0.5))

        finally:
            browser.close()

    df.to_excel(OUTPUT_FILE, index=False)

    # Статистика
    print(f"\n{'=' * 70}")
    print("РЕЗУЛЬТАТЫ")
    print(f"{'=' * 70}")
    print(f"Обновлено: {updated}")
    print(f"Файл: {OUTPUT_FILE}")

    methods = df["placement_method"].value_counts()
    print(f"\nСпособы размещения:")
    for m, cnt in methods.items():
        if m:
            print(f"  {m}: {cnt}")

if __name__ == "__main__":
    main()
