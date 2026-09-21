"""
Парсер рейтингов ЭМИТЕНТОВ от АКРА (Selenium).
Ищет по ИНН (столбец TIN) → первый результат с бейджом "Рейтингуемое лицо".
Парсит рейтинги эмитента из ПЕРВОГО div.rating-list.rating-list--js.

Столбцы:
  - acra_issuer_rating, _date, _forecast, _flag, _name

Вход/выход: bonds_retry.xlsx
Кэш: issuer_cache_acra.json (ключ = имя эмитента)

Требования:
  pip install selenium beautifulsoup4 pandas openpyxl tqdm
  pip install webdriver-manager   (опционально)
"""

import json
import os
import re
import time
import random

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm

# ============================================================
# НАСТРОЙКИ
# ============================================================

INPUT_FILE = "bonds_ext_ratings.xlsx"
OUTPUT_FILE = "bonds_ext_ratings.xlsx"
CACHE_FILE = "issuer_cache_acra.json"

TYPING_DELAY = (0.01, 0.05)
ACTION_DELAY = (0.5, 1.5)
PAGE_LOAD_DELAY = (1.0, 2.0)
BETWEEN_EMITTERS = (1.0, 2.0)
LONG_PAUSE_EVERY = 25
LONG_PAUSE = (5.0, 10.0)

MONTHS_RU = {
    'янв': '01', 'фев': '02', 'мар': '03', 'апр': '04',
    'май': '05', 'мая': '05', 'июн': '06', 'июл': '07',
    'авг': '08', 'сен': '09', 'окт': '10', 'ноя': '11', 'дек': '12',
}


# ============================================================
# КЭШ
# ============================================================

def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


# ============================================================
# УТИЛИТЫ
# ============================================================

def human_delay(rng):
    time.sleep(random.uniform(*rng))

def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(*TYPING_DELAY))

def random_mouse_move(driver, element):
    actions = ActionChains(driver)
    actions.move_to_element_with_offset(element, random.randint(-5, 5), random.randint(-5, 5))
    actions.perform()
    human_delay((0.3, 0.8))


# ============================================================
# БРАУЗЕР
# ============================================================

def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1440,900")
    options.add_argument("--lang=ru-RU")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    try:
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
    except ImportError:
        driver = webdriver.Chrome(options=options)

    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3]});
            Object.defineProperty(navigator, 'languages', {get: () => ['ru-RU','ru','en']});
            window.chrome = { runtime: {} };
        """
    })
    return driver


# ============================================================
# ПОИСК ПО ИНН НА АКРА
# ============================================================

def accept_cookies(driver):
    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH,
                "//button[contains(text(), 'Принять') or contains(text(), 'Accept') "
                "or contains(text(), 'Согласен')]"
            ))
        )
        random_mouse_move(driver, btn)
        btn.click()
        human_delay((1.0, 2.0))
    except TimeoutException:
        pass


def go_to_main_page(driver):
    driver.get("https://www.acra-ratings.ru/")
    human_delay(PAGE_LOAD_DELAY)
    accept_cookies(driver)


def search_by_inn(driver, inn: str):
    """
    Ищет по ИНН на АКРА. Возвращает (url, site_name) первого результата
    с бейджом "Рейтингуемое лицо".
    """
    # Прямой переход на страницу поиска — проще и надёжнее с ИНН
    driver.get(f"https://www.acra-ratings.ru/search/?q={inn}")
    human_delay(PAGE_LOAD_DELAY)

    return _find_issuer_link(driver)


def _find_issuer_link(driver):
    soup = BeautifulSoup(driver.page_source, "lxml")

    for item in soup.select("div.search-result__item"):
        tag_div = item.select_one("div.tag")
        if not tag_div:
            continue
        if "Рейтингуемое лицо" not in tag_div.get_text(strip=True):
            continue

        link = item.select_one("a.search-result__item-text")
        if not link:
            continue
        href = link.get("href", "")
        if "/ratings/issuers/" not in href:
            continue

        site_name = link.get_text(strip=True)
        full_url = ("https://www.acra-ratings.ru" + href) if href.startswith("/") else href
        return full_url, site_name

    # Fallback
    for a in soup.select("a.search-result__item-text"):
        href = a.get("href", "")
        if "/ratings/issuers/" in href:
            site_name = a.get_text(strip=True)
            full_url = ("https://www.acra-ratings.ru" + href) if href.startswith("/") else href
            return full_url, site_name

    return None, None


# ============================================================
# ПАРСИНГ СТРАНИЦЫ ЭМИТЕНТА
# ============================================================

def parse_issuer_ratings(driver):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
    human_delay((0.5, 1.0))

    soup = BeautifulSoup(driver.page_source, "lxml")
    history = []

    rating_list = soup.select_one("div.rating-list.rating-list--js")
    if not rating_list:
        return history

    for item in rating_list.select("div.rating-item"):
        rate_div = item.select_one("div.item-info[data-type='rate']")
        if not rate_div:
            continue

        full_text = rate_div.get_text(" ", strip=True)
        rating = full_text
        forecast = ""
        if "прогноз" in full_text.lower():
            parts = re.split(r'прогноз', full_text, flags=re.IGNORECASE)
            rating = parts[0].strip()
            forecast = parts[1].strip() if len(parts) > 1 else ""

        date_el = item.select_one("a[data-type='pressRelease'], a.item-info[data-type='pressRelease']")
        date_str = date_el.get_text(strip=True) if date_el else ""

        date_parsed = None
        if date_str:
            parts = date_str.split()
            if len(parts) >= 3:
                day = parts[0].zfill(2)
                mon_short = parts[1][:3].lower()
                year = parts[-1]
                month = MONTHS_RU.get(mon_short, "01")
                date_parsed = f"{year}-{month}-{day}"

        if rating and "отозван" not in rating.lower():
            history.append({
                "rating": rating,
                "forecast": forecast,
                "date": date_parsed,
                "date_raw": date_str,
            })

    return history


# ============================================================
# РЕЙТИНГ НА ДАТУ
# ============================================================

def get_rating_on_date(history, target_date):
    if not history or not target_date:
        return None, None, None, None

    dated = [h for h in history if h.get("date") and h.get("rating")]
    if not dated:
        return None, None, None, None

    dated.sort(key=lambda x: x["date"])

    best = None
    for r in dated:
        if r["date"] <= target_date:
            best = r

    if best:
        return best["rating"], best.get("forecast", ""), best["date"], "до размещения"

    earliest = dated[0]
    return earliest["rating"], earliest.get("forecast", ""), earliest["date"], "после размещения"


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("АКРА — рейтинги ЭМИТЕНТОВ (поиск по ИНН)")
    print("=" * 60)

    df = pd.read_excel(INPUT_FILE)
    print(f"Загружено: {len(df)} строк")

    pairs = df[["Эмитент", "TIN"]].dropna().drop_duplicates()
    pairs = pairs[pairs["Эмитент"].apply(lambda x: isinstance(x, str) and len(x.strip()) > 1)]
    pairs["TIN"] = pairs["TIN"].astype(str).str.strip().str.split('.').str[0].str.zfill(10)
    print(f"Уникальных эмитентов: {len(pairs)}")

    cache = load_cache()

    # Парсим заново ВСЕ
    if input("Запустить браузер? (y/n): ").strip().lower() != "y":
        print("Прервано. Применяю то, что есть в кэше.")
    else:
        print("Запуск Chrome (headless)...")
        driver = create_driver()

        try:
            go_to_main_page(driver)
            print("Сайт открыт\n")

            success = 0
            errors = 0
            not_found = 0

            for i, (_, row) in enumerate(tqdm(pairs.iterrows(), total=len(pairs), desc="АКРА (ИНН)")):
                emitter = row["Эмитент"]
                inn = row["TIN"]

                entry = {"history": [], "url": None, "site_name": None, "error": None, "inn": inn}

                try:
                    url, site_name = search_by_inn(driver, inn)

                    if url:
                        driver.get(url)
                        human_delay(PAGE_LOAD_DELAY)
                        accept_cookies(driver)

                        history = parse_issuer_ratings(driver)
                        entry["history"] = history
                        entry["url"] = url
                        entry["site_name"] = site_name

                        if history:
                            success += 1
                            tqdm.write(f"  {emitter} (ИНН {inn}) → {site_name} | рейтингов: {len(history)}")
                        else:
                            not_found += 1
                            tqdm.write(f"  {emitter} (ИНН {inn}) → {site_name} | рейтингов: 0")
                    else:
                        not_found += 1
                        entry["error"] = "Не найден по ИНН"
                        tqdm.write(f"  {emitter} (ИНН {inn}) → не найден")

                except Exception as e:
                    errors += 1
                    entry["error"] = str(e)
                    tqdm.write(f"  {emitter} (ИНН {inn}) → ОШИБКА: {e}")

                    if "net::ERR" in str(e) or "unreachable" in str(e).lower():
                        tqdm.write("Сеть недоступна. Останавливаюсь.")
                        cache[emitter] = entry
                        break

                cache[emitter] = entry
                human_delay(BETWEEN_EMITTERS)

                if (i + 1) % 10 == 0:
                    save_cache(cache)

                if (i + 1) % LONG_PAUSE_EVERY == 0:
                    pause = random.uniform(*LONG_PAUSE)
                    tqdm.write(f"  Пауза {pause:.0f}с (ok={success}, skip={not_found}, err={errors})")
                    time.sleep(pause)

            save_cache(cache)
            print(f"\nСпарсено: {success}, не найдено: {not_found}, ошибки: {errors}")

        finally:
            driver.quit()

    # --- Применяем к таблице ---
    print("\nПрименяю рейтинги к таблице...")

    col_rating = []
    col_date = []
    col_forecast = []
    col_flag = []
    col_name = []

    for _, row in df.iterrows():
        emitter = row.get("Эмитент")
        placement = row.get("Начало размещения")
        target = None
        if pd.notna(placement):
            target = pd.Timestamp(placement).strftime("%Y-%m-%d")

        rating, forecast, date, flag = None, None, None, None
        site_name = None

        if emitter and isinstance(emitter, str) and emitter in cache:
            entry = cache[emitter]
            hist = entry.get("history", [])
            site_name = entry.get("site_name")
            rating, forecast, date, flag = get_rating_on_date(hist, target)

        col_rating.append(rating)
        col_date.append(date)
        col_forecast.append(forecast)
        col_flag.append(flag)
        col_name.append(site_name)

    df["acra_issuer_rating"] = col_rating
    df["acra_issuer_rating_date"] = col_date
    df["acra_issuer_forecast"] = col_forecast
    df["acra_issuer_flag"] = col_flag
    df["acra_issuer_name"] = col_name

    found = sum(1 for r in col_rating if r is not None)
    after = sum(1 for f in col_flag if f == "после размещения")
    print(f"Найдено рейтингов: {found}/{len(df)} (из них {after} после размещения)")

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Сохранено: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
