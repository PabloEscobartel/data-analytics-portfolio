"""
Парсер рейтингов ЭМИТЕНТОВ от НКР (ratings.ru) через Selenium.
Ищет по названию из issuers.xlsx на странице ratings.ru/ratings/issuers/.
Переходит на страницу эмитента, парсит историю рейтингов из div.rate-items.

Добавляет столбцы в bonds_retry_nra.xlsx (по ключу TIN):
  - nkr_issuer_rating
  - nkr_issuer_rating_date
  - nkr_issuer_flag   ("до размещения" / "после размещения")

Кэш: issuer_cache_nkr.json (ключ = TIN)

Требования:
  pip install selenium beautifulsoup4 pandas openpyxl tqdm
  pip install webdriver-manager
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
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from tqdm import tqdm

# ============================================================
# НАСТРОЙКИ
# ============================================================

ISSUERS_FILE = "issuers.xlsx"
BONDS_FILE = "bonds_ext_ratings_cleaned.xlsx"
OUTPUT_FILE = "bonds_ext_ratings_cleaned.xlsx"
CACHE_FILE = "issuer_cache_nkr.json"

PAGE_LOAD_DELAY = (2.0, 3.5)
BETWEEN_ISSUERS = (1.5, 3.0)


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


def human_delay(rng):
    time.sleep(random.uniform(*rng))


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
        """
    })
    return driver


# ============================================================
# ПОИСК ЭМИТЕНТА ЧЕРЕЗ ПОЛЕ ПОИСКА
# ============================================================

def open_issuers_page(driver):
    """Открывает страницу списка эмитентов и ждёт загрузки таблицы."""
    driver.get("https://ratings.ru/ratings/issuers/")
    human_delay(PAGE_LOAD_DELAY)
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table#issuers-table tbody tr"))
        )
    except TimeoutException:
        pass


def search_issuer(driver, issuer_name: str):
    """
    Вбивает название в поле поиска на ratings.ru/ratings/issuers/.
    Берёт первую ссылку a.blue-link из отфильтрованной таблицы.
    Возвращает (url, site_name) или (None, None).
    """
    # Ищем поле ввода
    try:
        search_input = driver.find_element(By.CSS_SELECTOR,
            "input.filter-item__input[placeholder*='Название эмитента'], "
            "input.filter-item__input[name='filter-name']"
        )
    except NoSuchElementException:
        # Fallback: любой input в filter-item--search
        try:
            search_input = driver.find_element(By.CSS_SELECTOR,
                "div.filter-item--search input"
            )
        except NoSuchElementException:
            return None, None

    # Очищаем и вводим название
    search_input.clear()
    human_delay((0.3, 0.6))
    search_input.send_keys(issuer_name)
    human_delay((1.0, 2.0))  # ждём фильтрацию таблицы

    # Берём первую ссылку из отфильтрованной таблицы
    soup = BeautifulSoup(driver.page_source, "lxml")
    table = soup.select_one("table#issuers-table")
    if not table:
        return None, None

    for row in table.select("tbody tr"):
        # Пропускаем скрытые строки (display: none)
        style = row.get("style", "")
        if "display: none" in style or "display:none" in style:
            continue

        link = row.select_one("a.blue-link")
        if link:
            name = link.get_text(strip=True)
            href = link.get("href", "")
            if href:
                full_url = ("https://ratings.ru" + href) if href.startswith("/") else href
                # Очищаем поле перед следующим поиском
                search_input.clear()
                human_delay((0.3, 0.5))
                return full_url, name

    # Ничего не нашли — очищаем поле
    search_input.clear()
    human_delay((0.3, 0.5))
    return None, None


# ============================================================
# ПАРСИНГ СТРАНИЦЫ ЭМИТЕНТА
# ============================================================

def parse_issuer_ratings(driver) -> list:
    """
    Парсит историю рейтингов эмитента со страницы.
    Структура: div.rate-col (с title "Рейтинги эмитента") →
      div.rate-items → div.rate-item →
        p.rate-item__rating, p.rate-item__date
    """
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
    human_delay((0.5, 1.0))

    soup = BeautifulSoup(driver.page_source, "lxml")
    history = []

    # Ищем блок "Рейтинги эмитента"
    target_col = None
    for col in soup.select("div.rate-col"):
        title = col.select_one("p.rate-col__title")
        if title and "рейтинги эмитента" in title.get_text(strip=True).lower():
            target_col = col
            break

    if not target_col:
        # Fallback: берём первый rate-col с rate-items
        for col in soup.select("div.rate-col"):
            if col.select("div.rate-item"):
                target_col = col
                break

    if not target_col:
        return history

    for item in target_col.select("div.rate-item"):
        rating_p = item.select_one("p.rate-item__rating")
        date_p = item.select_one("p.rate-item__date")

        rating = rating_p.get_text(strip=True) if rating_p else ""
        date_raw = date_p.get_text(strip=True) if date_p else ""

        date_parsed = None
        match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', date_raw)
        if match:
            date_parsed = f"{match.group(3)}-{match.group(2)}-{match.group(1)}"

        if rating and "отозван" not in rating.lower():
            history.append({
                "rating": rating,
                "date": date_parsed,
                "date_raw": date_raw,
            })

    return history


# ============================================================
# РЕЙТИНГ НА ДАТУ
# ============================================================

def get_rating_on_date(history, target_date):
    if not history or not target_date:
        return None, None, None

    dated = [h for h in history if h.get("date") and h.get("rating")]
    if not dated:
        return None, None, None

    dated.sort(key=lambda x: x["date"])

    best = None
    for r in dated:
        if r["date"] <= target_date:
            best = r

    if best:
        return best["rating"], best["date"], "до размещения"

    earliest = dated[0]
    return earliest["rating"], earliest["date"], "после размещения"


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("НКР (ratings.ru) — рейтинги ЭМИТЕНТОВ")
    print("=" * 60)

    # Загружаем issuers.xlsx
    issuers_df = pd.read_excel(ISSUERS_FILE)
    issuers_df["TIN"] = issuers_df["TIN"].astype(str).str.strip().str.split('.').str[0].str.zfill(10)
    print(f"Эмитентов в issuers.xlsx: {len(issuers_df)}")

    cache = load_cache()

    # Определяем, кого парсить
    to_scrape = []
    for _, row in issuers_df.iterrows():
        tin = row["TIN"]
        if tin not in cache or cache[tin].get("history") is None:
            to_scrape.append(row)

    print(f"Уже в кэше: {len(issuers_df) - len(to_scrape)}")
    print(f"Нужно спарсить: {len(to_scrape)}")

    if to_scrape:
        if input("Запустить браузер? (y/n): ").strip().lower() != "y":
            print("Прервано. Применяю то, что есть в кэше.")
        else:
            print("Запуск Chrome (headless)...")
            driver = create_driver()

            try:
                # Открываем страницу списка эмитентов
                print("Открываю ratings.ru/ratings/issuers/...")
                open_issuers_page(driver)
                print("Страница загружена\n")

                success = 0
                not_found = 0
                errors = 0

                for row in tqdm(to_scrape, desc="НКР (эмитенты)"):
                    issuer_name = row["Issuer Name"]
                    tin = row["TIN"]

                    entry = {"history": [], "url": None, "site_name": None, "error": None}

                    try:
                        url, site_name = search_issuer(driver, issuer_name)

                        if url:
                            driver.get(url)
                            human_delay(PAGE_LOAD_DELAY)

                            history = parse_issuer_ratings(driver)
                            entry["history"] = history
                            entry["url"] = url
                            entry["site_name"] = site_name

                            # Возвращаемся на страницу списка для следующего поиска
                            open_issuers_page(driver)

                            if history:
                                success += 1
                                tqdm.write(f"  {issuer_name} (TIN {tin}) → {site_name} | рейтингов: {len(history)}")
                            else:
                                not_found += 1
                                tqdm.write(f"  {issuer_name} (TIN {tin}) → {site_name} | рейтингов: 0")
                        else:
                            not_found += 1
                            entry["error"] = "Не найден в списке эмитентов"
                            tqdm.write(f"  {issuer_name} (TIN {tin}) → не найден в списке")

                    except Exception as e:
                        errors += 1
                        entry["error"] = str(e)
                        tqdm.write(f"  {issuer_name} (TIN {tin}) → ОШИБКА: {e}")
                        # Возвращаемся на страницу списка
                        try:
                            open_issuers_page(driver)
                        except:
                            pass

                    cache[tin] = entry
                    human_delay(BETWEEN_ISSUERS)

                    if (success + not_found + errors) % 10 == 0:
                        save_cache(cache)

                save_cache(cache)
                print(f"\nСпарсено: {success}, не найдено: {not_found}, ошибки: {errors}")

            finally:
                driver.quit()

    # --- Применяем к bonds_retry_nra.xlsx ---
    print(f"\nПрименяю рейтинги к {BONDS_FILE}...")

    bonds = pd.read_excel(BONDS_FILE)
    bonds["_TIN_norm"] = bonds["TIN"].astype(str).str.strip().str.split('.').str[0].str.zfill(10)

    col_rating = []
    col_date = []
    col_flag = []

    for _, row in bonds.iterrows():
        tin = row["_TIN_norm"]
        placement = row.get("Начало размещения")
        target = None
        if pd.notna(placement):
            target = pd.Timestamp(placement).strftime("%Y-%m-%d")

        rating, date, flag = None, None, None

        if tin in cache:
            hist = cache[tin].get("history", [])
            rating, date, flag = get_rating_on_date(hist, target)

        col_rating.append(rating)
        col_date.append(date)
        col_flag.append(flag)

    bonds["nkr_issuer_rating"] = col_rating
    bonds["nkr_issuer_rating_date"] = col_date
    bonds["nkr_issuer_flag"] = col_flag

    bonds.drop(columns=["_TIN_norm"], inplace=True)

    found = sum(1 for r in col_rating if r is not None)
    after = sum(1 for f in col_flag if f == "после размещения")
    print(f"Найдено рейтингов: {found}/{len(bonds)} (из них {after} после размещения)")

    bonds.to_excel(OUTPUT_FILE, index=False)
    print(f"Сохранено: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
