"""
Парсер рейтингов АКРА через Selenium.
Для каждого ISIN из bonds_retry.xlsx ищет рейтинг, актуальный на дату размещения.
Добавляет два столбца: acra_rating, acra_rating_date.

Требования:
  pip install selenium beautifulsoup4 pandas openpyxl tqdm
  pip install webdriver-manager   (опционально, для автоскачивания chromedriver)
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    ElementClickInterceptedException, StaleElementReferenceException,
)
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
import json
import os
import re
from tqdm import tqdm

# ============================================================
# НАСТРОЙКИ
# ============================================================

INPUT_FILE = "bonds_ext_ratings.xlsx"
OUTPUT_FILE = "bonds_ext_ratings.xlsx"          # перезаписываем тот же файл
CACHE_FILE = "rating_cache_acra.json"

# Паузы (секунды)
TYPING_DELAY = (0.01, 0.05)
ACTION_DELAY = (0.5, 1.5)
PAGE_LOAD_DELAY = (1.0, 2.0)
BETWEEN_ISINS = (1.0, 2.0)
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

def load_cache() -> dict:
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_cache(cache: dict):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


# ============================================================
# УТИЛИТЫ ИМИТАЦИИ ЧЕЛОВЕКА
# ============================================================

def human_delay(range_tuple):
    time.sleep(random.uniform(*range_tuple))

def human_type(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(*TYPING_DELAY))

def human_scroll(driver, pixels=None):
    if pixels is None:
        pixels = random.randint(200, 500)
    driver.execute_script(f"window.scrollBy(0, {pixels})")
    human_delay((0.5, 1.5))

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
    # options.add_argument("--headless=new")  # раскомментировать для работы без окна
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
            Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3]});
            Object.defineProperty(navigator, 'languages', {get: () => ['ru-RU', 'ru', 'en']});
            window.chrome = {runtime: {}};
        """
    })
    return driver


# ============================================================
# ПАРСИНГ АКРА
# ============================================================

def accept_cookies(driver):
    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH,
                "//button[contains(text(), 'Принять') or contains(text(), 'Accept')]"
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

def search_isin_on_acra(driver, isin: str) -> str:
    """Ищет ISIN через поисковую строку АКРА. Возвращает URL или None."""
    try:
        search_buttons = driver.find_elements(By.CSS_SELECTOR,
            "button.search-btn, a.search-btn, .header-search, "
            "[data-action='search'], .search-icon, button[type='submit']"
        )
        search_input = None
        try:
            search_input = driver.find_element(By.CSS_SELECTOR,
                "input[type='search'], input[name='q'], input.search-form__input, "
                "input[placeholder*='Поиск'], input[placeholder*='поиск']"
            )
        except NoSuchElementException:
            pass

        if not search_input and search_buttons:
            random_mouse_move(driver, search_buttons[0])
            search_buttons[0].click()
            human_delay(ACTION_DELAY)
            search_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR,
                    "input[type='search'], input[name='q'], input.search-form__input"
                ))
            )

        if not search_input:
            driver.get(f"https://www.acra-ratings.ru/search/?q={isin}")
            human_delay(PAGE_LOAD_DELAY)
            return _find_rating_link(driver, isin)

        random_mouse_move(driver, search_input)
        search_input.click()
        human_delay((0.5, 1.0))
        search_input.clear()
        human_delay((0.3, 0.7))
        human_type(search_input, isin)
        human_delay((0.5, 1.5))
        search_input.send_keys(Keys.RETURN)
        human_delay(PAGE_LOAD_DELAY)

    except Exception:
        driver.get(f"https://www.acra-ratings.ru/search/?q={isin}")
        human_delay(PAGE_LOAD_DELAY)

    return _find_rating_link(driver, isin)


def _find_rating_link(driver, isin: str) -> str:
    soup = BeautifulSoup(driver.page_source, "lxml")

    for a in soup.select("a.search-result__item-text"):
        href = a.get("href", "")
        text = a.get_text(strip=True)
        if isin in text and "/ratings/" in href:
            return "https://www.acra-ratings.ru" + href

    for a in soup.select("a.search-result__item-text"):
        href = a.get("href", "")
        if "/ratings/emissions/" in href:
            return "https://www.acra-ratings.ru" + href

    return None


def click_to_rating_page(driver, url: str):
    path = url.replace("https://www.acra-ratings.ru", "")
    try:
        link = driver.find_element(By.CSS_SELECTOR, f"a[href='{path}']")
        human_scroll(driver, random.randint(50, 150))
        random_mouse_move(driver, link)
        human_delay((0.5, 1.0))
        link.click()
        human_delay(PAGE_LOAD_DELAY)
    except (NoSuchElementException, ElementClickInterceptedException):
        driver.get(url)
        human_delay(PAGE_LOAD_DELAY)


def parse_rating_history(driver) -> list:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
    human_delay((0.5, 1.0))

    soup = BeautifulSoup(driver.page_source, "lxml")
    history = []

    for item in soup.select("div.rating-item"):
        rate_div = item.select_one("div.item-info[data-type='rate']")
        rating = rate_div.get_text(strip=True) if rate_div else ""

        date_link = item.select_one("a.item-info[data-type='pressRelease']")
        if not date_link:
            date_link = item.select_one("a[data-type='pressRelease']")
        date_text = date_link.get_text(strip=True) if date_link else ""

        date_parsed = None
        if date_text:
            parts = date_text.split()
            if len(parts) == 3:
                day = parts[0].zfill(2)
                month = MONTHS_RU.get(parts[1][:3].lower(), "01")
                year = parts[2]
                date_parsed = f"{year}-{month}-{day}"

        if rating:
            history.append({
                "rating": rating,
                "date": date_parsed,
                "date_raw": date_text,
            })

    return history


def scrape_one_isin(driver, isin: str) -> list:
    rating_url = search_isin_on_acra(driver, isin)
    if not rating_url:
        return []
    click_to_rating_page(driver, rating_url)
    history = parse_rating_history(driver)
    driver.execute_script("window.scrollTo(0, 0)")
    human_delay((1.0, 2.0))
    return history


# ============================================================
# РЕЙТИНГ НА ДАТУ РАЗМЕЩЕНИЯ
# ============================================================

def get_rating_on_date(history: list, target_date: str):
    """
    Возвращает (rating, date) — рейтинг, актуальный на target_date.
    Берём самый свежий рейтинг, присвоенный не позже target_date.
    Если все рейтинги позже — берём самый ранний.
    """
    if not history or not target_date:
        return None, None

    dated = [h for h in history if h.get("date") and h.get("rating")]
    if not dated:
        return None, None

    dated.sort(key=lambda x: x["date"])

    best = None
    for r in dated:
        if r["date"] <= target_date:
            best = r

    if best:
        return best["rating"], best["date"]

    # Все позже — берём самый ранний
    return dated[0]["rating"], dated[0]["date"]


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print("Парсер АКРА (Selenium)")
    print("=" * 60)

    df = pd.read_excel(INPUT_FILE)
    cache = load_cache()
    print(f"Загружено: {len(df)} строк")

    # Определяем ISIN для парсинга: нет в кэше или кэш пустой/с ошибкой
    isins_all = df["ISIN"].dropna().unique().tolist()
    isins_all = [i for i in isins_all if isinstance(i, str) and len(i.strip()) > 10]

    to_scrape = []
    for isin in isins_all:
        if isin not in cache:
            to_scrape.append(isin)
        elif not cache[isin].get("history") and cache[isin].get("error"):
            to_scrape.append(isin)  # retry ошибок

    to_scrape = list(set(to_scrape))
    print(f"Всего ISIN: {len(isins_all)}")
    print(f"Уже в кэше: {len(isins_all) - len(to_scrape)}")
    print(f"Нужно спарсить: {len(to_scrape)}")

    if to_scrape:
        if input("Запустить браузер? (y/n): ").strip().lower() != "y":
            print("Прервано.")
            # Всё равно применяем то, что есть в кэше
        else:
            print("Запуск Chrome...")
            driver = create_driver()

            try:
                go_to_main_page(driver)
                print("Сайт открыт")

                success = 0
                errors = 0
                not_found = 0

                for i, isin in enumerate(tqdm(to_scrape, desc="АКРА")):
                    try:
                        history = scrape_one_isin(driver, isin)

                        cache[isin] = {"history": history, "error": None}

                        if history:
                            success += 1
                            if success <= 5:
                                print(f"\n  {isin}: {history[0]['rating']} ({history[0].get('date', '?')})")
                        else:
                            not_found += 1

                    except Exception as e:
                        errors += 1
                        cache[isin] = {"history": [], "error": str(e)}
                        if "net::ERR" in str(e) or "unreachable" in str(e).lower():
                            print(f"\nСеть недоступна. Останавливаюсь.")
                            break

                    human_delay(BETWEEN_ISINS)

                    if (i + 1) % LONG_PAUSE_EVERY == 0:
                        save_cache(cache)
                        pause = random.uniform(*LONG_PAUSE)
                        print(f"\n  Пауза {pause:.0f} сек... (ok={success}, skip={not_found}, err={errors})")
                        time.sleep(pause)

                save_cache(cache)
                print(f"\nСпарсено: {success}, не найдено: {not_found}, ошибки: {errors}")

            finally:
                driver.quit()

    # Применяем рейтинги к таблице
    print("\nПрименяем рейтинги к таблице...")
    ratings_col = []
    dates_col = []

    for _, row in df.iterrows():
        isin = row.get("ISIN")
        placement = row.get("Начало размещения")
        target = None
        if pd.notna(placement):
            target = pd.Timestamp(placement).strftime("%Y-%m-%d")

        rating, date = None, None
        if isin and isinstance(isin, str) and isin in cache:
            history = cache[isin].get("history", [])
            rating, date = get_rating_on_date(history, target)

        ratings_col.append(rating)
        dates_col.append(date)

    df["acra_rating"] = ratings_col
    df["acra_rating_date"] = dates_col

    found = sum(1 for r in ratings_col if r is not None)
    print(f"Найдено рейтингов: {found}/{len(df)}")

    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Сохранено: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
