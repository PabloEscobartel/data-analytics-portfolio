#!/usr/bin/env python3
import argparse
import html
import json
import re
import shutil
import time
import urllib.error
import urllib.parse
import urllib.request
from html.parser import HTMLParser
from pathlib import Path

import openpyxl

try:
    from bs4 import BeautifulSoup
except ImportError:
    BeautifulSoup = None

try:
    from selenium import webdriver
    from selenium.common.exceptions import TimeoutException, WebDriverException
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.support.ui import WebDriverWait
except ImportError:
    webdriver = None
    TimeoutException = None
    WebDriverException = None
    ChromeOptions = None
    WebDriverWait = None


SEARCH_URL = "https://bonds.finam.ru/issue/search/default.asp?emitterCustomName={isin}"
BASE_URL = "https://bonds.finam.ru"
CACHE_PATH = Path("finam_organizers_cache.json")
REFETCH_STATUSES = {
    "servicepipe_antibot",
    "issue_link_not_found",
    "organizers_not_found",
}


class TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []

    def handle_data(self, data):
        text = normalize_space(data)
        if text:
            self.parts.append(text)


def normalize_space(value):
    if value is None:
        return ""
    text = html.unescape(str(value)).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def decode_response(data, content_type=""):
    match = re.search(r"charset=([\w-]+)", content_type or "", flags=re.I)
    encodings = []
    if match:
        encodings.append(match.group(1))
    encodings.extend(["windows-1251", "cp1251", "utf-8"])
    for encoding in encodings:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="replace")


def fetch(url, timeout, retries, pause):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
        "Connection": "close",
    }
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=timeout) as response:
                data = response.read()
                return decode_response(data, response.headers.get("Content-Type", ""))
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(pause * attempt)
    raise RuntimeError(f"cannot fetch {url}: {last_error}")


class SeleniumFetcher:
    def __init__(self, timeout, pause, headless=False):
        if webdriver is None:
            raise RuntimeError("Selenium не установлен. Запустите через .venv/bin/python или установите selenium.")

        self.headless = headless
        self.timeout = timeout
        self.pause = pause
        self.current_url = ""
        self.driver = self._new_driver()

    def _new_driver(self):
        options = ChromeOptions()
        options.page_load_strategy = "eager"
        if self.headless:
            options.add_argument("--headless=new")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--blink-settings=imagesEnabled=false")
        options.add_argument("--window-size=1400,1000")
        options.add_argument(
            "--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
        )
        driver = webdriver.Chrome(options=options)
        driver.set_page_load_timeout(self.timeout)
        return driver

    def fetch(self, url):
        try:
            self.driver.get(url)
        except TimeoutException:
            try:
                self.driver.execute_script("window.stop();")
            except WebDriverException:
                pass
        except WebDriverException as exc:
            raise RuntimeError(compact_selenium_error(exc))

        self._wait_after_navigation()
        self.current_url = self.driver.current_url
        return self.driver.page_source

    def close(self):
        try:
            self.driver.quit()
        except Exception:
            pass

    def restart(self):
        self.close()
        self.driver = self._new_driver()
        self.current_url = ""

    def _wait_after_navigation(self):
        deadline = time.time() + self.timeout
        while time.time() < deadline:
            source = self.driver.page_source
            if not is_servicepipe_antibot(source):
                return
            time.sleep(max(self.pause, 1.0))


def find_issue_details_url(search_html):
    if BeautifulSoup is not None:
        soup = BeautifulSoup(search_html, "lxml")
        candidates = [
            a.get("href", "")
            for a in soup.find_all("a", href=True)
            if re.search(r"/issue/details[0-9A-Za-z]+/default\.asp", a.get("href", ""), flags=re.I)
        ]
    else:
        candidates = []

    if not candidates:
        candidates = re.findall(
            r"""href\s*=\s*["']([^"']*/issue/details[0-9A-Za-z]+/default\.asp)["']""",
            search_html,
            flags=re.I,
        )

    if not candidates:
        candidates = re.findall(
            r"""location\.href\s*=\s*["']([^"']*/issue/details[0-9A-Za-z]+/default\.asp)["']""",
            search_html,
            flags=re.I,
        )

    if not candidates:
        return ""

    return make_operators_url(candidates[0])


def make_operators_url(path):
    if path.startswith("http"):
        parsed = urllib.parse.urlparse(path)
        path = parsed.path

    path = re.sub(r"(/issue/details[0-9A-Za-z]+)(/default\.asp)$", r"\g<1>00004\2", path, flags=re.I)
    path = re.sub(r"(00004)+(/default\.asp)$", r"00004\2", path, flags=re.I)
    return urllib.parse.urljoin(BASE_URL, path)


def make_operators_url_from_current_page(current_url, page_html, isin):
    if not current_url:
        return ""
    parsed = urllib.parse.urlparse(current_url)
    if not re.search(r"/issue/details[0-9A-Za-z]+/default\.asp", parsed.path, flags=re.I):
        return ""
    if isin and isin not in page_html:
        return ""
    return make_operators_url(parsed.path)


def find_issue_details_url_regex_only(search_html):
    candidates = re.findall(
        r"""href\s*=\s*["']([^"']*/issue/details[0-9A-Za-z]+/default\.asp)["']""",
        search_html,
        flags=re.I,
    )
    if not candidates:
        candidates = re.findall(
            r"""location\.href\s*=\s*["']([^"']*/issue/details[0-9A-Za-z]+/default\.asp)["']""",
            search_html,
            flags=re.I,
        )
    if not candidates:
        return ""

    return make_operators_url(candidates[0])


def is_servicepipe_antibot(page_html):
    markers = (
        "servicepipe.ru/static/checkjs",
        "get_cookie_spsn",
        "get_cookie_spid",
        "id_captcha_frame_div",
    )
    return any(marker in page_html for marker in markers)


def compact_selenium_error(exc):
    text = str(exc).splitlines()[0].strip()
    if "net::ERR_CONNECTION_TIMED_OUT" in str(exc):
        return "selenium_connection_timed_out"
    if "Timed out receiving message from renderer" in str(exc):
        return "selenium_renderer_timed_out"
    if "timeout" in text.lower():
        return "selenium_timeout"
    return text[:180] or exc.__class__.__name__


def extract_organizers(details_html):
    if BeautifulSoup is not None:
        soup = BeautifulSoup(details_html, "lxml")
        for row in soup.find_all("tr"):
            row_cells = row.find_all("td", recursive=False)
            cells = [normalize_space(cell.get_text(" ", strip=True)).strip('"') for cell in row_cells]
            if not cells:
                continue
            if cells[0] == "Организатор":
                values = []
                for cell in row_cells[1:]:
                    for item in cell.stripped_strings:
                        value = normalize_space(item).strip('"')
                        if value and value not in {"Организатор", "|"}:
                            values.append(value)
                if values:
                    return "; ".join(dict.fromkeys(values))

    parser = TextExtractor()
    parser.feed(details_html)
    parts = [normalize_space(p).strip('"') for p in parser.parts if normalize_space(p)]

    for index, part in enumerate(parts):
        if part == "Организатор":
            values = []
            for candidate in parts[index + 1 :]:
                if candidate in {
                    "Агент по размещению",
                    "Андеррайтер",
                    "Платежный агент",
                    "Расчетный депозитарий",
                    "Регистратор",
                }:
                    break
                if candidate and candidate not in values:
                    values.append(candidate)
            if values:
                return "; ".join(values)

    match = re.search(
        r"Организатор\s*</b>.*?<td[^>]*>\s*(?:&quot;|\"|')?([^<\"']+)",
        details_html,
        flags=re.I | re.S,
    )
    if match:
        return normalize_space(match.group(1))
    return ""


def load_cache():
    if CACHE_PATH.exists():
        return json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    return {}


def save_cache(cache):
    CACHE_PATH.write_text(
        json.dumps(cache, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )


def find_header(headers, wanted):
    wanted_norm = wanted.casefold()
    for index, value in enumerate(headers, start=1):
        if normalize_space(value).casefold() == wanted_norm:
            return index
    return None


def should_fetch(cache_entry):
    if not cache_entry:
        return True
    if cache_entry.get("organizers"):
        return False
    status = cache_entry.get("status", "")
    return status.startswith("error:") or status in REFETCH_STATUSES


def is_retryable_error_status(status):
    retryable = (
        "selenium_connection_timed_out",
        "selenium_renderer_timed_out",
        "selenium_timeout",
        "net::ERR_CONNECTION_TIMED_OUT",
        "Timed out receiving message from renderer",
    )
    return any(marker in status for marker in retryable)


def save_debug_html(args, isin, name, page_html):
    if not args.debug_dir:
        return
    debug_dir = Path(args.debug_dir)
    debug_dir.mkdir(parents=True, exist_ok=True)
    safe_isin = re.sub(r"[^A-Za-z0-9_-]+", "_", isin)
    (debug_dir / f"{safe_isin}_{name}.html").write_text(page_html, encoding="utf-8")


def resolve_isin(isin, args, selenium_fetcher=None):
    search_url = SEARCH_URL.format(isin=urllib.parse.quote(isin))
    if args.backend == "selenium" and selenium_fetcher is not None:
        search_html = selenium_fetcher.fetch(search_url)
    else:
        search_html = fetch(search_url, args.timeout, args.retries, args.pause)

    if is_servicepipe_antibot(search_html):
        if args.backend == "requests":
            save_debug_html(args, isin, "search_antibot", search_html)
            return {"organizers": "", "status": "servicepipe_antibot"}
        if selenium_fetcher is None:
            save_debug_html(args, isin, "search_antibot", search_html)
            return {"organizers": "", "status": "servicepipe_antibot_selenium_unavailable"}
        search_html = selenium_fetcher.fetch(search_url)
        if is_servicepipe_antibot(search_html):
            save_debug_html(args, isin, "search_antibot_after_selenium", search_html)
            return {"organizers": "", "status": "servicepipe_antibot_after_selenium"}

    save_debug_html(args, isin, "search", search_html)
    details_url = find_issue_details_url(search_html)
    if not details_url and selenium_fetcher is not None:
        details_url = make_operators_url_from_current_page(
            selenium_fetcher.current_url,
            search_html,
            isin,
        )
    if not details_url:
        return {"organizers": "", "status": "issue_link_not_found"}

    if selenium_fetcher is not None:
        details_html = selenium_fetcher.fetch(details_url)
    else:
        details_html = fetch(details_url, args.timeout, args.retries, args.pause)

    if is_servicepipe_antibot(details_html):
        if args.backend == "requests" or selenium_fetcher is None:
            save_debug_html(args, isin, "details_antibot", details_html)
            return {"organizers": "", "status": "servicepipe_antibot_details", "details_url": details_url}
        details_html = selenium_fetcher.fetch(details_url)
        if is_servicepipe_antibot(details_html):
            save_debug_html(args, isin, "details_antibot_after_selenium", details_html)
            return {"organizers": "", "status": "servicepipe_antibot_details_after_selenium", "details_url": details_url}

    save_debug_html(args, isin, "details", details_html)
    organizers = extract_organizers(details_html)
    return {
        "organizers": organizers,
        "status": "ok" if organizers else "organizers_not_found",
        "details_url": details_url,
    }


def resolve_isin_with_retries(isin, args, selenium_fetcher=None):
    attempts = max(1, args.selenium_attempts if selenium_fetcher is not None else 1)
    last_result = None
    for attempt in range(1, attempts + 1):
        try:
            return resolve_isin(isin, args, selenium_fetcher)
        except Exception as exc:
            status = f"error: {compact_selenium_error(exc)}"
            last_result = {"organizers": "", "status": status}
            if selenium_fetcher is None or attempt >= attempts or not is_retryable_error_status(status):
                return last_result
            selenium_fetcher.restart()
            time.sleep(max(args.pause, 1.0) * attempt)
    return last_result or {"organizers": "", "status": "error: unknown"}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "xlsx",
        nargs="?",
        default="bonds_offer_updated_cbonds_orientir_gemini_retry.xlsx",
    )
    parser.add_argument("--sheet", default=None)
    parser.add_argument("--timeout", type=int, default=35)
    parser.add_argument("--retries", type=int, default=3)
    parser.add_argument("--pause", type=float, default=1.0)
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument(
        "--backend",
        choices=("auto", "requests", "selenium"),
        default="auto",
        help="requests = только urllib; selenium = все страницы через браузер; auto = urllib + Selenium при Servicepipe",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Запускать Chrome в headless-режиме. Для Servicepipe надежнее обычное окно браузера.",
    )
    parser.add_argument(
        "--only-isin",
        default="",
        help="Отладить один ISIN без изменения Excel.",
    )
    parser.add_argument(
        "--debug-dir",
        default="",
        help="Папка для сохранения HTML страниц поиска/details.",
    )
    parser.add_argument(
        "--selenium-attempts",
        type=int,
        default=3,
        help="Сколько раз повторять ISIN при Selenium timeout/network error.",
    )
    args = parser.parse_args()

    if args.only_isin:
        selenium_fetcher = None
        try:
            if args.backend in {"auto", "selenium"}:
                selenium_fetcher = SeleniumFetcher(args.timeout, args.pause, args.headless)
            result = resolve_isin_with_retries(normalize_space(args.only_isin), args, selenium_fetcher)
            print(json.dumps(result, ensure_ascii=False, indent=2))
        finally:
            if selenium_fetcher is not None:
                selenium_fetcher.close()
        return

    path = Path(args.xlsx)
    if not path.exists():
        raise FileNotFoundError(path)

    workbook = openpyxl.load_workbook(path)
    worksheet = workbook[args.sheet] if args.sheet else workbook.active
    headers = [worksheet.cell(1, col).value for col in range(1, worksheet.max_column + 1)]
    isin_col = find_header(headers, "ISIN")
    if not isin_col:
        raise RuntimeError("Не найден столбец ISIN")

    target_col = find_header(headers, "организаторы_финам")
    if not target_col:
        target_col = worksheet.max_column + 1
        worksheet.cell(1, target_col).value = "организаторы_финам"

    cache = load_cache()
    processed = 0
    updated = 0
    errors = 0
    selenium_fetcher = None

    if args.backend == "selenium":
        selenium_fetcher = SeleniumFetcher(args.timeout, args.pause, args.headless)

    try:
        for row in range(2, worksheet.max_row + 1):
            isin = normalize_space(worksheet.cell(row, isin_col).value)
            if not isin:
                continue
            existing = normalize_space(worksheet.cell(row, target_col).value)
            if existing:
                continue

            if should_fetch(cache.get(isin)):
                if args.backend == "auto" and selenium_fetcher is None:
                    selenium_fetcher = SeleniumFetcher(args.timeout, args.pause, args.headless)
                cache[isin] = resolve_isin_with_retries(isin, args, selenium_fetcher)
                if cache[isin].get("status", "").startswith("error:"):
                    errors += 1
                save_cache(cache)
                time.sleep(args.pause)

            organizers = cache.get(isin, {}).get("organizers", "")
            if organizers:
                worksheet.cell(row, target_col).value = organizers
                updated += 1

            processed += 1
            status = cache.get(isin, {}).get("status", "")
            print_status = status if len(status) <= 120 else status[:117] + "..."
            print(f"{processed}: {isin} status={print_status} organizers={organizers}", flush=True)
            if processed % 25 == 0:
                print(f"processed={processed} updated={updated} errors={errors}", flush=True)
            if args.limit and processed >= args.limit:
                break
    finally:
        if selenium_fetcher is not None:
            selenium_fetcher.close()

    backup = path.with_suffix(path.suffix + ".bak")
    if not backup.exists():
        shutil.copy2(path, backup)
    workbook.save(path)
    print(f"done sheet={worksheet.title} processed={processed} updated={updated} errors={errors}")
    print(f"backup={backup}")
    print(f"cache={CACHE_PATH}")


if __name__ == "__main__":
    main()
