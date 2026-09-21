#!/usr/bin/env python3
"""
Дозагрузка связанной эмиссии со страниц новостей Cbonds.

Берёт новости из articles, где:
    has_orientir_term = 1
    ИЛИ
    has_coupon_spread_term = 1

Для каждой новости открывает её url и ищет на странице блок вида:
    Эмиссия — <a href="https://cbonds.ru/bonds/1443472/">...</a>

Результат пишет в новую таблицу article_emissions, а в articles добавляет
технические поля статуса проверки.

Cookie берётся так же, как в cbonds_news_scrape.py:
- из переменной окружения CBONDS_COOKIE;
- или из файла cbonds_cookie.txt рядом со скриптом.

Запуск:
    python scrape_cbonds_news_emissions.py

Тестовый запуск на 10 новостях:
    python scrape_cbonds_news_emissions.py --limit 10

Перепроверить уже обработанные новости:
    python scrape_cbonds_news_emissions.py --refresh

Проверить все новости, привязанные к выпускам в issue_articles, без флагов:
    python scrape_cbonds_news_emissions.py --all-linked

Пропустить статьи, привязанные только к ВЭБ.РФ:
    python scrape_cbonds_news_emissions.py --all-linked --skip-veb-rf

Запуск без Cookie:
    python scrape_cbonds_news_emissions.py --no-cookie
"""

from __future__ import annotations

import argparse
import re
import sqlite3
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

from bs4 import BeautifulSoup

import cbonds_news_scrape as base


DB_FILE = "cbonds_news.sqlite3"
DELAY = 0.7
DEFAULT_DATE_TO = "2018-02-28"
BOND_HREF_RE = re.compile(r"/bonds/(\d+)/?", re.I)
EMISSION_LABEL_RE = re.compile(r"(?iu)\b(?:эмисси[яи]|emission|issue)\b")
SPACE_RE = re.compile(r"\s+")


@dataclass(frozen=True)
class EmissionLink:
    emission_cbonds_id: int | None
    emission_name: str
    emission_url: str
    label: str | None
    raw_context: str | None


def normalize_space(value: str | None) -> str:
    if not value:
        return ""
    return SPACE_RE.sub(" ", value.replace("\xa0", " ")).strip()


def existing_columns(conn: sqlite3.Connection, table_name: str) -> set[str]:
    rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    return {str(row[1]) for row in rows}


def ensure_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS article_emissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            article_row_id INTEGER NOT NULL,
            article_id INTEGER,
            article_url TEXT NOT NULL,
            emission_cbonds_id INTEGER,
            emission_name TEXT NOT NULL,
            emission_url TEXT NOT NULL,
            label TEXT,
            raw_context TEXT,
            fetched_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(article_row_id, emission_url),
            FOREIGN KEY(article_row_id) REFERENCES articles(id)
        );

        CREATE INDEX IF NOT EXISTS idx_article_emissions_article_row_id
        ON article_emissions(article_row_id);

        CREATE INDEX IF NOT EXISTS idx_article_emissions_emission_cbonds_id
        ON article_emissions(emission_cbonds_id);
        """
    )

    columns = existing_columns(conn, "articles")
    additions = [
        ("emissions_checked_at", "TEXT"),
        ("emissions_found_count", "INTEGER NOT NULL DEFAULT 0"),
        ("emissions_fetch_status", "TEXT"),
        ("emissions_fetch_error", "TEXT"),
    ]
    for column_name, column_sql in additions:
        if column_name not in columns:
            conn.execute(f"ALTER TABLE articles ADD COLUMN {column_name} {column_sql}")

    conn.commit()


def fetch_html(url: str) -> str:
    resp = base.request_with_retry(
        "GET",
        url,
        headers={
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Referer": base.SITE_BASE_URL + "/news/",
        },
    )
    return resp.text


def bond_id_from_url(url: str) -> int | None:
    match = BOND_HREF_RE.search(url)
    if not match:
        return None
    return int(match.group(1))


def parent_chain(node: Any, max_depth: int = 6) -> list[Any]:
    result = []
    current = getattr(node, "parent", None)
    for _ in range(max_depth):
        if current is None or getattr(current, "name", None) in {"body", "html"}:
            break
        result.append(current)
        current = getattr(current, "parent", None)
    return result


def has_emission_label_near_anchor(anchor: Any) -> tuple[bool, str | None, str | None]:
    contexts: list[str] = []

    for parent in parent_chain(anchor):
        text = normalize_space(parent.get_text(" ", strip=True))
        if text:
            contexts.append(text)
        if text and EMISSION_LABEL_RE.search(text):
            label = EMISSION_LABEL_RE.search(text).group(0)
            return True, label, text[:600]

    previous_text = normalize_space(anchor.find_previous(string=True))
    next_text = normalize_space(anchor.find_next(string=True))
    sibling_context = normalize_space(" ".join(x for x in [previous_text, anchor.get_text(" ", strip=True), next_text] if x))
    if sibling_context:
        contexts.append(sibling_context)
    if sibling_context and EMISSION_LABEL_RE.search(sibling_context):
        label = EMISSION_LABEL_RE.search(sibling_context).group(0)
        return True, label, sibling_context[:600]

    return False, None, contexts[0][:600] if contexts else None


def parse_emissions_from_html(html: str, page_url: str) -> list[EmissionLink]:
    soup = BeautifulSoup(html, "html.parser")
    emissions_by_url: dict[str, EmissionLink] = {}

    for anchor in soup.find_all("a", href=True):
        href = str(anchor.get("href") or "")
        if not BOND_HREF_RE.search(href):
            continue

        is_emission, label, raw_context = has_emission_label_near_anchor(anchor)
        if not is_emission:
            continue

        emission_url = urljoin(page_url, href)
        emission_name = normalize_space(anchor.get_text(" ", strip=True))
        if not emission_name:
            continue

        emissions_by_url[emission_url] = EmissionLink(
            emission_cbonds_id=bond_id_from_url(emission_url),
            emission_name=emission_name,
            emission_url=emission_url,
            label=label,
            raw_context=raw_context,
        )

    return list(emissions_by_url.values())


def select_target_articles(
    conn: sqlite3.Connection,
    limit: int | None,
    refresh: bool,
    all_linked: bool,
    skip_veb_rf: bool,
    date_to: str | None,
) -> list[sqlite3.Row]:
    if all_linked:
        sql = """
            SELECT DISTINCT a.id, a.article_id, a.url, a.title
            FROM issue_articles ia
            JOIN articles a ON a.article_id = ia.article_id
        """
        where_parts = []
    else:
        sql = """
            SELECT id, article_id, url, title
            FROM articles
        """
        where_parts = [
            "(has_orientir_term = 1 OR has_coupon_spread_term = 1)",
        ]

    if skip_veb_rf:
        where_parts.append(
            """
            (
                NOT EXISTS (
                    SELECT 1
                    FROM issue_articles ia_any
                    WHERE ia_any.article_id = a.article_id
                )
                OR EXISTS (
                    SELECT 1
                    FROM issue_articles ia_non_veb
                    WHERE ia_non_veb.article_id = a.article_id
                      AND REPLACE(LOWER(TRIM(ia_non_veb.emitent_name)), 'ё', 'е')
                          NOT IN ('вэб.рф', 'веб.рф')
                )
            )
            """
        )

    params_list: list[Any] = []
    if date_to:
        where_parts.append(
            """
            DATE(
                CASE
                    WHEN a.date_time GLOB '??.??.????*'
                    THEN SUBSTR(a.date_time, 7, 4) || '-' || SUBSTR(a.date_time, 4, 2) || '-' || SUBSTR(a.date_time, 1, 2)
                    ELSE SUBSTR(a.date_time, 1, 10)
                END
            ) <= DATE(?)
            """
        )
        params_list.append(date_to)

    if not refresh:
        where_parts.append("(emissions_checked_at IS NULL OR emissions_fetch_status = 'error')")

    if where_parts:
        sql += f" WHERE {' AND '.join(where_parts)}"
    sql += " ORDER BY a.id" if all_linked else " ORDER BY id"

    if limit is not None:
        sql += " LIMIT ?"
        params_list.append(limit)

    return conn.execute(sql, tuple(params_list)).fetchall()


def save_emissions(conn: sqlite3.Connection, article: sqlite3.Row, emissions: list[EmissionLink]) -> None:
    now = datetime.now().isoformat(timespec="seconds")

    conn.execute("DELETE FROM article_emissions WHERE article_row_id = ?", (article["id"],))
    conn.executemany(
        """
        INSERT OR IGNORE INTO article_emissions (
            article_row_id, article_id, article_url, emission_cbonds_id,
            emission_name, emission_url, label, raw_context, fetched_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        [
            (
                article["id"],
                article["article_id"],
                article["url"],
                emission.emission_cbonds_id,
                emission.emission_name,
                emission.emission_url,
                emission.label,
                emission.raw_context,
                now,
            )
            for emission in emissions
        ],
    )
    conn.execute(
        """
        UPDATE articles
        SET
            emissions_checked_at = ?,
            emissions_found_count = ?,
            emissions_fetch_status = ?,
            emissions_fetch_error = NULL
        WHERE id = ?
        """,
        (now, len(emissions), "ok", article["id"]),
    )
    conn.commit()


def save_fetch_error(conn: sqlite3.Connection, article: sqlite3.Row, error: Exception) -> None:
    now = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        """
        UPDATE articles
        SET
            emissions_checked_at = ?,
            emissions_fetch_status = ?,
            emissions_fetch_error = ?
        WHERE id = ?
        """,
        (now, "error", str(error)[:1000], article["id"]),
    )
    conn.commit()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Собрать связанные эмиссии со страниц новостей Cbonds.")
    parser.add_argument("--db", default=DB_FILE, help=f"SQLite-база, по умолчанию {DB_FILE}")
    parser.add_argument("--limit", type=int, default=None, help="Ограничить количество новостей для тестового запуска.")
    parser.add_argument("--refresh", action="store_true", help="Перепроверить новости, которые уже были обработаны.")
    parser.add_argument("--delay", type=float, default=DELAY, help=f"Пауза между запросами, по умолчанию {DELAY}.")
    parser.add_argument(
        "--date-to",
        default=DEFAULT_DATE_TO,
        help=f"Обрабатывать только новости с датой не позже YYYY-MM-DD, по умолчанию {DEFAULT_DATE_TO}.",
    )
    parser.add_argument("--no-date-to", action="store_true", help="Отключить ограничение верхней даты новости.")
    parser.add_argument("--no-cookie", action="store_true", help="Не подключать Cookie даже из env/файла.")
    parser.add_argument(
        "--all-linked",
        action="store_true",
        help="Проверить все статьи, привязанные к выпускам в issue_articles, без фильтра has_orientir/has_coupon_spread.",
    )
    parser.add_argument(
        "--skip-veb-rf",
        action="store_true",
        help="Пропустить статьи, которые привязаны только к эмитенту ВЭБ.РФ / ВЭБ.РФ.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    db_path = Path(args.db)
    if not db_path.exists():
        raise FileNotFoundError(f"Не найдена база: {db_path.resolve()}")

    if args.no_cookie:
        base.session.headers.pop("Cookie", None)
    else:
        cookie_header = base.load_cookie_header()
        if cookie_header:
            base.session.headers["Cookie"] = cookie_header

    print(f"Cookie: {'загружены' if 'Cookie' in base.session.headers else 'не заданы'}")
    print(f"База: {db_path.resolve()}")

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    ensure_schema(conn)

    articles = select_target_articles(
        conn,
        limit=args.limit,
        refresh=args.refresh,
        all_linked=args.all_linked,
        skip_veb_rf=args.skip_veb_rf,
        date_to=None if args.no_date_to else args.date_to,
    )
    print(f"Новостей к проверке: {len(articles)}")
    print(f"Режим: {'all-linked issue_articles' if args.all_linked else 'orientir/coupon flags'}")
    print(f"Дата новости не позже: {'без ограничения' if args.no_date_to else args.date_to}")
    if args.skip_veb_rf:
        print("ВЭБ.РФ: пропускаем статьи, привязанные только к этому эмитенту")

    found_total = 0
    error_total = 0
    for idx, article in enumerate(articles, start=1):
        print(f"[{idx}/{len(articles)}] {article['article_id']} | {article['url']}")
        try:
            html = fetch_html(article["url"])
            emissions = parse_emissions_from_html(html, article["url"])
            save_emissions(conn, article, emissions)
            found_total += len(emissions)
            print(f"    эмиссий найдено: {len(emissions)}")
            for emission in emissions:
                print(f"    - {emission.emission_name} | {emission.emission_url}")
        except Exception as exc:  # noqa: BLE001
            error_total += 1
            save_fetch_error(conn, article, exc)
            print(f"    [error] {exc}")

        if idx < len(articles):
            time.sleep(args.delay)

    conn.close()
    print(f"\nГотово. Найдено связей article_emissions: {found_total}; ошибок: {error_total}")


if __name__ == "__main__":
    main()
