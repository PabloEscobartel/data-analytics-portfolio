#!/usr/bin/env python3
"""
Разметка новостей Cbonds по наличию контекста ориентира купона / спреда.

Скрипт добавляет в таблицу articles флаги:
- has_orientir_term: есть слово "ориентир" с вариациями;
- has_coupon_spread_term: есть "купон" или "спред" с вариациями;
- has_orientir_coupon_spread_context: оба признака найдены рядом в тексте;
- orientir_context_snippet: короткий фрагмент вокруг найденного контекста;
- orientir_flags_updated_at: время последнего пересчёта флагов.

По умолчанию обновляет существующую cbonds_news.sqlite3.
Если нужно создать отдельную размеченную копию, используй --copy-to.

Запуск:
    python flag_cbonds_orientir_news.py

Создать новую базу-копию и разметить её:
    python flag_cbonds_orientir_news.py --copy-to cbonds_news_orientir.sqlite3
"""

from __future__ import annotations

import argparse
import html
import re
import shutil
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Iterable


DB_FILE = "cbonds_news.sqlite3"

ORIENTIR_RE = re.compile(
    r"(?iu)\b(?:ориентир\w*|guidance)\b"
)
COUPON_SPREAD_RE = re.compile(
    r"(?iu)\b(?:купон\w*|спред\w*|coupon\w*|spread\w*)\b"
)

CONTEXT_WINDOW_CHARS = 280
SNIPPET_RADIUS_CHARS = 220


def normalize_text(value: str | None) -> str:
    if not value:
        return ""
    text = html.unescape(value)
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def combined_article_text(title: str | None, introtext: str | None, content_text: str | None) -> str:
    return normalize_text(" ".join(part for part in (title, introtext, content_text) if part))


def has_near_context(orientir_matches: Iterable[re.Match], coupon_spread_matches: Iterable[re.Match]) -> tuple[int, int | None]:
    orientir_positions = [(m.start(), m.end()) for m in orientir_matches]
    coupon_spread_positions = [(m.start(), m.end()) for m in coupon_spread_matches]

    best_distance: int | None = None
    best_position: int | None = None

    for o_start, o_end in orientir_positions:
        for c_start, c_end in coupon_spread_positions:
            if c_start >= o_end:
                distance = c_start - o_end
            elif o_start >= c_end:
                distance = o_start - c_end
            else:
                distance = 0

            if best_distance is None or distance < best_distance:
                best_distance = distance
                best_position = min(o_start, c_start)

    if best_distance is not None and best_distance <= CONTEXT_WINDOW_CHARS:
        return 1, best_position
    return 0, best_position


def make_snippet(text: str, center: int | None) -> str | None:
    if center is None:
        return None

    start = max(0, center - SNIPPET_RADIUS_CHARS)
    end = min(len(text), center + SNIPPET_RADIUS_CHARS)
    snippet = text[start:end].strip()

    if start > 0:
        snippet = "..." + snippet
    if end < len(text):
        snippet += "..."

    return snippet


def classify_article(title: str | None, introtext: str | None, content_text: str | None) -> tuple[int, int, int, str | None]:
    text = combined_article_text(title, introtext, content_text)
    orientir_matches = list(ORIENTIR_RE.finditer(text))
    coupon_spread_matches = list(COUPON_SPREAD_RE.finditer(text))

    has_orientir = int(bool(orientir_matches))
    has_coupon_spread = int(bool(coupon_spread_matches))
    has_context, context_position = has_near_context(orientir_matches, coupon_spread_matches)

    # Если оба слова есть в заголовке/тексте, но дальше выбранного окна, всё равно
    # оставляем отдельные флаги; context-флаг более строгий и удобен для кандидатов.
    snippet = make_snippet(text, context_position if has_context else None)
    return has_orientir, has_coupon_spread, has_context, snippet


def existing_columns(conn: sqlite3.Connection, table_name: str) -> set[str]:
    rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    return {str(row[1]) for row in rows}


def ensure_columns(conn: sqlite3.Connection) -> None:
    columns = existing_columns(conn, "articles")
    additions = [
        ("has_orientir_term", "INTEGER NOT NULL DEFAULT 0"),
        ("has_coupon_spread_term", "INTEGER NOT NULL DEFAULT 0"),
        ("has_orientir_coupon_spread_context", "INTEGER NOT NULL DEFAULT 0"),
        ("orientir_context_snippet", "TEXT"),
        ("orientir_flags_updated_at", "TEXT"),
    ]

    for column_name, column_sql in additions:
        if column_name not in columns:
            conn.execute(f"ALTER TABLE articles ADD COLUMN {column_name} {column_sql}")

    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_articles_orientir_flags
        ON articles (
            has_orientir_coupon_spread_context,
            has_orientir_term,
            has_coupon_spread_term
        )
        """
    )
    conn.commit()


def mark_articles(db_path: Path) -> dict[str, int]:
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")

    ensure_columns(conn)

    rows = conn.execute(
        """
        SELECT id, title, introtext, content_text
        FROM articles
        """
    ).fetchall()

    updated_at = datetime.now().isoformat(timespec="seconds")
    stats = {
        "articles": len(rows),
        "has_orientir_term": 0,
        "has_coupon_spread_term": 0,
        "has_orientir_coupon_spread_context": 0,
    }

    updates = []
    for article_id, title, introtext, content_text in rows:
        has_orientir, has_coupon_spread, has_context, snippet = classify_article(title, introtext, content_text)
        stats["has_orientir_term"] += has_orientir
        stats["has_coupon_spread_term"] += has_coupon_spread
        stats["has_orientir_coupon_spread_context"] += has_context
        updates.append((has_orientir, has_coupon_spread, has_context, snippet, updated_at, article_id))

    conn.executemany(
        """
        UPDATE articles
        SET
            has_orientir_term = ?,
            has_coupon_spread_term = ?,
            has_orientir_coupon_spread_context = ?,
            orientir_context_snippet = ?,
            orientir_flags_updated_at = ?
        WHERE id = ?
        """,
        updates,
    )
    conn.commit()
    conn.close()

    return stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Разметить новости Cbonds по словам ориентир / купон / спред.")
    parser.add_argument("--db", default=DB_FILE, help=f"Исходная SQLite-база, по умолчанию {DB_FILE}")
    parser.add_argument(
        "--copy-to",
        default=None,
        help="Если указано, сначала создать копию базы по этому пути и разметить уже её.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source_db = Path(args.db)

    if not source_db.exists():
        raise FileNotFoundError(f"Не найдена база: {source_db.resolve()}")

    target_db = Path(args.copy_to) if args.copy_to else source_db
    if args.copy_to:
        shutil.copyfile(source_db, target_db)

    stats = mark_articles(target_db)

    print(f"База: {target_db.resolve()}")
    print(f"Новостей обработано: {stats['articles']}")
    print(f"has_orientir_term=1: {stats['has_orientir_term']}")
    print(f"has_coupon_spread_term=1: {stats['has_coupon_spread_term']}")
    print(f"has_orientir_coupon_spread_context=1: {stats['has_orientir_coupon_spread_context']}")


if __name__ == "__main__":
    main()
