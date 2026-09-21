#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пост-обработка результатов Gemini: для диапазонов ориентира берёт верхнюю границу.

Пример:
    7.75–8%      -> 8
    275-300 б.п. -> 300
    0.85–1%      -> 1

Обновляет:
- JSON-кэш Gemini;
- Excel-файл с результатами;
- таблицу issue_orientir_bookbuilding в SQLite.

Старое значение сохраняется в JSON-кэше в поле orient_value_original,
если оно было изменено.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sqlite3
from pathlib import Path
from typing import Any

import openpyxl


DEFAULT_CACHE = Path("gemini_bookbuilding_orientir_multi_news_cache.json")
DEFAULT_WORKBOOK = Path("bonds_offer_updated_cbonds_orientir_gemini.xlsx")
DEFAULT_DB = Path("cbonds_news.sqlite3")
DEFAULT_SHEET = "main_no_na"

RANGE_RE = re.compile(
    r"(?<!\d)(\d{1,3}(?:[.,]\d+)?)\s*(?:%|б\.?\s*п\.?|bp|bps)?\s*[-–—−]\s*(\d{1,3}(?:[.,]\d+)?)\s*(?:%|б\.?\s*п\.?|bp|bps)?(?!\d)",
    re.IGNORECASE,
)


def normalize_number(value: str) -> float:
    return float(value.replace(",", "."))


def upper_bound_from_text(text: Any) -> float | None:
    if not text:
        return None
    text = str(text).replace("\xa0", " ")
    matches = list(RANGE_RE.finditer(text))
    if not matches:
        return None

    candidates: list[float] = []
    for match in matches:
        left = normalize_number(match.group(1))
        right = normalize_number(match.group(2))

        # Отсекаем очевидные диапазоны дат/лет, если такие случайно попадут в evidence_text.
        if left > 1000 or right > 1000:
            continue
        candidates.append(max(left, right))

    if not candidates:
        return None
    return candidates[0]


def safe_number(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def load_cache(path: Path) -> dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_cache(path: Path, cache: dict[str, Any]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def adjust_cache(cache: dict[str, Any]) -> tuple[dict[str, Any], int]:
    changed = 0
    for item in cache.values():
        if not isinstance(item, dict):
            continue
        orient_type = item.get("orient_type")
        if orient_type not in {"coupon", "spread"}:
            continue

        upper = upper_bound_from_text(item.get("value_text"))
        if upper is None:
            continue

        current = safe_number(item.get("orient_value"))
        if current is None or abs(current - upper) < 1e-9:
            continue

        item.setdefault("orient_value_original", current)
        item["orient_value"] = upper
        if orient_type == "coupon":
            item["coupon_guide"] = upper
            item["spread_guide"] = None
        elif orient_type == "spread":
            item["spread_guide"] = upper
            item["coupon_guide"] = None
        item["upper_bound_adjusted"] = 1
        changed += 1

    return cache, changed


def update_workbook(workbook_path: Path, sheet_name: str, cache: dict[str, Any]) -> int:
    wb = openpyxl.load_workbook(workbook_path)
    ws = wb[sheet_name]
    headers = {str(ws.cell(1, c).value).strip(): c for c in range(1, ws.max_column + 1)}

    required = [
        "bb_gemini_orient_type",
        "bb_gemini_orient_value",
        "bb_gemini_coupon_guide",
        "bb_gemini_spread_guide",
    ]
    missing = [name for name in required if name not in headers]
    if missing:
        raise RuntimeError(f"В Excel не найдены колонки: {missing}")

    changed = 0
    for key, item in cache.items():
        if not isinstance(item, dict):
            continue
        try:
            row_idx = int(key)
        except ValueError:
            continue
        if row_idx < 2 or row_idx > ws.max_row:
            continue
        if not item.get("upper_bound_adjusted"):
            continue

        orient_type = item.get("orient_type")
        orient_value = item.get("orient_value")
        ws.cell(row_idx, headers["bb_gemini_orient_value"], orient_value)
        ws.cell(row_idx, headers["bb_gemini_coupon_guide"], orient_value if orient_type == "coupon" else None)
        ws.cell(row_idx, headers["bb_gemini_spread_guide"], orient_value if orient_type == "spread" else None)
        changed += 1

    wb.save(workbook_path)
    return changed


def update_db(db_path: Path, cache: dict[str, Any]) -> int:
    conn = sqlite3.connect(db_path)
    changed = 0
    for key, item in cache.items():
        if not isinstance(item, dict) or not item.get("upper_bound_adjusted"):
            continue
        try:
            excel_row = int(key)
        except ValueError:
            continue
        orient_type = item.get("orient_type")
        orient_value = item.get("orient_value")
        cursor = conn.execute(
            """
            UPDATE issue_orientir_bookbuilding
            SET
                orient_value = ?,
                coupon_guide = ?,
                spread_guide = ?
            WHERE excel_row = ?
            """,
            (
                orient_value,
                orient_value if orient_type == "coupon" else None,
                orient_value if orient_type == "spread" else None,
                excel_row,
            ),
        )
        changed += max(cursor.rowcount, 0)
    conn.commit()
    conn.close()
    return changed


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Заменить нижнюю границу диапазона ориентира на верхнюю.")
    parser.add_argument("--cache", type=Path, default=DEFAULT_CACHE)
    parser.add_argument("--workbook", type=Path, default=DEFAULT_WORKBOOK)
    parser.add_argument("--db", type=Path, default=DEFAULT_DB)
    parser.add_argument("--sheet", default=DEFAULT_SHEET)
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    cache = load_cache(args.cache)
    cache, changed_cache = adjust_cache(cache)

    print(f"Диапазонов к корректировке: {changed_cache}")
    if args.dry_run:
        return

    save_cache(args.cache, cache)
    changed_workbook = update_workbook(args.workbook, args.sheet, cache)
    changed_db = update_db(args.db, cache)

    print(f"Кэш обновлён: {args.cache}")
    print(f"Excel строк обновлено: {changed_workbook}")
    print(f"SQLite update operations: {changed_db}")


if __name__ == "__main__":
    main()
