#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Проверка результатов Gemini по объему размещения.

Для каждой строки Excel:
1. Берем `bond_name` в формате "эмитент, номер выпуска".
2. Берем часть после первой запятой как номер выпуска.
3. Проверяем, есть ли этот номер:
   - в `volume_gemini_pre_source_title` ИЛИ `volume_gemini_pre_evidence_text`;
   - в `volume_gemini_final_source_title` ИЛИ `volume_gemini_final_evidence_text`.
4. Помечаем строку только если все четыре проверяемые ячейки непустые
   и хотя бы одна из двух пар не содержит номер выпуска.

Скрипт добавляет/обновляет колонки:
- volume_issue_number
- volume_pre_issue_number_found
- volume_final_issue_number_found
- volume_issue_number_mismatch
- volume_issue_number_mismatch_reason

По умолчанию файл обновляется на месте.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any

import openpyxl


DEFAULT_WORKBOOK = Path("panel_final_v2_placement_volume_gemini.xlsx")
DEFAULT_SHEET = "panel"

REQUIRED_COLUMNS = [
    "bond_name",
    "volume_gemini_pre_source_title",
    "volume_gemini_pre_evidence_text",
    "volume_gemini_final_source_title",
    "volume_gemini_final_evidence_text",
]

OUTPUT_COLUMNS = [
    "volume_issue_number",
    "volume_pre_issue_number_found",
    "volume_final_issue_number_found",
    "volume_issue_number_mismatch",
    "volume_issue_number_mismatch_reason",
]

CYR_TO_LAT = str.maketrans(
    {
        "А": "A",
        "В": "B",
        "Е": "E",
        "К": "K",
        "М": "M",
        "Н": "H",
        "О": "O",
        "Р": "P",
        "С": "C",
        "Т": "T",
        "У": "Y",
        "Х": "X",
        "а": "A",
        "в": "B",
        "е": "E",
        "к": "K",
        "м": "M",
        "н": "H",
        "о": "O",
        "р": "P",
        "с": "C",
        "т": "T",
        "у": "Y",
        "х": "X",
    }
)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def normalize_for_issue_match(value: Any) -> str:
    text = normalize_text(value).translate(CYR_TO_LAT).upper()
    return re.sub(r"[^0-9A-ZА-Я]+", "", text)


def extract_issue_number(bond_name: Any) -> str:
    text = normalize_text(bond_name)
    if "," not in text:
        return ""
    return text.split(",", 1)[1].strip()


def contains_issue_number(haystack_parts: list[Any], issue_number: str) -> bool:
    needle = normalize_for_issue_match(issue_number)
    if not needle:
        return False
    haystack = normalize_for_issue_match(" ".join(normalize_text(x) for x in haystack_parts))
    return needle in haystack


def all_nonblank(values: list[Any]) -> bool:
    return all(normalize_text(value) for value in values)


def build_header_index(ws) -> dict[str, int]:
    return {normalize_text(ws.cell(1, col).value): col for col in range(1, ws.max_column + 1)}


def ensure_columns(ws, headers: dict[str, int], columns: list[str]) -> dict[str, int]:
    for column in columns:
        if column not in headers:
            col_idx = ws.max_column + 1
            ws.cell(1, col_idx, column)
            headers[column] = col_idx
    return headers


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Flag volume Gemini rows where source/evidence does not mention issue number.")
    parser.add_argument("--input", type=Path, default=DEFAULT_WORKBOOK, help=f"Workbook to check, default: {DEFAULT_WORKBOOK}")
    parser.add_argument("--output", type=Path, default=None, help="Output workbook. If omitted, updates input file in place.")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help=f"Sheet name, default: {DEFAULT_SHEET!r}")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_path = args.output or args.input

    wb = openpyxl.load_workbook(args.input)
    if args.sheet not in wb.sheetnames:
        raise RuntimeError(f"Лист {args.sheet!r} не найден. Доступные листы: {', '.join(wb.sheetnames)}")
    ws = wb[args.sheet]

    headers = build_header_index(ws)
    missing = [column for column in REQUIRED_COLUMNS if column not in headers]
    if missing:
        raise RuntimeError(f"Не найдены обязательные колонки: {missing}")
    headers = ensure_columns(ws, headers, OUTPUT_COLUMNS)

    checked = 0
    flagged = 0
    no_issue_number = 0
    skipped_blank_cells = 0

    for row_idx in range(2, ws.max_row + 1):
        bond_name = ws.cell(row_idx, headers["bond_name"]).value
        issue_number = extract_issue_number(bond_name)
        if not issue_number:
            no_issue_number += 1

        pre_values = [
            ws.cell(row_idx, headers["volume_gemini_pre_source_title"]).value,
            ws.cell(row_idx, headers["volume_gemini_pre_evidence_text"]).value,
        ]
        final_values = [
            ws.cell(row_idx, headers["volume_gemini_final_source_title"]).value,
            ws.cell(row_idx, headers["volume_gemini_final_evidence_text"]).value,
        ]
        has_all_checked_cells = all_nonblank([*pre_values, *final_values])

        pre_found = contains_issue_number(
            pre_values,
            issue_number,
        )
        final_found = contains_issue_number(
            final_values,
            issue_number,
        )

        reasons = []
        if not issue_number:
            reasons.append("Не удалось извлечь номер выпуска из bond_name")
        if issue_number and not has_all_checked_cells:
            skipped_blank_cells += 1
        if issue_number and has_all_checked_cells and not pre_found:
            reasons.append("Номер выпуска не найден в pre source_title/evidence_text")
        if issue_number and has_all_checked_cells and not final_found:
            reasons.append("Номер выпуска не найден в final source_title/evidence_text")

        mismatch = bool(reasons) and (not issue_number or has_all_checked_cells)
        checked += 1
        flagged += int(mismatch)

        ws.cell(row_idx, headers["volume_issue_number"], issue_number)
        ws.cell(row_idx, headers["volume_pre_issue_number_found"], int(pre_found))
        ws.cell(row_idx, headers["volume_final_issue_number_found"], int(final_found))
        ws.cell(row_idx, headers["volume_issue_number_mismatch"], int(mismatch))
        ws.cell(row_idx, headers["volume_issue_number_mismatch_reason"], "; ".join(reasons))

    wb.save(output_path)
    print(f"Workbook: {output_path}")
    print(f"Checked rows: {checked}")
    print(f"Flagged rows: {flagged}")
    print(f"Rows without issue number: {no_issue_number}")
    print(f"Rows skipped because at least one checked cell is blank: {skipped_blank_cells}")


if __name__ == "__main__":
    main()
