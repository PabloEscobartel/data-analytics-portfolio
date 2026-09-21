#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Combine GJRM bivariate-probit cluster-bootstrap SE files into one table.

Usage:
    python combine_gjrm_bootstrap_se.py
    python combine_gjrm_bootstrap_se.py gjrm_biprobit_bootstrap_clean
    python combine_gjrm_bootstrap_se.py gjrm_biprobit_bootstrap_clean gjrm_biprobit_bootstrap_combined.csv
    python combine_gjrm_bootstrap_se.py --all-samples gjrm_biprobit_bootstrap_clean gjrm_biprobit_bootstrap_combined.csv --xlsx gjrm_biprobit_bootstrap_combined.xlsx
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd


DEFAULT_INPUT_DIR = "gjrm_biprobit_bootstrap_clean"
DEFAULT_OUTPUT_CSV = "gjrm_biprobit_bootstrap_se_combined.csv"
SAMPLE_LABELS = {
    "all": "All bonds",
    "fixed": "Fixed coupon bonds",
    "floater": "Floating-rate bonds",
}


def parse_scenario(path: Path) -> dict:
    match = re.match(r"^(open|close)_(\d+)_cluster_bootstrap_se\.csv$", path.name)
    if not match:
        return {
            "window": None,
            "min_messages": None,
            "scenario": path.stem.replace("_cluster_bootstrap_se", ""),
        }

    short_window, min_messages = match.groups()
    window = "book_open" if short_window == "open" else "book_close"
    return {
        "window": window,
        "min_messages": int(min_messages),
        "scenario": f"{short_window}_{min_messages}",
    }


def combine_bootstrap_files(input_dir: Path, sample_mode: str = "all") -> pd.DataFrame:
    paths = sorted(input_dir.glob("*_cluster_bootstrap_se.csv"))
    if not paths:
        raise FileNotFoundError(f"No *_cluster_bootstrap_se.csv files found in {input_dir}")

    frames = []
    for path in paths:
        meta = parse_scenario(path)
        df = pd.read_csv(path)
        df.insert(0, "source_file", str(path))
        df.insert(0, "sample_label", SAMPLE_LABELS.get(sample_mode, sample_mode))
        df.insert(0, "sample_mode", sample_mode)
        df.insert(0, "scenario", meta["scenario"])
        df.insert(0, "min_messages", meta["min_messages"])
        df.insert(0, "window", meta["window"])
        frames.append(df)

    out = pd.concat(frames, ignore_index=True)

    sort_cols = [
        col for col in ["window", "min_messages", "equation", "term"]
        if col in out.columns
    ]
    if sort_cols:
        out = out.sort_values(sort_cols, na_position="last").reset_index(drop=True)

    return out


def sample_input_dirs(base_input_dir: Path) -> list[tuple[str, Path]]:
    return [
        ("all", base_input_dir),
        ("fixed", base_input_dir.with_name(f"{base_input_dir.name}_fixed")),
        ("floater", base_input_dir.with_name(f"{base_input_dir.name}_floater")),
    ]


def main() -> None:
    parser = argparse.ArgumentParser(description="Combine GJRM bootstrap SE CSV files.")
    parser.add_argument("input_dir", nargs="?", default=DEFAULT_INPUT_DIR)
    parser.add_argument("output_csv", nargs="?", default=DEFAULT_OUTPUT_CSV)
    parser.add_argument(
        "--all-samples",
        action="store_true",
        help="Combine base, _fixed, and _floater output directories into one file.",
    )
    parser.add_argument(
        "--xlsx",
        default=None,
        help="Optional XLSX output path. Defaults to output_csv with .xlsx suffix if provided without value is not supported.",
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_csv = Path(args.output_csv)

    if args.all_samples:
        frames = []
        missing = []
        for sample_mode, sample_dir in sample_input_dirs(input_dir):
            try:
                frames.append(combine_bootstrap_files(sample_dir, sample_mode=sample_mode))
            except FileNotFoundError:
                missing.append(str(sample_dir))
        if not frames:
            raise FileNotFoundError(f"No bootstrap files found. Checked: {', '.join(missing)}")
        combined = pd.concat(frames, ignore_index=True)
        sort_cols = [col for col in ["sample_mode", "window", "min_messages", "equation", "term"] if col in combined.columns]
        combined = combined.sort_values(sort_cols, na_position="last").reset_index(drop=True)
        if missing:
            print("Missing sample dirs skipped:", ", ".join(missing))
    else:
        combined = combine_bootstrap_files(input_dir)
    combined.to_csv(output_csv, index=False, encoding="utf-8-sig")

    if args.xlsx:
        combined.to_excel(args.xlsx, index=False)

    print(f"Combined rows: {len(combined)}")
    print(f"Saved CSV: {output_csv}")
    if args.xlsx:
        print(f"Saved XLSX: {args.xlsx}")


if __name__ == "__main__":
    main()
