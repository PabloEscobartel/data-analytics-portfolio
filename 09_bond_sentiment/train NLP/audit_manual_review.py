#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path
from typing import Dict, List

import pandas as pd


# -----------------------------
# helpers
# -----------------------------

def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def read_sql(conn: sqlite3.Connection, query: str) -> pd.DataFrame:
    return pd.read_sql_query(query, conn)


def add_review_columns(df: pd.DataFrame, group_name: str) -> pd.DataFrame:
    df = df.copy()
    df.insert(0, "review_group", group_name)
    df["audit_label"] = ""   # 1 = correct, 0 = incorrect, empty = not reviewed
    df["audit_note"] = ""
    return df


def normalize_label(x):
    if pd.isna(x):
        return None
    s = str(x).strip().lower()
    if s in {"1", "true", "yes", "y", "correct", "tp"}:
        return 1
    if s in {"0", "false", "no", "n", "incorrect", "fp"}:
        return 0
    return None


# -----------------------------
# export samples
# -----------------------------

def export_telegram_samples(conn: sqlite3.Connection, out_dir: Path, n: int) -> List[Path]:
    files = []

    queries: Dict[str, str] = {
        # 1) suppress сработал и сообщение реально выкинуто
        "tg_suppressed_dropped": f"""
            SELECT
                r.source_key,
                r.channel,
                r.message_id,
                r.parent_id,
                r.published_at,
                r.explicit_status,
                r.match_origin,
                r.issuer,
                r.entity_level,
                r.offering_id,
                r.bond_name,
                r.isin,
                r.matched_aliases,
                r.alias_types,
                r.match_score,
                r.resolution_note,
                r.inherit_anchor_key,
                r.inherit_depth,
                r.text_len,
                r.is_short_reply,
                r.bond_context,
                r.primary_context,
                r.broker_noise,
                m.text AS message_text,
                p.text AS parent_text
            FROM telegram_matches_resolved r
            JOIN messages m
              ON r.channel = m.channel AND r.message_id = m.id
            LEFT JOIN messages p
              ON r.channel = p.channel AND r.parent_id = p.id
            WHERE r.explicit_status = 'suppressed'
              AND r.issuer IS NULL
            ORDER BY RANDOM()
            LIMIT {int(n)}
        """,

        # 2) suppress сработал, но потом сообщение унаследовало сущность
        "tg_suppressed_but_inherited": f"""
            SELECT
                r.source_key,
                r.channel,
                r.message_id,
                r.parent_id,
                r.published_at,
                r.explicit_status,
                r.match_origin,
                r.issuer,
                r.entity_level,
                r.offering_id,
                r.bond_name,
                r.isin,
                r.matched_aliases,
                r.alias_types,
                r.match_score,
                r.resolution_note,
                r.inherit_anchor_key,
                r.inherit_depth,
                r.text_len,
                r.is_short_reply,
                r.bond_context,
                r.primary_context,
                r.broker_noise,
                m.text AS message_text,
                p.text AS parent_text
            FROM telegram_matches_resolved r
            JOIN messages m
              ON r.channel = m.channel AND r.message_id = m.id
            LEFT JOIN messages p
              ON r.channel = p.channel AND r.parent_id = p.id
            WHERE r.explicit_status = 'suppressed'
              AND r.match_origin = 'inherited'
            ORDER BY RANDOM()
            LIMIT {int(n)}
        """,

        # 3) самая подозрительная группа inherited
        "tg_inherited_issuer_no_bond": f"""
            SELECT
                r.source_key,
                r.channel,
                r.message_id,
                r.parent_id,
                r.published_at,
                r.explicit_status,
                r.match_origin,
                r.issuer,
                r.entity_level,
                r.offering_id,
                r.bond_name,
                r.isin,
                r.matched_aliases,
                r.alias_types,
                r.match_score,
                r.resolution_note,
                r.inherit_anchor_key,
                r.inherit_depth,
                r.text_len,
                r.is_short_reply,
                r.bond_context,
                r.primary_context,
                r.broker_noise,
                m.text AS reply_text,
                p.text AS parent_text
            FROM telegram_matches_resolved r
            JOIN messages m
              ON r.channel = m.channel AND r.message_id = m.id
            LEFT JOIN messages p
              ON r.channel = p.channel AND r.parent_id = p.id
            WHERE r.match_origin = 'inherited'
              AND r.entity_level = 'issuer'
              AND r.bond_context = 0
            ORDER BY RANDOM()
            LIMIT {int(n)}
        """,

        # 4) offering-level, но evidence слабее: code_exact/code_variant без isin/full name
        "tg_offering_code_only": f"""
            SELECT
                r.source_key,
                r.channel,
                r.message_id,
                r.parent_id,
                r.published_at,
                r.explicit_status,
                r.match_origin,
                r.issuer,
                r.entity_level,
                r.offering_id,
                r.bond_name,
                r.isin,
                r.matched_aliases,
                r.alias_types,
                r.match_score,
                r.resolution_note,
                r.inherit_anchor_key,
                r.inherit_depth,
                r.text_len,
                r.is_short_reply,
                r.bond_context,
                r.primary_context,
                r.broker_noise,
                m.text AS message_text,
                p.text AS parent_text
            FROM telegram_matches_resolved r
            JOIN messages m
              ON r.channel = m.channel AND r.message_id = m.id
            LEFT JOIN messages p
              ON r.channel = p.channel AND r.parent_id = p.id
            WHERE r.offering_id IS NOT NULL
              AND (
                    r.alias_types LIKE '%code_exact%'
                 OR r.alias_types LIKE '%code_variant%'
              )
              AND r.alias_types NOT LIKE '%isin%'
              AND r.alias_types NOT LIKE '%bond_name_full%'
            ORDER BY RANDOM()
            LIMIT {int(n)}
        """,
    }

    for group_name, query in queries.items():
        df = read_sql(conn, query)
        df = add_review_columns(df, group_name)
        path = out_dir / f"{group_name}.csv"
        df.to_csv(path, index=False, encoding="utf-8-sig")
        files.append(path)

    return files


def export_smartlab_samples(conn: sqlite3.Connection, out_dir: Path, n: int) -> List[Path]:
    files = []

    queries: Dict[str, str] = {
        "sl_posts_suppressed": f"""
            SELECT
                r.source_key,
                r.post_id,
                r.published_at,
                r.explicit_status,
                r.match_origin,
                r.issuer,
                r.entity_level,
                r.offering_id,
                r.bond_name,
                r.isin,
                r.matched_aliases,
                r.alias_types,
                r.match_score,
                r.resolution_note,
                r.bond_context,
                r.primary_context,
                r.broker_noise,
                p.title,
                p.tags,
                p.text AS post_text
            FROM smartlab_post_matches_resolved r
            JOIN posts p
              ON r.post_id = p.post_id
            WHERE r.explicit_status = 'suppressed'
            ORDER BY RANDOM()
            LIMIT {int(n)}
        """,

        "sl_comments_suppressed": f"""
            SELECT
                r.source_key,
                r.comment_id,
                r.post_id,
                r.published_at,
                r.explicit_status,
                r.match_origin,
                r.issuer,
                r.entity_level,
                r.offering_id,
                r.bond_name,
                r.isin,
                r.matched_aliases,
                r.alias_types,
                r.match_score,
                r.resolution_note,
                r.bond_context,
                r.primary_context,
                r.broker_noise,
                c.text AS comment_text,
                p.title AS post_title
            FROM smartlab_comment_matches_resolved r
            JOIN comments c
              ON r.comment_id = c.comment_id
            LEFT JOIN posts p
              ON r.post_id = p.post_id
            WHERE r.explicit_status = 'suppressed'
            ORDER BY RANDOM()
            LIMIT {int(n)}
        """,
    }

    for group_name, query in queries.items():
        df = read_sql(conn, query)
        df = add_review_columns(df, group_name)
        path = out_dir / f"{group_name}.csv"
        df.to_csv(path, index=False, encoding="utf-8-sig")
        files.append(path)

    return files


# -----------------------------
# summarize labels
# -----------------------------

def summarize_reviews(out_dir: Path) -> pd.DataFrame:
    rows = []

    for csv_path in sorted(out_dir.glob("*.csv")):
        if csv_path.name in {
            "manual_review_summary.csv",
            "manual_review_row_details.csv",
        }:
            continue

        try:
            df = pd.read_csv(csv_path)
        except Exception:
            continue

        if "audit_label" not in df.columns:
            continue

        labels = df["audit_label"].apply(normalize_label)
        reviewed = labels.notna()

        n_total = len(df)
        n_reviewed = int(reviewed.sum())
        n_correct = int((labels == 1).sum())
        n_incorrect = int((labels == 0).sum())
        precision = (n_correct / n_reviewed) if n_reviewed else None

        rows.append({
            "file": csv_path.name,
            "review_group": df["review_group"].iloc[0] if "review_group" in df.columns and len(df) else csv_path.stem,
            "n_total_rows": n_total,
            "n_reviewed": n_reviewed,
            "n_correct": n_correct,
            "n_incorrect": n_incorrect,
            "precision": precision,
        })

    summary = pd.DataFrame(rows)
    if not summary.empty:
        summary = summary.sort_values(["precision", "n_reviewed"], ascending=[True, False], na_position="last")
        summary.to_csv(out_dir / "manual_review_summary.csv", index=False, encoding="utf-8-sig")

    return summary


# -----------------------------
# main
# -----------------------------

def main():
    parser = argparse.ArgumentParser(description="Export manual QA samples and summarize labels.")
    parser.add_argument("--telegram-db", type=Path, default=Path("data.db"))
    parser.add_argument("--smartlab-db", type=Path, default=Path("smartlab_bonds.db"))
    parser.add_argument("--out-dir", type=Path, default=Path("manual_review_output"))
    parser.add_argument("--sample-size", type=int, default=50)
    parser.add_argument(
        "--mode",
        choices=["export", "summarize", "both"],
        default="both",
        help="export samples, summarize reviewed CSVs, or both",
    )
    args = parser.parse_args()

    ensure_dir(args.out_dir)

    if args.mode in {"export", "both"}:
        tg_conn = sqlite3.connect(args.telegram_db)
        sl_conn = sqlite3.connect(args.smartlab_db)

        tg_files = export_telegram_samples(tg_conn, args.out_dir, args.sample_size)
        sl_files = export_smartlab_samples(sl_conn, args.out_dir, args.sample_size)

        tg_conn.close()
        sl_conn.close()

        print("Exported files:")
        for p in tg_files + sl_files:
            print(f"  - {p}")

        print("\nРазметка:")
        print("  audit_label = 1  -> match/filter looks correct")
        print("  audit_label = 0  -> incorrect / suspicious")
        print("  audit_note       -> short comment why")

    if args.mode in {"summarize", "both"}:
        summary = summarize_reviews(args.out_dir)
        if summary.empty:
            print("\nNo reviewed CSVs found yet.")
        else:
            print("\nManual review summary:")
            print(summary.to_string(index=False))


if __name__ == "__main__":
    main()
