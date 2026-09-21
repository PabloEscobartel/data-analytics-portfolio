#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build the Gemini input universe for primary-market issuer sentiment.

This version is compatible with bond_entity_matcher_v10.py. The matcher is
treated as a candidate generator only: this script does not filter by bond
context, primary context, broker noise, or relevance keywords. It selects all
matched texts for the issuer in the event window before each book close, and
Gemini decides whether each text is relevant to the issuer as a bond issuer.
"""

from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path
from typing import List, Optional

import numpy as np
import pandas as pd


# ---------------------------------------------------------
# Panel / event window helpers
# ---------------------------------------------------------

def _find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    lower_map = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


def _combine_date_and_time(date_series: pd.Series, time_series: pd.Series) -> pd.Series:
    date_part = pd.to_datetime(date_series, errors="coerce")
    time_str = time_series.astype("string").str.strip()

    out = pd.Series(pd.NaT, index=date_part.index, dtype="datetime64[ns]")
    has_date = date_part.notna()
    has_time = has_date & time_str.notna() & (~time_str.isin(["", "NaT", "nan", "None"]))

    if has_time.any():
        combined = pd.to_datetime(
            date_part.loc[has_time].dt.strftime("%Y-%m-%d") + " " + time_str.loc[has_time],
            errors="coerce",
        )
        out.loc[has_time] = combined.values

    fallback = has_date & ~has_time
    if fallback.any():
        out.loc[fallback] = date_part.loc[fallback].dt.normalize().values

    return out


def load_panel(panel_path: Path) -> pd.DataFrame:
    panel = pd.read_excel(panel_path).copy()

    if "offering_id" not in panel.columns:
        panel["offering_id"] = range(1, len(panel) + 1)

    issuer_col = _find_col(panel, ["issuer"])
    bond_col = _find_col(panel, ["bond_name", "bond", "issue_name"])
    isin_col = _find_col(panel, ["ISIN", "isin"])

    if issuer_col is None or bond_col is None or isin_col is None:
        raise ValueError(
            "Panel must contain issuer, bond_name, and ISIN columns "
            "(case-insensitive for ISIN)."
        )

    if issuer_col != "issuer":
        panel = panel.rename(columns={issuer_col: "issuer"})
    if bond_col != "bond_name":
        panel = panel.rename(columns={bond_col: "bond_name"})
    if isin_col != "ISIN":
        panel = panel.rename(columns={isin_col: "ISIN"})

    # Common date / time fields
    book_col = _find_col(panel, ["book_date", "bookbuilding_date", "pricing_date"])
    placement_col = _find_col(panel, ["placement_date", "issue_date"])
    marketing_start_col = _find_col(panel, ["marketing_start", "book_open_date", "book_start_date"])
    pricing_col = _find_col(panel, ["pricing_date", "coupon_fix_date", "book_close_date", "book_date"])

    close_dt_col = _find_col(panel, [
        "bookbuilding_close_dt", "book_close_dt", "bookbuilding_end_dt",
        "book_close_datetime", "bookbuilding_close_datetime",
        "cbonds_book_end_datetime",
    ])
    close_time_col = _find_col(panel, [
        "bookbuilding_close_time", "book_close_time", "book_end_time", "pricing_time"
    ])
    open_dt_col = _find_col(panel, [
        "bookbuilding_open_dt", "book_open_dt", "bookbuilding_start_dt",
        "book_open_datetime", "bookbuilding_open_datetime",
        "cbonds_book_start_datetime",
    ])
    exact_time_flag_col = _find_col(panel, [
        "bookbuilding_exact_time_available", "bookbuilding_close_time_available"
    ])

    if book_col is not None:
        panel["book_date"] = pd.to_datetime(panel[book_col], errors="coerce")
    else:
        panel["book_date"] = pd.NaT

    if placement_col is not None:
        panel["placement_date"] = pd.to_datetime(panel[placement_col], errors="coerce")
    else:
        panel["placement_date"] = pd.NaT

    if marketing_start_col is not None:
        panel["marketing_start"] = pd.to_datetime(panel[marketing_start_col], errors="coerce")
    else:
        panel["marketing_start"] = pd.NaT

    if pricing_col is not None:
        panel["pricing_date"] = pd.to_datetime(panel[pricing_col], errors="coerce")
    else:
        panel["pricing_date"] = pd.NaT

    if open_dt_col is not None:
        panel["bookbuilding_open_dt"] = pd.to_datetime(panel[open_dt_col], errors="coerce")
    else:
        panel["bookbuilding_open_dt"] = pd.NaT

    if close_dt_col is not None:
        panel["bookbuilding_close_dt"] = pd.to_datetime(panel[close_dt_col], errors="coerce")
    elif close_time_col is not None and book_col is not None:
        panel["bookbuilding_close_dt"] = _combine_date_and_time(panel[book_col], panel[close_time_col])
    else:
        panel["bookbuilding_close_dt"] = pd.NaT

    if close_time_col is not None and "bookbuilding_close_time" not in panel.columns:
        panel["bookbuilding_close_time"] = panel[close_time_col]
    elif "bookbuilding_close_time" not in panel.columns and panel["bookbuilding_close_dt"].notna().any():
        panel["bookbuilding_close_time"] = panel["bookbuilding_close_dt"].dt.strftime("%H:%M:%S")
    elif "bookbuilding_close_time" not in panel.columns:
        panel["bookbuilding_close_time"] = pd.NA

    if exact_time_flag_col is not None:
        panel["bookbuilding_exact_time_available"] = pd.to_numeric(
            panel[exact_time_flag_col], errors="coerce"
        ).fillna(0).astype(int)
    elif close_time_col is not None:
        panel["bookbuilding_exact_time_available"] = (
            panel[close_time_col].astype("string").str.strip().fillna("").ne("")
        ).astype(int)
    elif panel["bookbuilding_close_dt"].notna().any():
        panel["bookbuilding_exact_time_available"] = (
            panel["bookbuilding_close_dt"].dt.time.astype("string").str.strip().fillna("").ne("")
        ).astype(int)
    else:
        panel["bookbuilding_exact_time_available"] = 0

    # Event anchor: close time of the order book when available.
    # Window is [event_end - window_days, event_end), i.e. strictly less than close time.
    panel["event_end"] = panel["bookbuilding_close_dt"]
    panel["event_end_source"] = "bookbuilding_close_dt"

    fallback_mask = panel["event_end"].isna() & panel["pricing_date"].notna()
    panel.loc[fallback_mask, "event_end"] = panel.loc[fallback_mask, "pricing_date"]
    panel.loc[fallback_mask, "event_end_source"] = "pricing_date"

    fallback_mask = panel["event_end"].isna() & panel["book_date"].notna()
    panel.loc[fallback_mask, "event_end"] = panel.loc[fallback_mask, "book_date"]
    panel.loc[fallback_mask, "event_end_source"] = "book_date"

    fallback_mask = panel["event_end"].isna() & panel["placement_date"].notna()
    panel.loc[fallback_mask, "event_end"] = panel.loc[fallback_mask, "placement_date"]
    panel.loc[fallback_mask, "event_end_source"] = "placement_date"

    if panel["event_end"].isna().all():
        raise ValueError(
            "Could not infer event end datetime. Need at least one of: "
            "bookbuilding_close_dt / (book_date + bookbuilding_close_time) / pricing_date / book_date / placement_date."
        )

    return panel

def build_event_windows(panel: pd.DataFrame, window_days: int, use_marketing_window: bool) -> pd.DataFrame:
    p = panel.copy()

    if use_marketing_window and "marketing_start" in p.columns:
        p["event_start"] = p["marketing_start"]
        missing_start = p["event_start"].isna()
        p.loc[missing_start, "event_start"] = p.loc[missing_start, "event_end"] - pd.Timedelta(days=window_days)
    else:
        p["event_start"] = p["event_end"] - pd.Timedelta(days=window_days)

    p = p[p["event_end"].notna()].copy()
    p = p[p["event_start"].notna()].copy()
    return p


# ---------------------------------------------------------
# Reading v10 mention candidates
# ---------------------------------------------------------

def table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    q = "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?"
    return conn.execute(q, (table_name,)).fetchone() is not None


def normalize_text_for_csv(value) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).split())


def read_telegram_mentions(conn: sqlite3.Connection) -> pd.DataFrame:
    if not table_exists(conn, "telegram_matches"):
        return pd.DataFrame()
    q = """
        SELECT
            'telegram' AS source,
            r.source_key,
            r.channel,
            CAST(r.message_id AS TEXT) AS local_id,
            r.parent_id,
            r.published_at,
            r.issuer,
            r.entity_level,
            r.offering_id AS matched_offering_id,
            r.bond_name AS matched_bond_name,
            r.isin AS matched_isin,
            r.match_origin,
            r.inherit_anchor_key,
            r.text_len,
            r.matched_aliases,
            r.alias_types,
            m.text AS text,
            p.text AS parent_text
        FROM telegram_matches r
        JOIN messages m
          ON r.channel = m.channel AND r.message_id = m.id
        LEFT JOIN messages p
          ON r.channel = p.channel AND r.parent_id = p.id
    """
    return pd.read_sql_query(q, conn)


def read_smartlab_posts(conn: sqlite3.Connection) -> pd.DataFrame:
    if not table_exists(conn, "smartlab_post_matches"):
        return pd.DataFrame()
    q = """
        SELECT
            'smartlab_post' AS source,
            r.source_key,
            NULL AS channel,
            CAST(r.post_id AS TEXT) AS local_id,
            NULL AS parent_id,
            r.published_at,
            r.issuer,
            r.entity_level,
            r.offering_id AS matched_offering_id,
            r.bond_name AS matched_bond_name,
            r.isin AS matched_isin,
            r.match_origin,
            r.inherit_anchor_key,
            r.text_len,
            r.matched_aliases,
            r.alias_types,
            TRIM(COALESCE(p.title, '') || ' ' || COALESCE(p.tags, '') || ' ' || COALESCE(p.text, '')) AS text,
            NULL AS parent_text
        FROM smartlab_post_matches r
        JOIN posts p ON r.post_id = p.post_id
    """
    return pd.read_sql_query(q, conn)


def read_smartlab_comments(conn: sqlite3.Connection) -> pd.DataFrame:
    if not table_exists(conn, "smartlab_comment_matches"):
        return pd.DataFrame()
    q = """
        SELECT
            'smartlab_comment' AS source,
            r.source_key,
            NULL AS channel,
            CAST(r.comment_id AS TEXT) AS local_id,
            r.post_id AS parent_id,
            r.published_at,
            r.issuer,
            r.entity_level,
            r.offering_id AS matched_offering_id,
            r.bond_name AS matched_bond_name,
            r.isin AS matched_isin,
            r.match_origin,
            r.inherit_anchor_key,
            r.text_len,
            r.matched_aliases,
            r.alias_types,
            c.text AS text,
            TRIM(COALESCE(p.title, '') || ' ' || COALESCE(p.tags, '')) AS parent_text
        FROM smartlab_comment_matches r
        JOIN comments c ON r.comment_id = c.comment_id
        LEFT JOIN posts p ON r.post_id = p.post_id
    """
    return pd.read_sql_query(q, conn)


def load_all_mentions(telegram_db: Path, smartlab_db: Path) -> pd.DataFrame:
    tg_conn = sqlite3.connect(telegram_db)
    sl_conn = sqlite3.connect(smartlab_db)

    parts = [
        read_telegram_mentions(tg_conn),
        read_smartlab_posts(sl_conn),
        read_smartlab_comments(sl_conn),
    ]

    tg_conn.close()
    sl_conn.close()

    parts = [p for p in parts if not p.empty]
    if not parts:
        return pd.DataFrame()

    mentions = pd.concat(parts, ignore_index=True)
    mentions["published_at"] = pd.to_datetime(
        mentions["published_at"], errors="coerce", utc=True, format="mixed"
    ).dt.tz_localize(None)
    mentions = mentions[mentions["published_at"].notna()].copy()
    for col in ["text", "parent_text"]:
        if col in mentions.columns:
            mentions[col] = mentions[col].map(normalize_text_for_csv)
    return mentions


# ---------------------------------------------------------
# Expansion into event-level issuer sentiment universe
# ---------------------------------------------------------

def expand_to_primary_event_universe(mentions: pd.DataFrame, panel_windows: pd.DataFrame) -> pd.DataFrame:
    if mentions.empty or panel_windows.empty:
        return pd.DataFrame()

    panel_small_cols = [
        "offering_id", "issuer", "bond_name", "ISIN",
        "book_date", "placement_date", "marketing_start", "pricing_date",
        "bookbuilding_open_dt", "bookbuilding_close_time", "bookbuilding_close_dt",
        "bookbuilding_exact_time_available", "event_end_source",
        "event_start", "event_end",
    ]
    panel_small = panel_windows[[c for c in panel_small_cols if c in panel_windows.columns]].copy()
    panel_small = panel_small[panel_small["issuer"].notna()].copy()

    min_start = panel_small["event_start"].min()
    max_end = panel_small["event_end"].max()
    mentions = mentions[
        mentions["issuer"].notna()
        & (mentions["published_at"] >= min_start)
        & (mentions["published_at"] < max_end)
    ].copy()
    if mentions.empty:
        return pd.DataFrame()

    # Sentiment target is the issuer, measured before each new primary event.
    # Therefore every mention candidate for the issuer can enter every event
    # window of the same issuer, regardless of whether the matcher hit the
    # current issue, another issue, or only an issuer alias.
    uni = mentions.merge(panel_small, on="issuer", how="inner", suffixes=("_match", ""))
    uni = uni[
        (uni["published_at"] >= uni["event_start"])
        & (uni["published_at"] < uni["event_end"])
    ].copy()
    if uni.empty:
        return pd.DataFrame()

    uni["candidate_rows_for_text_issuer_event"] = (
        uni.groupby(["offering_id", "source", "source_key", "issuer"])["source_key"].transform("size")
    )
    uni["match_origin_priority"] = (uni["match_origin"] != "explicit").astype(int)
    uni["entity_level_priority"] = (uni["entity_level"] != "offering").astype(int)
    uni["has_matched_offering"] = uni["matched_offering_id"].notna().astype(int)
    uni = uni.sort_values(
        [
            "offering_id", "source", "source_key", "issuer",
            "match_origin_priority", "entity_level_priority", "has_matched_offering",
        ],
        ascending=[True, True, True, True, True, True, False],
    )
    uni = uni.drop_duplicates(subset=["offering_id", "source", "source_key", "issuer"]).copy()
    uni = uni.drop(columns=["match_origin_priority", "entity_level_priority", "has_matched_offering"])

    uni["mention_uid"] = (
        "event:" + uni["offering_id"].astype("Int64").astype(str)
        + "|" + uni["source"].astype(str)
        + "|" + uni["source_key"].astype(str)
        + "|issuer:" + uni["issuer"].astype(str)
    )

    # Helpful event-time fields
    sec_to_end = (uni["event_end"] - uni["published_at"]).dt.total_seconds()
    sec_from_start = (uni["published_at"] - uni["event_start"]).dt.total_seconds()

    uni["seconds_to_event_end"] = sec_to_end
    uni["hours_to_event_end"] = sec_to_end / 3600.0
    uni["days_to_event_end"] = sec_to_end / 86400.0

    uni["seconds_from_event_start"] = sec_from_start
    uni["hours_from_event_start"] = sec_from_start / 3600.0
    uni["days_from_event_start"] = sec_from_start / 86400.0

    # Source-specific buzz flags (useful later in aggregation)
    uni["is_tg"] = (uni["source"] == "telegram").astype(int)
    uni["is_sl_post"] = (uni["source"] == "smartlab_post").astype(int)
    uni["is_sl_comment"] = (uni["source"] == "smartlab_comment").astype(int)

    keep_cols = [
        "mention_uid", "source", "source_key", "channel", "local_id", "parent_id",
        "published_at", "issuer", "entity_level", "offering_id", "bond_name", "ISIN",
        "matched_offering_id", "matched_bond_name", "matched_isin",
        "book_date", "placement_date", "marketing_start", "pricing_date",
        "bookbuilding_open_dt", "bookbuilding_close_time", "bookbuilding_close_dt",
        "bookbuilding_exact_time_available", "event_end_source",
        "event_start", "event_end",
        "seconds_from_event_start", "hours_from_event_start", "days_from_event_start",
        "seconds_to_event_end", "hours_to_event_end", "days_to_event_end",
        "match_origin", "inherit_anchor_key", "text_len",
        "candidate_rows_for_text_issuer_event",
        "matched_aliases", "alias_types",
        "is_tg", "is_sl_post", "is_sl_comment",
        "text", "parent_text",
    ]
    keep_cols = [c for c in keep_cols if c in uni.columns]
    return uni[keep_cols].copy()


# ---------------------------------------------------------
# Optional event-level pre-sentiment features
# ---------------------------------------------------------

def build_pre_sentiment_event_features(universe: pd.DataFrame) -> pd.DataFrame:
    if universe.empty:
        return pd.DataFrame()

    agg = (
        universe.groupby(["offering_id", "issuer", "bond_name", "ISIN", "event_start", "event_end"], dropna=False)
        .agg(
            n_texts=("mention_uid", "nunique"),
            n_source_texts=("source_key", "nunique"),
            n_tg=("is_tg", "sum"),
            n_sl_post=("is_sl_post", "sum"),
            n_sl_comment=("is_sl_comment", "sum"),
            share_inherited=("match_origin", lambda s: (s == "inherited").mean()),
            share_offering_level=("entity_level", lambda s: (s == "offering").mean()),
            mean_candidate_rows_per_text=("candidate_rows_for_text_issuer_event", "mean"),
        )
        .reset_index()
    )
    agg["log_buzz"] = np.log1p(agg["n_texts"])
    return agg


# ---------------------------------------------------------
# Main
# ---------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create an offering-level text universe for building primary-market sentiment indices."
    )
    parser.add_argument("--panel", type=Path, default=Path("panel_final_v3.xlsx"))
    parser.add_argument("--telegram-db", type=Path, default=Path("data.db"))
    parser.add_argument("--smartlab-db", type=Path, default=Path("smartlab_bonds.db"))
    parser.add_argument("--out-dir", type=Path, default=Path("primary_sentiment_ready"))
    parser.add_argument("--window-days", type=int, default=14,
                        help="Backward window length before event_end. Final window is [event_end - window_days, event_end).")
    parser.add_argument("--use-marketing-window", action="store_true",
                        help="Use marketing_start -> event_end when marketing_start exists; otherwise fallback to [-window_days, 0].")
    args = parser.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)

    panel = load_panel(args.panel)
    panel_windows = build_event_windows(panel, window_days=args.window_days, use_marketing_window=args.use_marketing_window)

    mentions = load_all_mentions(args.telegram_db, args.smartlab_db)
    universe = expand_to_primary_event_universe(mentions, panel_windows)

    universe_path = args.out_dir / "sentiment_universe_primary_v10.csv"
    universe.to_csv(universe_path, index=False, encoding="utf-8-sig")

    # Helpful auxiliary file before sentiment scoring
    features = build_pre_sentiment_event_features(universe)
    features_path = args.out_dir / "event_features_pre_sentiment_v10.csv"
    features.to_csv(features_path, index=False, encoding="utf-8-sig")

    print(f"Saved: {universe_path}")
    print(f"Saved: {features_path}")
    print(f"Rows in universe: {len(universe):,}")
    if not universe.empty:
        print(f"Unique offerings covered: {universe['offering_id'].nunique():,}")
        print(f"Unique texts: {universe['mention_uid'].nunique():,}")


if __name__ == "__main__":
    main()
