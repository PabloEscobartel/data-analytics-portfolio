#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Post-classification pipeline: classified.csv → regression-ready dataset.

Replaces three separate scripts:
  1. add_engagement_metrics_to_universe.py
  2. aggregate_sentiment_offering.py
  3. build_regression_dataset.py

Input:
  - classified.csv          (output of train_bond_classifier.py --predict)
  - data.db                 (Telegram messages with views/reactions)
  - smartlab_bonds.db       (Smart-Lab posts/comments with views/votes)
  - panel_final_v2.xlsx     (offering panel with control variables)

Output (in --output-dir):
  - message_level.csv           all messages with engagement + panel info
  - aggregated_by_offering.csv  offering-level sentiment metrics
  - regression_ready.csv        panel + sentiment, ready for regression
  - regression_ready.xlsx       same as Excel
  - pipeline_summary.txt        stats

Usage:
  python build_pipeline.py \
    --classified classified.csv \
    --panel panel_final_v2.xlsx \
    --message-cutoff book_open \
    --lookback-days 14
"""

from __future__ import annotations

import argparse
import math
import sqlite3
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd


# ============================================================================
#  Step 1: Load and parse classified.csv
# ============================================================================

def load_classified(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, low_memory=False)
    print(f"Classified: {len(df):,} rows")

    # Parse source_key → source-specific IDs
    # Telegram: "channel_name:12345"
    # SL post:  "post:678"
    # SL comment: "comment:901"
    df["channel"] = None
    df["local_id"] = None

    tg_mask = df["source"] == "telegram"
    if tg_mask.any():
        parts = df.loc[tg_mask, "source_key"].str.split(":", n=1, expand=True)
        df.loc[tg_mask, "channel"] = parts[0]
        df.loc[tg_mask, "local_id"] = parts[1]

    slp_mask = df["source"] == "smartlab_post"
    if slp_mask.any():
        df.loc[slp_mask, "local_id"] = df.loc[slp_mask, "source_key"].str.replace("post:", "", regex=False)

    slc_mask = df["source"] == "smartlab_comment"
    if slc_mask.any():
        df.loc[slc_mask, "local_id"] = df.loc[slc_mask, "source_key"].str.replace("comment:", "", regex=False)

    # Source flags
    df["is_tg"] = (df["source"] == "telegram").astype(int)
    df["is_sl_post"] = (df["source"] == "smartlab_post").astype(int)
    df["is_sl_comment"] = (df["source"] == "smartlab_comment").astype(int)

    return df


# ============================================================================
#  Step 2: Add published_at + engagement from databases
# ============================================================================

def add_telegram_metadata(df: pd.DataFrame, db_path: Path) -> pd.DataFrame:
    tg = df[df["source"] == "telegram"].copy()
    if tg.empty or not db_path.exists():
        return tg

    conn = sqlite3.connect(str(db_path))
    msg = pd.read_sql_query("""
        SELECT channel, CAST(id AS TEXT) AS local_id,
               date AS published_at, views, reactions, reply_count
        FROM messages
    """, conn)
    conn.close()

    tg = tg.merge(msg, on=["channel", "local_id"], how="left", suffixes=("", "_db"))

    # Use DB published_at if not already present
    if "published_at" not in df.columns or tg["published_at"].isna().all():
        pass  # already merged as published_at
    elif "published_at_db" in tg.columns:
        tg["published_at"] = tg["published_at_db"].fillna(tg.get("published_at", ""))

    tg["platform_views"] = pd.to_numeric(tg.get("views"), errors="coerce")
    tg["platform_reactions"] = pd.to_numeric(tg.get("reactions"), errors="coerce")
    tg["platform_replies"] = pd.to_numeric(tg.get("reply_count"), errors="coerce")
    tg["engagement_total"] = tg["platform_reactions"].fillna(0) + tg["platform_replies"].fillna(0)

    return tg


def add_smartlab_post_metadata(df: pd.DataFrame, db_path: Path) -> pd.DataFrame:
    sp = df[df["source"] == "smartlab_post"].copy()
    if sp.empty or not db_path.exists():
        return sp

    conn = sqlite3.connect(str(db_path))
    posts = pd.read_sql_query("""
        SELECT CAST(post_id AS TEXT) AS local_id,
               published_at, views, votes, comments_count
        FROM posts
    """, conn)
    conn.close()

    sp = sp.merge(posts, on="local_id", how="left", suffixes=("", "_db"))
    if "published_at_db" in sp.columns:
        sp["published_at"] = sp["published_at_db"].fillna(sp.get("published_at", ""))

    sp["platform_views"] = pd.to_numeric(sp.get("views"), errors="coerce")
    sp["platform_reactions"] = pd.to_numeric(sp.get("votes"), errors="coerce")
    sp["platform_replies"] = pd.to_numeric(sp.get("comments_count"), errors="coerce")
    sp["engagement_total"] = sp["platform_reactions"].fillna(0) + sp["platform_replies"].fillna(0)

    return sp


def add_smartlab_comment_metadata(df: pd.DataFrame, db_path: Path) -> pd.DataFrame:
    sc = df[df["source"] == "smartlab_comment"].copy()
    if sc.empty or not db_path.exists():
        return sc

    conn = sqlite3.connect(str(db_path))
    comments = pd.read_sql_query("""
        SELECT CAST(comment_id AS TEXT) AS local_id,
               published_at, votes
        FROM comments
    """, conn)
    conn.close()

    sc = sc.merge(comments, on="local_id", how="left", suffixes=("", "_db"))
    if "published_at_db" in sc.columns:
        sc["published_at"] = sc["published_at_db"].fillna(sc.get("published_at", ""))

    sc["platform_views"] = np.nan
    sc["platform_reactions"] = pd.to_numeric(sc.get("votes"), errors="coerce")
    sc["platform_replies"] = np.nan
    sc["engagement_total"] = sc["platform_reactions"].fillna(0)

    return sc


def enrich_with_metadata(df: pd.DataFrame, tg_db: Path, sl_db: Path) -> pd.DataFrame:
    parts = [
        add_telegram_metadata(df, tg_db),
        add_smartlab_post_metadata(df, sl_db),
        add_smartlab_comment_metadata(df, sl_db),
    ]
    out = pd.concat([p for p in parts if not p.empty], ignore_index=True, sort=False)
    out["published_at"] = pd.to_datetime(out["published_at"], errors="coerce", utc=True)
    print(f"Enriched: {len(out):,} rows with metadata")
    return out


# ============================================================================
#  Step 3: Add panel info (book_date, ISIN, etc.)
# ============================================================================

def add_panel_info(df: pd.DataFrame, panel: pd.DataFrame) -> pd.DataFrame:
    """Join offering_id → panel columns for time-based features."""
    panel = panel.copy()
    panel["offering_id"] = range(1, len(panel) + 1)
    panel["book_date"] = pd.to_datetime(panel["book_date"], errors="coerce", utc=True)
    panel["placement_date"] = pd.to_datetime(panel["placement_date"], errors="coerce", utc=True)
    if "cbonds_book_start_datetime" in panel.columns:
        panel["book_open_datetime"] = pd.to_datetime(panel["cbonds_book_start_datetime"], errors="coerce", utc=True)
    else:
        panel["book_open_datetime"] = panel["book_date"]
    if "cbonds_book_end_datetime" in panel.columns:
        panel["book_close_datetime"] = pd.to_datetime(panel["cbonds_book_end_datetime"], errors="coerce", utc=True)
    else:
        panel["book_close_datetime"] = panel["book_date"]

    panel_cols = panel[[
        "offering_id", "issuer", "bond_name", "ISIN", "book_date",
        "book_open_datetime", "book_close_datetime", "placement_date",
    ]].copy()
    panel_cols = panel_cols.rename(columns={
        "issuer": "panel_issuer", "bond_name": "panel_bond_name",
    })

    df["offering_id"] = pd.to_numeric(df["offering_id"], errors="coerce")
    merged = df.merge(panel_cols, on="offering_id", how="left")
    merged["message_scope"] = np.where(merged["offering_id"].notna(), "offering", "issuer")

    # For issuer-level matches (no offering_id), try joining on issuer name
    no_oid = merged["ISIN"].isna() & merged["offering_id"].isna()
    if no_oid.any():
        print(f"  {no_oid.sum():,} rows without offering_id (issuer-level matches)")

    # Time features
    if "published_at" in merged.columns and "book_date" in merged.columns:
        merged["days_to_book"] = (
            merged["book_date"] - merged["published_at"]
        ).dt.total_seconds() / 86400
    if "published_at" in merged.columns and "book_open_datetime" in merged.columns:
        merged["days_to_book_open"] = (
            merged["book_open_datetime"] - merged["published_at"]
        ).dt.total_seconds() / 86400
    if "published_at" in merged.columns and "book_close_datetime" in merged.columns:
        merged["days_to_book_close"] = (
            merged["book_close_datetime"] - merged["published_at"]
        ).dt.total_seconds() / 86400

    return merged


def issuer_key(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.casefold()


def message_window_mask(df: pd.DataFrame, cutoff_col: str, lookback_days: float) -> pd.Series:
    """Return mask for messages in (cutoff - lookback_days; cutoff)."""
    cutoff_ts = df[cutoff_col]
    lower_ts = cutoff_ts - pd.to_timedelta(lookback_days, unit="D")
    return (
        df["published_at"].notna()
        & cutoff_ts.notna()
        & (df["published_at"] > lower_ts)
        & (df["published_at"] < cutoff_ts)
    )


def apply_message_cutoff(df: pd.DataFrame, cutoff: str, lookback_days: float) -> pd.DataFrame:
    """Keep only offering-level messages in (selected book timestamp - lookback; timestamp)."""
    if cutoff == "none":
        print("Temporal filter: none")
        return df

    cutoff_col = {
        "book_open": "book_open_datetime",
        "book_close": "book_close_datetime",
    }[cutoff]
    if cutoff_col not in df.columns:
        raise KeyError(f"Не найдена колонка для временного фильтра: {cutoff_col}")

    before = len(df)
    offering_mask = df["offering_id"].notna()
    keep_offering = message_window_mask(df, cutoff_col, lookback_days)
    keep = ~offering_mask | keep_offering
    out = df.loc[keep].copy()

    before_offering = int(offering_mask.sum())
    after_offering = int((out["offering_id"].notna()).sum())
    print(
        f"Temporal filter: {cutoff} ({cutoff_col}), window=(-{lookback_days:g}d, cutoff); "
        f"offering-level messages {before_offering:,} → {after_offering:,}; "
        f"all rows {before:,} → {len(out):,}"
    )
    return out


# ============================================================================
#  Step 4: Aggregate to offering level
# ============================================================================

def safe_mean(s):
    s = pd.to_numeric(s, errors="coerce").dropna()
    return float(s.mean()) if len(s) else np.nan

def safe_std(s):
    s = pd.to_numeric(s, errors="coerce").dropna()
    return float(s.std(ddof=1)) if len(s) > 1 else np.nan

def weighted_mean(values, weights):
    v = pd.to_numeric(values, errors="coerce")
    w = pd.to_numeric(weights, errors="coerce")
    mask = v.notna() & w.notna() & (w > 0)
    if mask.sum() == 0:
        return np.nan
    return float(np.average(v[mask], weights=w[mask]))

def first_valid(s):
    s = s.dropna()
    return s.iloc[0] if len(s) else np.nan


def aggregate_offering(g: pd.DataFrame) -> dict:
    """Aggregate one offering's messages into summary metrics."""
    rel = g[g["relevance_pred"] == "yes"]

    n_all = len(g)
    n_rel = len(rel)

    # Sentiment counts (from discrete labels)
    n_pos = int((rel["sentiment_pred"] == "positive").sum())
    n_neg = int((rel["sentiment_pred"] == "negative").sum())
    n_neu = int((rel["sentiment_pred"] == "neutral").sum())

    # Continuous sentiment score
    scores = pd.to_numeric(rel["sentiment_score"], errors="coerce").dropna()

    out = {
        # Identifiers
        "panel_issuer": first_valid(g["panel_issuer"]) if "panel_issuer" in g.columns else np.nan,
        "panel_bond_name": first_valid(g["panel_bond_name"]) if "panel_bond_name" in g.columns else np.nan,
        "ISIN": first_valid(g["ISIN"]) if "ISIN" in g.columns else np.nan,
        "book_date": first_valid(g["book_date"]) if "book_date" in g.columns else np.nan,
        "book_open_datetime": first_valid(g["book_open_datetime"]) if "book_open_datetime" in g.columns else np.nan,
        "book_close_datetime": first_valid(g["book_close_datetime"]) if "book_close_datetime" in g.columns else np.nan,
        "placement_date": first_valid(g["placement_date"]) if "placement_date" in g.columns else np.nan,

        # Coverage
        "n_messages_all": n_all,
        "n_messages_relevant": n_rel,
        "log_buzz_all": math.log1p(n_all),
        "log_buzz_relevant": math.log1p(n_rel),
        "share_relevant": n_rel / n_all if n_all > 0 else np.nan,
        "n_sources": int(g["source"].nunique()) if "source" in g.columns else np.nan,

        # Discrete sentiment counts
        "n_positive": n_pos,
        "n_negative": n_neg,
        "n_neutral": n_neu,
        "share_positive": n_pos / n_rel if n_rel > 0 else np.nan,
        "share_negative": n_neg / n_rel if n_rel > 0 else np.nan,

        # Index metrics (from discrete labels)
        "si": (n_pos - n_neg) / n_rel if n_rel > 0 else np.nan,
        "ndi": (n_pos - n_neg) / (n_pos + n_neg) if (n_pos + n_neg) > 0 else np.nan,

        # Continuous sentiment (from P(pos) - P(neg))
        "mean_sentiment_score": float(scores.mean()) if len(scores) > 0 else np.nan,
        "median_sentiment_score": float(scores.median()) if len(scores) > 0 else np.nan,
        "std_sentiment_score": float(scores.std(ddof=1)) if len(scores) > 1 else np.nan,

        # Probability-based metrics
        "mean_p_positive": safe_mean(rel["p_positive"]) if "p_positive" in rel.columns else np.nan,
        "mean_p_negative": safe_mean(rel["p_negative"]) if "p_negative" in rel.columns else np.nan,
        "mean_p_neutral": safe_mean(rel["p_neutral"]) if "p_neutral" in rel.columns else np.nan,

        # Platform composition
        "n_tg": int(g["is_tg"].sum()) if "is_tg" in g.columns else 0,
        "n_sl_post": int(g["is_sl_post"].sum()) if "is_sl_post" in g.columns else 0,
        "n_sl_comment": int(g["is_sl_comment"].sum()) if "is_sl_comment" in g.columns else 0,

        # Engagement
        "views_sum": float(g["platform_views"].fillna(0).sum()) if "platform_views" in g.columns else np.nan,
        "views_mean": safe_mean(g["platform_views"]) if "platform_views" in g.columns else np.nan,
        "reactions_sum": float(g["platform_reactions"].fillna(0).sum()) if "platform_reactions" in g.columns else np.nan,
        "engagement_sum": float(g["engagement_total"].fillna(0).sum()) if "engagement_total" in g.columns else np.nan,
    }

    # Weighted sentiment
    if "platform_views" in rel.columns and "sentiment_score" in rel.columns:
        out["sentiment_view_weighted"] = weighted_mean(rel["sentiment_score"], rel["platform_views"])
    if "engagement_total" in rel.columns and "sentiment_score" in rel.columns:
        out["sentiment_engagement_weighted"] = weighted_mean(rel["sentiment_score"], rel["engagement_total"])

    # Timing
    if "days_to_book" in g.columns:
        out["mean_days_to_book"] = safe_mean(g["days_to_book"])
    if "days_to_book_open" in g.columns:
        out["mean_days_to_book_open"] = safe_mean(g["days_to_book_open"])
    if "days_to_book_close" in g.columns:
        out["mean_days_to_book_close"] = safe_mean(g["days_to_book_close"])

    return out


def build_panel_meta(panel: pd.DataFrame) -> pd.DataFrame:
    panel_meta = panel.copy()
    panel_meta["offering_id"] = range(1, len(panel_meta) + 1)
    panel_meta["book_date"] = pd.to_datetime(panel_meta["book_date"], errors="coerce", utc=True)
    panel_meta["placement_date"] = pd.to_datetime(panel_meta["placement_date"], errors="coerce", utc=True)
    if "cbonds_book_start_datetime" in panel_meta.columns:
        panel_meta["book_open_datetime"] = pd.to_datetime(panel_meta["cbonds_book_start_datetime"], errors="coerce", utc=True)
    else:
        panel_meta["book_open_datetime"] = panel_meta["book_date"]
    if "cbonds_book_end_datetime" in panel_meta.columns:
        panel_meta["book_close_datetime"] = pd.to_datetime(panel_meta["cbonds_book_end_datetime"], errors="coerce", utc=True)
    else:
        panel_meta["book_close_datetime"] = panel_meta["book_date"]
    panel_meta["issuer_key"] = issuer_key(panel_meta["issuer"])
    return panel_meta[[
        "offering_id", "issuer_key", "issuer", "bond_name", "ISIN", "book_date",
        "book_open_datetime", "book_close_datetime", "placement_date",
    ]].rename(columns={"issuer": "panel_issuer", "bond_name": "panel_bond_name"})


def assign_issuer_level_messages(
    df: pd.DataFrame,
    panel: pd.DataFrame,
    message_cutoff: str,
    lookback_days: float,
) -> pd.DataFrame:
    """Assign issuer-level messages to same-issuer offerings inside each offering's time window."""
    issuer_rows = df[df["offering_id"].isna() & df["issuer"].notna()].copy()
    if issuer_rows.empty:
        return issuer_rows

    needed_cols = [c for c in [
        "source_key", "source", "channel", "local_id", "published_at", "issuer",
        "entity_level", "bond_name", "match_origin", "relevance_pred", "relevance_conf",
        "sentiment_pred", "sentiment_score", "p_negative", "p_neutral", "p_positive",
        "platform_views", "platform_reactions", "engagement_total",
        "is_tg", "is_sl_post", "is_sl_comment",
    ] if c in issuer_rows.columns]
    issuer_rows = issuer_rows[needed_cols].copy()
    issuer_rows["issuer_key"] = issuer_key(issuer_rows["issuer"])

    assigned = issuer_rows.merge(build_panel_meta(panel), on="issuer_key", how="inner")
    assigned["message_scope"] = "issuer"

    if message_cutoff != "none":
        cutoff_col = {
            "book_open": "book_open_datetime",
            "book_close": "book_close_datetime",
        }[message_cutoff]
        assigned = assigned.loc[message_window_mask(assigned, cutoff_col, lookback_days)].copy()

    if "published_at" in assigned.columns:
        assigned["days_to_book"] = (
            assigned["book_date"] - assigned["published_at"]
        ).dt.total_seconds() / 86400
        assigned["days_to_book_open"] = (
            assigned["book_open_datetime"] - assigned["published_at"]
        ).dt.total_seconds() / 86400
        assigned["days_to_book_close"] = (
            assigned["book_close_datetime"] - assigned["published_at"]
        ).dt.total_seconds() / 86400

    return assigned.drop(columns=["issuer_key"], errors="ignore")


def aggregate_to_offerings(
    df: pd.DataFrame,
    panel: pd.DataFrame,
    min_messages: int = 3,
    message_cutoff: str = "book_open",
    lookback_days: float = 14,
    include_issuer_level: bool = True,
    min_messages_field: str = "all",
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Aggregate direct offering-level and optionally issuer-level messages to offering-level."""
    work_parts = []

    offering_work = df[df["offering_id"].notna()].copy()
    if not offering_work.empty:
        offering_work["offering_id"] = offering_work["offering_id"].astype(int)
        offering_work["message_scope"] = "offering"
        work_parts.append(offering_work)

    issuer_assigned_n = 0
    if include_issuer_level:
        issuer_work = assign_issuer_level_messages(df, panel, message_cutoff, lookback_days)
        issuer_assigned_n = len(issuer_work)
        if not issuer_work.empty:
            issuer_work["offering_id"] = issuer_work["offering_id"].astype(int)
            work_parts.append(issuer_work)

    if not work_parts:
        empty = pd.DataFrame(columns=["offering_id"])
        return empty, empty.copy(), {
            "offering_level_messages": 0,
            "issuer_level_assignments": 0,
            "deduped_messages": 0,
        }

    work = pd.concat(work_parts, ignore_index=True, sort=False)
    before_dedupe = len(work)
    if "source_key" in work.columns:
        work["_scope_priority"] = np.where(work["message_scope"].eq("offering"), 0, 1)
        work = (
            work.sort_values(["offering_id", "source_key", "_scope_priority"])
            .drop_duplicates(["offering_id", "source_key"], keep="first")
            .drop(columns=["_scope_priority"])
        )

    grouped = (
        work.groupby("offering_id")
        .apply(aggregate_offering, include_groups=False)
        .apply(pd.Series)
        .reset_index()
    )
    grouped = grouped.sort_values("n_messages_all", ascending=False).reset_index(drop=True)

    full = grouped.copy()
    min_col = "n_messages_relevant" if min_messages_field == "relevant" else "n_messages_all"
    filtered = grouped[grouped[min_col] >= min_messages].copy()

    stats = {
        "offering_level_messages": int(len(offering_work)),
        "issuer_level_assignments": int(issuer_assigned_n),
        "deduped_messages": int(before_dedupe - len(work)),
        "aggregation_message_rows": int(len(work)),
        "min_messages_field": min_messages_field,
        "lookback_days": float(lookback_days),
    }
    issuer_note = f", issuer assignments={issuer_assigned_n:,}" if include_issuer_level else ", issuer-level excluded"
    print(
        f"Aggregated: {len(full):,} offerings total, {len(filtered):,} with >= {min_messages} {min_col}"
        f" (message rows={len(work):,}{issuer_note}, deduped={stats['deduped_messages']:,})"
    )
    return full, filtered, stats


# ============================================================================
#  Step 5: Build regression dataset
# ============================================================================

def build_regression(aggregated: pd.DataFrame, panel: pd.DataFrame) -> pd.DataFrame:
    """Merge aggregated sentiment with panel control variables."""
    panel = panel.copy()
    panel = panel.rename(columns={
        "tizer_value": "teaser_value",
        "tizer_volume": "teaser_volume",
    })
    panel["offering_id"] = range(1, len(panel) + 1)

    for col in ["book_date", "placement_date"]:
        panel[col] = pd.to_datetime(panel[col], errors="coerce")

    merged = panel.merge(aggregated, on="offering_id", how="left", suffixes=("", "_sent"))
    sentiment_rows = int(merged["n_messages_all"].notna().sum()) if "n_messages_all" in merged.columns else 0

    # Dependent variable helpers
    if "coupon_reduction_bp" in merged.columns:
        merged["coupon_reduction_bp"] = pd.to_numeric(merged["coupon_reduction_bp"], errors="coerce")
        merged["coupon_reduced_dummy"] = (merged["coupon_reduction_bp"] > 0).astype("Int64")
    if "volume_ratio" in merged.columns:
        merged["volume_ratio"] = pd.to_numeric(merged["volume_ratio"], errors="coerce")
        merged["volume_increased_dummy"] = (merged["volume_ratio"] > 1).astype("Int64")
        merged["log_volume_ratio"] = np.where(merged["volume_ratio"] > 0, np.log(merged["volume_ratio"]), np.nan)

    print(f"Regression dataset: {len(merged):,} rows ({sentiment_rows:,} with sentiment)")
    return merged


def remove_excel_timezones(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with timezone-aware datetimes converted to Excel-safe values."""
    out = df.copy()

    for col in out.columns:
        series = out[col]
        if isinstance(series.dtype, pd.DatetimeTZDtype):
            out[col] = series.dt.tz_convert(None)
        elif series.dtype == "object":
            out[col] = series.map(remove_timezone_from_value)

    return out


def remove_timezone_from_value(value):
    if isinstance(value, pd.Timestamp) and value.tzinfo is not None:
        return value.tz_convert(None)
    if isinstance(value, datetime) and value.tzinfo is not None:
        return value.replace(tzinfo=None)
    return value


# ============================================================================
#  Main
# ============================================================================

def main():
    p = argparse.ArgumentParser(description="Post-classification pipeline → regression dataset")
    p.add_argument("--classified", type=Path, default=Path("classified.csv"))
    p.add_argument("--panel", type=Path, default=Path("panel_final_v2.xlsx"))
    p.add_argument("--telegram-db", type=Path, default=Path("data.db"))
    p.add_argument("--smartlab-db", type=Path, default=Path("smartlab_bonds.db"))
    p.add_argument("--output-dir", type=Path, default=Path("regression_dataset"))
    p.add_argument("--min-messages", type=int, default=3)
    p.add_argument(
        "--lookback-days",
        type=float,
        default=14,
        help="Length of the pre-book message window. Uses strict interval (cutoff - days; cutoff).",
    )
    p.add_argument(
        "--min-messages-field",
        choices=["all", "relevant"],
        default="relevant",
        help="Which offering-level message count is used for the min-messages filter.",
    )
    p.add_argument(
        "--message-cutoff",
        choices=["book_open", "book_close", "none"],
        default="book_open",
        help="Temporal cutoff for offering-level messages before aggregation.",
    )
    p.add_argument(
        "--issuer-level-mode",
        choices=["include", "exclude"],
        default="include",
        help="Include issuer-level messages by assigning them to same-issuer offerings before each offering cutoff.",
    )
    args = p.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)

    # Load
    panel = pd.read_excel(args.panel)
    df = load_classified(args.classified)

    # Enrich with published_at + engagement
    df = enrich_with_metadata(df, args.telegram_db, args.smartlab_db)

    # Add panel info (ISIN, book_date, etc.)
    df = add_panel_info(df, panel)
    df = apply_message_cutoff(df, args.message_cutoff, args.lookback_days)

    # Save message-level
    msg_path = args.output_dir / "message_level.csv"
    msg_cols = [c for c in [
        "source_key", "source", "channel", "local_id", "published_at",
        "issuer", "entity_level", "offering_id", "bond_name", "ISIN",
        "match_origin", "relevance_pred", "relevance_conf",
        "sentiment_pred", "sentiment_score", "p_negative", "p_neutral", "p_positive",
        "platform_views", "platform_reactions", "engagement_total",
        "book_date", "book_open_datetime", "book_close_datetime", "placement_date",
        "days_to_book", "days_to_book_open", "days_to_book_close",
    ] if c in df.columns]
    df[msg_cols].to_csv(msg_path, index=False, encoding="utf-8-sig")
    print(f"Saved: {msg_path}")

    # Aggregate
    full_agg, filtered_agg, agg_stats = aggregate_to_offerings(
        df,
        panel,
        min_messages=args.min_messages,
        message_cutoff=args.message_cutoff,
        lookback_days=args.lookback_days,
        include_issuer_level=args.issuer_level_mode == "include",
        min_messages_field=args.min_messages_field,
    )
    full_agg.to_csv(args.output_dir / "aggregated_all.csv", index=False, encoding="utf-8-sig")
    filtered_agg.to_csv(args.output_dir / "aggregated_filtered.csv", index=False, encoding="utf-8-sig")

    # Regression dataset
    regression = build_regression(filtered_agg, panel)
    reg_csv = args.output_dir / "regression_ready.csv"
    reg_xlsx = args.output_dir / "regression_ready.xlsx"
    regression.to_csv(reg_csv, index=False, encoding="utf-8-sig")
    remove_excel_timezones(regression).to_excel(reg_xlsx, index=False)

    # Summary
    summary = args.output_dir / "pipeline_summary.txt"
    with open(summary, "w") as f:
        f.write(f"Classified input:    {len(load_classified.__wrapped__ if hasattr(load_classified, '__wrapped__') else pd.read_csv(args.classified)):,} rows\n" if False else "")
        f.write(f"Panel offerings:     {len(panel):,}\n")
        f.write(f"Enriched messages:   {len(df):,}\n")
        f.write(f"Relevant messages:   {(df.get('relevance_pred','') == 'yes').sum():,}\n")
        f.write(f"Aggregated offerings:{len(full_agg):,}\n")
        f.write(f"  with >={args.min_messages} messages: {len(filtered_agg):,}\n")
        f.write(f"Regression rows:     {len(regression):,}\n")
        sentiment_rows = int(regression["n_messages_all"].notna().sum()) if "n_messages_all" in regression.columns else 0
        f.write(f"  with sentiment:    {sentiment_rows:,}\n")
        f.write(f"Message cutoff:      {args.message_cutoff}\n")
        f.write(f"Lookback days:       {args.lookback_days:g}\n")
        f.write(f"Issuer-level mode:   {args.issuer_level_mode}\n")
        f.write(f"Min messages field:  {args.min_messages_field}\n")
        f.write(f"Offering msg rows:   {agg_stats['offering_level_messages']:,}\n")
        f.write(f"Issuer assignments:  {agg_stats['issuer_level_assignments']:,}\n")
        f.write(f"Aggregation rows:    {agg_stats['aggregation_message_rows']:,}\n")
        f.write(f"Deduped rows:        {agg_stats['deduped_messages']:,}\n")

    n_rel = int((df.get("relevance_pred", "") == "yes").sum())
    print(f"\n{'='*60}")
    print(f"Pipeline complete")
    print(f"{'='*60}")
    print(f"Messages:    {len(df):,}  (relevant: {n_rel:,})")
    print(f"Offerings:   {len(filtered_agg):,}  (with >={args.min_messages} messages)")
    sentiment_rows = int(regression["n_messages_all"].notna().sum()) if "n_messages_all" in regression.columns else 0
    print(f"Regression:  {len(regression):,}  ({sentiment_rows:,} with sentiment)")
    print(f"Cutoff:      {args.message_cutoff}")
    print(f"Lookback:    {args.lookback_days:g} days")
    print(f"Issuer mode: {args.issuer_level_mode}")
    print(f"Min field:   {args.min_messages_field}")
    print(f"Output:      {args.output_dir}/")


if __name__ == "__main__":
    main()
