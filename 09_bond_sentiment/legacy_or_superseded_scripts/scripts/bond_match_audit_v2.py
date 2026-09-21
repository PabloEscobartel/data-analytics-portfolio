#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Audit resolved bond-entity matches and produce a short QA report.

Compatible with bond_entity_matcher v8 (topical context columns).

Inputs
------
- Telegram DB (default: data.db)
- Smart-Lab DB (default: smartlab_bonds.db)
- Coverage file from bond_entity_matcher (xlsx or csv)

Outputs
-------
Creates an output directory with:
- audit_report.txt                Human-readable summary with topical distribution
- coverage_summary.csv            Coverage distribution (three layers if v8)
- top_issuers.csv                 Top issuers with bond/primary/noise breakdown
- top_offerings.csv               Top offerings by matched texts
- telegram_inherited_flags.csv    Issuers/channels where inherited may dominate
- alias_ambiguity.csv             Aliases used across many issuers/offers
- samples_tg_explicit.csv         Random sample with topical flags
- samples_tg_inherited.csv        Random sample with topical flags
- samples_sl_posts.csv            Random sample with topical flags
- samples_sl_comments.csv         Random sample with topical flags

Purpose
-------
This script does not estimate sentiment. It only checks whether the linkage
(text -> issuer/offering) looks sane enough to proceed, and whether the
topical context filtering (bond/primary/broker_noise) is working correctly.
"""

from __future__ import annotations

import argparse
import json
import os
import sqlite3
from pathlib import Path
from typing import Iterable, Optional

import numpy as np
import pandas as pd


# -----------------------------
# Helpers
# -----------------------------

def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def read_sql(conn: sqlite3.Connection, query: str, params: Optional[Iterable] = None) -> pd.DataFrame:
    return pd.read_sql_query(query, conn, params=params or ())


def table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    q = "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?"
    row = conn.execute(q, (table_name,)).fetchone()
    return row is not None


def load_coverage(path: Optional[Path]) -> Optional[pd.DataFrame]:
    if path is None or not path.exists():
        return None
    if path.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(path)
    return pd.read_csv(path)


def safe_json_list(x) -> list:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return []
    if isinstance(x, list):
        return x
    s = str(x).strip()
    if not s:
        return []
    try:
        v = json.loads(s)
        return v if isinstance(v, list) else []
    except Exception:
        return []


def pct(num: float, den: float) -> float:
    return float(num) / float(den) if den else 0.0


def fmt_pct(x: float) -> str:
    return f"{100*x:.1f}%"


def flatten_alias_counts(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    if df.empty or "matched_aliases" not in df.columns:
        return pd.DataFrame(columns=["alias", "source", "rows", "distinct_issuers", "distinct_offerings"])
    tmp = df.copy()
    tmp["alias_list"] = tmp["matched_aliases"].apply(safe_json_list)
    tmp = tmp.explode("alias_list")
    tmp = tmp[tmp["alias_list"].notna() & (tmp["alias_list"].astype(str).str.len() > 0)].copy()
    if tmp.empty:
        return pd.DataFrame(columns=["alias", "source", "rows", "distinct_issuers", "distinct_offerings"])
    out = (
        tmp.groupby("alias_list", dropna=False)
        .agg(
            rows=("source_key", "nunique"),
            distinct_issuers=("issuer", lambda s: s.dropna().nunique()),
            distinct_offerings=("offering_id", lambda s: s.dropna().nunique()),
        )
        .reset_index()
        .rename(columns={"alias_list": "alias"})
    )
    out["source"] = source_name
    return out[["alias", "source", "rows", "distinct_issuers", "distinct_offerings"]]


# -----------------------------
# Main audit pieces
# -----------------------------

def audit_telegram(conn: sqlite3.Connection, sample_size: int, seed: int):
    required = ["telegram_matches_resolved", "messages"]
    missing = [t for t in required if not table_exists(conn, t)]
    if missing:
        return {"missing": missing}

    stats = read_sql(conn, """
        SELECT
            COUNT(*) AS n_total,
            SUM(CASE WHEN issuer IS NOT NULL THEN 1 ELSE 0 END) AS n_resolved,
            SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS n_explicit,
            SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS n_inherited,
            SUM(CASE WHEN entity_level='issuer' THEN 1 ELSE 0 END) AS n_issuer,
            SUM(CASE WHEN entity_level='offering' THEN 1 ELSE 0 END) AS n_offering,
            SUM(CASE WHEN issuer IS NULL THEN 1 ELSE 0 END) AS n_unresolved,

            -- было неправильно:
            -- SUM(CASE WHEN resolution_note='suppressed_broker_noise' THEN 1 ELSE 0 END) AS n_suppressed,

            -- правильно для Telegram:
            SUM(CASE WHEN explicit_status='suppressed' THEN 1 ELSE 0 END) AS n_suppressed,
            SUM(CASE WHEN explicit_status='suppressed' AND issuer IS NULL THEN 1 ELSE 0 END) AS n_suppressed_dropped,
            SUM(CASE WHEN explicit_status='suppressed' AND match_origin='inherited' THEN 1 ELSE 0 END) AS n_suppressed_but_inherited,

            AVG(CASE WHEN match_origin='inherited' THEN is_short_reply END) AS inherited_short_reply_share,
            AVG(CASE WHEN match_origin='inherited' THEN text_len END) AS inherited_mean_text_len,
            AVG(CASE WHEN match_origin='explicit' THEN text_len END) AS explicit_mean_text_len,
            SUM(CASE WHEN issuer IS NOT NULL AND bond_context=1 THEN 1 ELSE 0 END) AS n_bond_context,
            SUM(CASE WHEN issuer IS NOT NULL AND primary_context=1 THEN 1 ELSE 0 END) AS n_primary_context,
            SUM(CASE WHEN issuer IS NOT NULL AND broker_noise=1 THEN 1 ELSE 0 END) AS n_broker_noise
        FROM telegram_matches_resolved
    """).iloc[0].to_dict()

    top_issuers = read_sql(conn, """
        SELECT issuer,
               COUNT(*) AS matched_texts,
               SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS explicit_texts,
               SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS inherited_texts,
               SUM(CASE WHEN entity_level='offering' THEN 1 ELSE 0 END) AS offering_level,
               SUM(CASE WHEN entity_level='issuer' THEN 1 ELSE 0 END) AS issuer_level,
               SUM(CASE WHEN bond_context=1 THEN 1 ELSE 0 END) AS bond_context,
               SUM(CASE WHEN primary_context=1 THEN 1 ELSE 0 END) AS primary_context,
               SUM(CASE WHEN broker_noise=1 THEN 1 ELSE 0 END) AS broker_noise
        FROM telegram_matches_resolved
        WHERE issuer IS NOT NULL
        GROUP BY issuer
        ORDER BY matched_texts DESC
        LIMIT 100
    """)

    top_offerings = read_sql(conn, """
        SELECT COALESCE(bond_name, CAST(offering_id AS TEXT)) AS offering,
               issuer,
               offering_id,
               COUNT(*) AS matched_texts,
               SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS explicit_texts,
               SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS inherited_texts
        FROM telegram_matches_resolved
        WHERE offering_id IS NOT NULL
        GROUP BY offering_id, bond_name, issuer
        ORDER BY matched_texts DESC
        LIMIT 100
    """)

    inherited_flags = read_sql(conn, """
        SELECT issuer,
               channel,
               COUNT(*) AS matched_texts,
               SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS explicit_texts,
               SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS inherited_texts,
               ROUND(1.0 * SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END)
                     / NULLIF(COUNT(*), 0), 4) AS inherited_share,
               ROUND(AVG(CASE WHEN match_origin='inherited' THEN is_short_reply END), 4) AS inherited_short_share,
               ROUND(AVG(CASE WHEN match_origin='inherited' THEN text_len END), 1) AS inherited_mean_text_len
        FROM telegram_matches_resolved
        WHERE issuer IS NOT NULL
        GROUP BY issuer, channel
        HAVING COUNT(*) >= 20
        ORDER BY inherited_share DESC, matched_texts DESC
    """)
    if not inherited_flags.empty:
        inherited_flags["flag_inherited_dominates"] = (
            (inherited_flags["inherited_share"] >= 0.75) & (inherited_flags["matched_texts"] >= 30)
        )
        inherited_flags["flag_short_reply_heavy"] = (
            inherited_flags["inherited_short_share"].fillna(0) >= 0.55
        )

    samples_tg_explicit = read_sql(conn, f"""
        SELECT r.channel, r.message_id, r.parent_id, r.issuer, r.bond_name, r.entity_level,
               r.match_origin, r.match_score, r.resolution_note, r.explicit_status,
               r.matched_aliases, r.alias_types,
               r.bond_context, r.primary_context, r.broker_noise,
               m.text
        FROM telegram_matches_resolved r
        JOIN messages m
          ON r.channel = m.channel AND r.message_id = m.id
        WHERE r.issuer IS NOT NULL
          AND r.match_origin = 'explicit'
        ORDER BY RANDOM()
        LIMIT {int(sample_size)}
    """)

    samples_tg_inherited = read_sql(conn, f"""
        SELECT r.channel, r.message_id, r.parent_id, r.issuer, r.bond_name, r.entity_level,
               r.match_origin, r.match_score, r.resolution_note, r.inherit_depth,
               r.is_short_reply, r.explicit_status,
               r.bond_context, r.primary_context, r.broker_noise,
               m.text AS reply_text,
               p.text AS parent_text
        FROM telegram_matches_resolved r
        JOIN messages m
          ON r.channel = m.channel AND r.message_id = m.id
        LEFT JOIN messages p
          ON r.channel = p.channel AND r.parent_id = p.id
        WHERE r.issuer IS NOT NULL
          AND r.match_origin = 'inherited'
        ORDER BY RANDOM()
        LIMIT {int(sample_size)}
    """)

    alias_df = pd.DataFrame()
    if table_exists(conn, "telegram_match_candidates"):
        cand = read_sql(conn, "SELECT source_key, issuer, offering_id, matched_aliases FROM telegram_match_candidates")
        alias_df = flatten_alias_counts(cand, "telegram")

    return {
        "stats": stats,
        "top_issuers": top_issuers,
        "top_offerings": top_offerings,
        "inherited_flags": inherited_flags,
        "samples_tg_explicit": samples_tg_explicit,
        "samples_tg_inherited": samples_tg_inherited,
        "alias_counts": alias_df,
    }


def audit_smartlab(conn: sqlite3.Connection, sample_size: int):
    out = {}

    if table_exists(conn, "smartlab_post_matches_resolved") and table_exists(conn, "posts"):
        post_stats = read_sql(conn, """
            SELECT COUNT(*) AS n_total,
                   SUM(CASE WHEN issuer IS NOT NULL THEN 1 ELSE 0 END) AS n_resolved,
                   SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS n_explicit,
                   SUM(CASE WHEN entity_level='issuer' THEN 1 ELSE 0 END) AS n_issuer,
                   SUM(CASE WHEN entity_level='offering' THEN 1 ELSE 0 END) AS n_offering,
                   SUM(CASE WHEN issuer IS NULL THEN 1 ELSE 0 END) AS n_unresolved,
                   SUM(CASE WHEN resolution_note='suppressed_broker_noise' THEN 1 ELSE 0 END) AS n_suppressed,
                   SUM(CASE WHEN issuer IS NOT NULL AND bond_context=1 THEN 1 ELSE 0 END) AS n_bond_context,
                   SUM(CASE WHEN issuer IS NOT NULL AND primary_context=1 THEN 1 ELSE 0 END) AS n_primary_context
            FROM smartlab_post_matches_resolved
        """).iloc[0].to_dict()
        out["post_stats"] = post_stats
        out["samples_sl_posts"] = read_sql(conn, f"""
            SELECT r.post_id, r.issuer, r.bond_name, r.entity_level, r.match_origin,
                   r.match_score, r.resolution_note, r.matched_aliases, r.alias_types,
                   r.bond_context, r.primary_context, r.broker_noise,
                   p.title, p.tags, p.text
            FROM smartlab_post_matches_resolved r
            JOIN posts p ON r.post_id = p.post_id
            WHERE r.issuer IS NOT NULL
            ORDER BY RANDOM()
            LIMIT {int(sample_size)}
        """)
    else:
        out["post_missing"] = True

    if table_exists(conn, "smartlab_comment_matches_resolved") and table_exists(conn, "comments"):
        comment_stats = read_sql(conn, """
            SELECT COUNT(*) AS n_total,
                   SUM(CASE WHEN issuer IS NOT NULL THEN 1 ELSE 0 END) AS n_resolved,
                   SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS n_explicit,
                   SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS n_inherited,
                   SUM(CASE WHEN entity_level='issuer' THEN 1 ELSE 0 END) AS n_issuer,
                   SUM(CASE WHEN entity_level='offering' THEN 1 ELSE 0 END) AS n_offering,
                   SUM(CASE WHEN issuer IS NULL THEN 1 ELSE 0 END) AS n_unresolved,
                   SUM(CASE WHEN resolution_note='suppressed_broker_noise' THEN 1 ELSE 0 END) AS n_suppressed,
                   SUM(CASE WHEN issuer IS NOT NULL AND bond_context=1 THEN 1 ELSE 0 END) AS n_bond_context,
                   SUM(CASE WHEN issuer IS NOT NULL AND primary_context=1 THEN 1 ELSE 0 END) AS n_primary_context
            FROM smartlab_comment_matches_resolved
        """).iloc[0].to_dict()
        out["comment_stats"] = comment_stats
        out["samples_sl_comments"] = read_sql(conn, f"""
            SELECT r.comment_id, r.post_id, r.issuer, r.bond_name, r.entity_level,
                   r.match_origin, r.match_score, r.resolution_note,
                   r.bond_context, r.primary_context, r.broker_noise,
                   c.text AS comment_text,
                   p.title AS post_title
            FROM smartlab_comment_matches_resolved r
            JOIN comments c ON r.comment_id = c.comment_id
            LEFT JOIN posts p ON r.post_id = p.post_id
            WHERE r.issuer IS NOT NULL
            ORDER BY RANDOM()
            LIMIT {int(sample_size)}
        """)
    else:
        out["comment_missing"] = True

    alias_parts = []
    if table_exists(conn, "smartlab_post_match_candidates"):
        post_cand = read_sql(conn, "SELECT source_key, issuer, offering_id, matched_aliases FROM smartlab_post_match_candidates")
        alias_parts.append(flatten_alias_counts(post_cand, "smartlab_posts"))
    if table_exists(conn, "smartlab_comment_match_candidates"):
        comment_cand = read_sql(conn, "SELECT source_key, issuer, offering_id, matched_aliases FROM smartlab_comment_match_candidates")
        alias_parts.append(flatten_alias_counts(comment_cand, "smartlab_comments"))
    out["alias_counts"] = pd.concat(alias_parts, ignore_index=True) if alias_parts else pd.DataFrame()

    return out


def audit_coverage(coverage_df: Optional[pd.DataFrame]):
    if coverage_df is None or coverage_df.empty:
        return {"missing": True}

    df = coverage_df.copy()

    # v8 three-layer columns; fall back to v7 column names if present
    has_v8 = "all_entity_only" in df.columns
    if has_v8:
        layer_cols = {
            "entity_only": "all_entity_only",
            "bond": "all_bond",
            "primary": "all_primary",
        }
    else:
        # legacy v7 format — single layer
        layer_cols = {"entity_only": "all_mentions"}

    for col in layer_cols.values():
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    main_col = layer_cols.get("primary", layer_cols["entity_only"])

    summary = {
        "n_offerings": int(len(df)),
    }

    for label, col in layer_cols.items():
        if col not in df.columns:
            continue
        summary[f"n_with_{label}"] = int((df[col] > 0).sum())
        summary[f"n_zero_{label}"] = int((df[col] == 0).sum())
        summary[f"mean_{label}"] = float(df[col].mean())
        summary[f"median_{label}"] = float(df[col].median())
        summary[f"p90_{label}"] = float(df[col].quantile(0.90))
        summary[f"p95_{label}"] = float(df[col].quantile(0.95))

    dist = pd.DataFrame([
        {"metric": k, "value": v} for k, v in summary.items()
    ])

    zero_cov = df.loc[
        df[main_col] == 0,
        [c for c in ["offering_id", "issuer", "bond_name", "ISIN", "book_date", "placement_date"] + list(layer_cols.values()) if c in df.columns]
    ].copy()
    high_cov = df.sort_values(main_col, ascending=False).head(100)

    return {"summary": summary, "distribution": dist, "zero_coverage": zero_cov, "high_coverage": high_cov, "has_v8": has_v8}


# -----------------------------
# Report writing
# -----------------------------

def build_report(tg, sl, cov) -> str:
    lines = []
    lines.append("BOND MATCH AUDIT REPORT (v8 — with topical context)")
    lines.append("=" * 80)
    lines.append("")

    lines.append("1) TELEGRAM")
    lines.append("-" * 80)
    if "missing" in tg:
        lines.append(f"Missing tables: {', '.join(tg['missing'])}")
    else:
        s = tg["stats"]
        lines.append(f"Total messages: {int(s['n_total']):,}")
        lines.append(f"Resolved: {int(s['n_resolved']):,} ({fmt_pct(pct(s['n_resolved'], s['n_total']))})")
        lines.append(f"Explicit: {int(s['n_explicit']):,}")
        lines.append(f"Inherited: {int(s['n_inherited']):,}")
        lines.append(f"Offering-level: {int(s['n_offering']):,}")
        lines.append(f"Issuer-level: {int(s['n_issuer']):,}")
        lines.append(f"Unresolved: {int(s['n_unresolved']):,} ({fmt_pct(pct(s['n_unresolved'], s['n_total']))})")
        lines.append(f"Suppressed explicit: {int(s.get('n_suppressed', 0)):,}")
        lines.append(f"Suppressed and dropped finally: {int(s.get('n_suppressed_dropped', 0)):,}")
        lines.append(f"Suppressed but inherited later: {int(s.get('n_suppressed_but_inherited', 0)):,}")
        inh_share = pct(s['n_inherited'], s['n_resolved']) if s['n_resolved'] else 0
        lines.append(f"Inherited share among resolved: {fmt_pct(inh_share)}")
        if s.get("inherited_short_reply_share") is not None and not pd.isna(s.get("inherited_short_reply_share")):
            lines.append(f"Short-reply share among inherited: {fmt_pct(float(s['inherited_short_reply_share']))}")
        if s.get("inherited_mean_text_len") is not None and not pd.isna(s.get("inherited_mean_text_len")):
            lines.append(f"Mean inherited text length: {float(s['inherited_mean_text_len']):.1f}")
        if s.get("explicit_mean_text_len") is not None and not pd.isna(s.get("explicit_mean_text_len")):
            lines.append(f"Mean explicit text length: {float(s['explicit_mean_text_len']):.1f}")

        # Topical distribution
        n_res = int(s['n_resolved']) or 1
        lines.append("")
        lines.append("Topical context among resolved:")
        lines.append(f"  bond_context=1:    {int(s.get('n_bond_context', 0)):,} ({fmt_pct(pct(s.get('n_bond_context', 0), n_res))})")
        lines.append(f"  primary_context=1: {int(s.get('n_primary_context', 0)):,} ({fmt_pct(pct(s.get('n_primary_context', 0), n_res))})")
        lines.append(f"  broker_noise=1:    {int(s.get('n_broker_noise', 0)):,} ({fmt_pct(pct(s.get('n_broker_noise', 0), n_res))})")

        lines.append("")
        lines.append("Flags:")
        if inh_share >= 0.70:
            lines.append("- Inherited share is high. Check inherited samples carefully.")
        else:
            lines.append("- Inherited share is not extreme.")
        if float(s.get("inherited_short_reply_share") or 0) >= 0.55:
            lines.append("- Many inherited matches are short replies. Consider down-weighting or filtering very short replies.")
        else:
            lines.append("- Short replies do not dominate inherited matches.")
        if int(s.get('n_suppressed', 0)) > 0:
            lines.append(
                f"- {int(s['n_suppressed']):,} explicit matches were suppressed by broker-noise filter; "
                f"{int(s.get('n_suppressed_dropped', 0)):,} were dropped finally and "
                f"{int(s.get('n_suppressed_but_inherited', 0)):,} later inherited context."
            )
        if not tg["inherited_flags"].empty:
            suspicious = tg["inherited_flags"].loc[
                tg["inherited_flags"]["flag_inherited_dominates"] | tg["inherited_flags"]["flag_short_reply_heavy"]
            ]
            lines.append(f"- Issuer/channel combinations flagged: {len(suspicious):,}")

    lines.append("")
    lines.append("2) SMART-LAB")
    lines.append("-" * 80)
    if "post_stats" in sl:
        s = sl["post_stats"]
        lines.append(f"Posts total: {int(s['n_total']):,}; resolved: {int(s['n_resolved']):,} ({fmt_pct(pct(s['n_resolved'], s['n_total']))})")
        lines.append(f"Posts explicit: {int(s['n_explicit']):,}; offering-level: {int(s['n_offering']):,}; issuer-level: {int(s['n_issuer']):,}")
        lines.append(f"Posts suppressed: {int(s.get('n_suppressed', 0)):,}")
        n_res = int(s['n_resolved']) or 1
        lines.append(f"Posts bond_context: {int(s.get('n_bond_context', 0)):,} ({fmt_pct(pct(s.get('n_bond_context', 0), n_res))})")
        lines.append(f"Posts primary_context: {int(s.get('n_primary_context', 0)):,} ({fmt_pct(pct(s.get('n_primary_context', 0), n_res))})")
    else:
        lines.append("Posts resolved table missing.")

    if "comment_stats" in sl:
        s = sl["comment_stats"]
        lines.append(f"Comments total: {int(s['n_total']):,}; resolved: {int(s['n_resolved']):,} ({fmt_pct(pct(s['n_resolved'], s['n_total']))})")
        lines.append(f"Comments explicit: {int(s['n_explicit']):,}; inherited: {int(s['n_inherited']):,}; offering-level: {int(s['n_offering']):,}; issuer-level: {int(s['n_issuer']):,}")
        lines.append(f"Comments suppressed: {int(s.get('n_suppressed', 0)):,}")
        n_res = int(s['n_resolved']) or 1
        lines.append(f"Comments bond_context: {int(s.get('n_bond_context', 0)):,} ({fmt_pct(pct(s.get('n_bond_context', 0), n_res))})")
        lines.append(f"Comments primary_context: {int(s.get('n_primary_context', 0)):,} ({fmt_pct(pct(s.get('n_primary_context', 0), n_res))})")
    else:
        lines.append("Comments resolved table missing.")

    lines.append("")
    lines.append("3) COVERAGE AROUND OFFERINGS")
    lines.append("-" * 80)
    if cov.get("missing"):
        lines.append("Coverage file not found or empty.")
    else:
        s = cov["summary"]
        has_v8 = cov.get("has_v8", False)
        lines.append(f"Offerings: {s['n_offerings']:,}")

        if has_v8:
            for label in ["entity_only", "bond", "primary"]:
                n_with = s.get(f"n_with_{label}", 0)
                n_zero = s.get(f"n_zero_{label}", 0)
                mean_v = s.get(f"mean_{label}", 0)
                median_v = s.get(f"median_{label}", 0)
                p90_v = s.get(f"p90_{label}", 0)
                lines.append(f"  [{label}] with mentions: {n_with:,}; zero: {n_zero:,}; mean: {mean_v:.2f}; median: {median_v:.2f}; P90: {p90_v:.2f}")
        else:
            n_with = s.get("n_with_entity_only", 0)
            n_zero = s.get("n_zero_entity_only", 0)
            lines.append(f"With any mention: {n_with:,} ({fmt_pct(pct(n_with, s['n_offerings']))})")
            lines.append(f"Zero coverage: {n_zero:,}")
            for k in ["mean_entity_only", "median_entity_only", "p90_entity_only", "p95_entity_only"]:
                if k in s:
                    lines.append(f"  {k}: {s[k]:.2f}")

        lines.append("")
        lines.append("Flags:")
        # Use primary layer if v8, else entity_only
        key = "n_zero_primary" if has_v8 else "n_zero_entity_only"
        zero_frac = pct(s.get(key, 0), s['n_offerings'])
        if zero_frac >= 0.50:
            lines.append("- A large share of offerings has zero primary-context coverage. This is expected for less-discussed issuers.")
        elif zero_frac >= 0.35:
            lines.append("- Moderate zero-coverage rate. Check early years separately.")
        else:
            lines.append("- Zero coverage is not extreme.")

    lines.append("")
    lines.append("4) WHAT TO READ FIRST")
    lines.append("-" * 80)
    lines.append("- samples_tg_explicit.csv  (check bond_context / primary_context columns)")
    lines.append("- samples_tg_inherited.csv")
    lines.append("- samples_sl_posts.csv")
    lines.append("- samples_sl_comments.csv")
    lines.append("- top_issuers.csv  (bond_context vs broker_noise per issuer)")
    lines.append("- telegram_inherited_flags.csv")
    lines.append("- alias_ambiguity.csv")
    lines.append("")
    lines.append("Suggested rule of thumb:")
    lines.append("- Primary regression: use primary_context=1 mentions only.")
    lines.append("- Robustness: use bond_context=1 mentions.")
    lines.append("- Baseline: use entity_only (all resolved, no topical filter).")
    lines.append("- If broker_noise share is high for a dual-role issuer, the suppress filter is working.")
    return "\n".join(lines)


# -----------------------------
# Main
# -----------------------------

def main():
    parser = argparse.ArgumentParser(description="Audit bond entity matching outputs and produce QA files.")
    parser.add_argument("--telegram-db", type=Path, default=Path("data.db"))
    parser.add_argument("--smartlab-db", type=Path, default=Path("smartlab_bonds.db"))
    parser.add_argument("--coverage", type=Path, default=Path("mention_coverage_window7.csv"))
    parser.add_argument("--out-dir", type=Path, default=Path("match_audit_output"))
    parser.add_argument("--sample-size", type=int, default=50)
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    ensure_dir(args.out_dir)

    # Telegram
    tg_conn = sqlite3.connect(args.telegram_db)
    tg = audit_telegram(tg_conn, sample_size=args.sample_size, seed=args.seed)
    tg_conn.close()

    # Smart-Lab
    sl_conn = sqlite3.connect(args.smartlab_db)
    sl = audit_smartlab(sl_conn, sample_size=args.sample_size)
    sl_conn.close()

    # Coverage
    cov_df = load_coverage(args.coverage)
    cov = audit_coverage(cov_df)

    # Save tables
    if isinstance(tg, dict):
        for key in ["top_issuers", "top_offerings", "inherited_flags", "samples_tg_explicit", "samples_tg_inherited"]:
            df = tg.get(key)
            if isinstance(df, pd.DataFrame):
                df.to_csv(args.out_dir / f"{key}.csv", index=False, encoding="utf-8-sig")

    if isinstance(sl, dict):
        for key in ["samples_sl_posts", "samples_sl_comments"]:
            df = sl.get(key)
            if isinstance(df, pd.DataFrame):
                df.to_csv(args.out_dir / f"{key}.csv", index=False, encoding="utf-8-sig")

    alias_parts = []
    if isinstance(tg.get("alias_counts"), pd.DataFrame) and not tg["alias_counts"].empty:
        alias_parts.append(tg["alias_counts"])
    if isinstance(sl.get("alias_counts"), pd.DataFrame) and not sl["alias_counts"].empty:
        alias_parts.append(sl["alias_counts"])
    if alias_parts:
        alias_ambiguity = pd.concat(alias_parts, ignore_index=True)
        alias_ambiguity = (
            alias_ambiguity.groupby("alias", as_index=False)
            .agg(
                rows=("rows", "sum"),
                distinct_issuers=("distinct_issuers", "max"),
                distinct_offerings=("distinct_offerings", "max"),
                sources=("source", lambda s: ", ".join(sorted(set(s))))
            )
            .sort_values(["distinct_issuers", "distinct_offerings", "rows"], ascending=[False, False, False])
        )
        alias_ambiguity["alias_len"] = alias_ambiguity["alias"].astype(str).str.len()
        alias_ambiguity["flag_suspicious"] = (
            (alias_ambiguity["distinct_issuers"] >= 3) |
            (alias_ambiguity["distinct_offerings"] >= 5) |
            ((alias_ambiguity["alias_len"] <= 4) & (alias_ambiguity["rows"] >= 20))
        )
        alias_ambiguity.to_csv(args.out_dir / "alias_ambiguity.csv", index=False, encoding="utf-8-sig")

    if isinstance(cov, dict) and not cov.get("missing"):
        cov["distribution"].to_csv(args.out_dir / "coverage_summary.csv", index=False, encoding="utf-8-sig")
        cov["zero_coverage"].to_csv(args.out_dir / "coverage_zero_cases.csv", index=False, encoding="utf-8-sig")
        cov["high_coverage"].to_csv(args.out_dir / "coverage_high_cases.csv", index=False, encoding="utf-8-sig")

    report = build_report(tg, sl, cov)
    (args.out_dir / "audit_report.txt").write_text(report, encoding="utf-8")
    print(report)
    print(f"\nSaved outputs to: {args.out_dir.resolve()}")


if __name__ == "__main__":
    main()
