#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Audit bond-entity matcher v10 outputs.

Compatible with bond_entity_matcher_v10.py.

v10 philosophy
--------------
The matcher is now a candidate generator. It finds entity mentions and saves
ALL matches. It does not decide whether the text is bond-relevant, whether the
issuer is the subject, whether the mention is about a deposit, or whether an
agent/broker mention should count. Those questions belong downstream to the
LLM layer.

This audit therefore focuses on candidate quality and LLM input readiness:
  - coverage and volume of candidate rows
  - texts with multiple issuer/offering candidates
  - aliases that create broad or suspicious candidate sets
  - inherited matches that may inflate the LLM workload
  - likely false negatives: bond-looking texts with no entity candidates

Outputs
-------
Creates an output directory with:

  audit_report.txt
  top_issuers.csv
  top_offerings.csv
  multi_issuer_texts.csv
  multi_offering_texts.csv
  samples_tg_explicit.csv
  samples_tg_inherited.csv
  samples_tg_multi_issuer.csv
  samples_tg_unmatched_bond_like.csv
  samples_sl_posts.csv
  samples_sl_comments.csv
  samples_sl_multi_issuer.csv
  sl_post_multi_issuer.csv
  sl_comment_multi_issuer.csv
  inherited_flags.csv
  alias_ambiguity.csv
  coverage_summary.csv
  coverage_zero_cases.csv
  coverage_high_cases.csv
  char_budget.csv
"""

from __future__ import annotations

import argparse
import json
import sqlite3
from pathlib import Path
from typing import Iterable, Optional

import numpy as np
import pandas as pd


BOND_KEYWORDS = [
    "облигац",
    "купон",
    "размещ",
    "размещение",
    "эмитент",
    "выпуск",
    "оферта",
    "доходност",
    "номинал",
    "дюрац",
    "book",
    "букбилд",
]


# -----------------------------
# Helpers
# -----------------------------

def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def read_sql(conn: sqlite3.Connection, query: str, params: Optional[Iterable] = None) -> pd.DataFrame:
    return pd.read_sql_query(query, conn, params=params or ())


def table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    q = "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?"
    return conn.execute(q, (table_name,)).fetchone() is not None


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
    return f"{100 * x:.1f}%"


def sql_like_any(column: str, keywords: list[str]) -> str:
    return " OR ".join([f"LOWER({column}) LIKE '%{kw.lower()}%'" for kw in keywords])


def flatten_alias_counts(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    cols = ["alias", "source", "rows", "distinct_texts", "distinct_issuers", "distinct_offerings"]
    if df.empty or "matched_aliases" not in df.columns:
        return pd.DataFrame(columns=cols)

    tmp = df.copy()
    tmp["alias_list"] = tmp["matched_aliases"].apply(safe_json_list)
    tmp = tmp.explode("alias_list")
    tmp = tmp[tmp["alias_list"].notna() & (tmp["alias_list"].astype(str).str.len() > 0)].copy()
    if tmp.empty:
        return pd.DataFrame(columns=cols)

    out = (
        tmp.groupby("alias_list", dropna=False)
        .agg(
            rows=("source_key", "size"),
            distinct_texts=("source_key", "nunique"),
            distinct_issuers=("issuer", lambda s: s.dropna().nunique()),
            distinct_offerings=("offering_id", lambda s: s.dropna().nunique()),
        )
        .reset_index()
        .rename(columns={"alias_list": "alias"})
    )
    out["source"] = source_name
    return out[cols]


def sample_df(df: pd.DataFrame, sample_size: int, seed: int) -> pd.DataFrame:
    if df.empty or len(df) <= sample_size:
        return df
    return df.sample(n=sample_size, random_state=seed)


def token_range(chars: float) -> str:
    low = chars / 4.0
    high = chars / 2.5
    return f"{low:,.0f}-{high:,.0f}"


# -----------------------------
# Telegram
# -----------------------------

def audit_telegram(conn: sqlite3.Connection, sample_size: int, seed: int) -> dict:
    required = ["telegram_matches", "messages"]
    missing = [t for t in required if not table_exists(conn, t)]
    if missing:
        return {"missing": missing}

    text_count = read_sql(conn, "SELECT COUNT(*) AS n FROM messages").iloc[0]["n"]

    stats = read_sql(conn, """
        SELECT
            COUNT(*) AS n_candidate_rows,
            COUNT(DISTINCT source_key) AS n_matched_texts,
            COUNT(DISTINCT issuer) AS n_issuers,
            COUNT(DISTINCT offering_id) AS n_offerings,
            SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS n_explicit_rows,
            SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS n_inherited_rows,
            COUNT(DISTINCT CASE WHEN match_origin='explicit' THEN source_key END) AS n_explicit_texts,
            COUNT(DISTINCT CASE WHEN match_origin='inherited' THEN source_key END) AS n_inherited_texts,
            SUM(CASE WHEN entity_level='issuer' THEN 1 ELSE 0 END) AS n_issuer_rows,
            SUM(CASE WHEN entity_level='offering' THEN 1 ELSE 0 END) AS n_offering_rows,
            AVG(text_len) AS mean_text_len
        FROM telegram_matches
    """).iloc[0].to_dict()
    stats["n_total_texts"] = int(text_count)
    stats["n_unmatched_texts"] = int(text_count - stats["n_matched_texts"])

    char_budget = read_sql(conn, """
        WITH per_text AS (
            SELECT source_key, MAX(text_len) AS chars, COUNT(*) AS candidate_rows
            FROM telegram_matches
            GROUP BY source_key
        )
        SELECT
            'telegram' AS source,
            COUNT(*) AS matched_texts,
            SUM(candidate_rows) AS candidate_rows,
            SUM(chars) AS unique_chars,
            SUM(chars * candidate_rows) AS candidate_weighted_chars,
            ROUND(1.0 * SUM(chars * candidate_rows) / NULLIF(SUM(chars), 0), 4) AS duplication_factor
        FROM per_text
    """)

    text_complexity = read_sql(conn, """
        SELECT
            source_key,
            channel,
            message_id,
            parent_id,
            MIN(published_at) AS published_at,
            COUNT(*) AS candidate_rows,
            COUNT(DISTINCT issuer) AS issuer_candidates,
            COUNT(DISTINCT offering_id) AS offering_candidates,
            SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS inherited_rows,
            GROUP_CONCAT(DISTINCT issuer) AS candidate_issuers,
            GROUP_CONCAT(DISTINCT COALESCE(bond_name, CAST(offering_id AS TEXT))) AS candidate_offerings
        FROM telegram_matches
        GROUP BY source_key, channel, message_id, parent_id
    """)

    multi_issuer_texts = text_complexity[text_complexity["issuer_candidates"] >= 2].copy()
    multi_issuer_texts = multi_issuer_texts.sort_values(
        ["issuer_candidates", "candidate_rows"], ascending=[False, False]
    )

    multi_offering_texts = text_complexity[text_complexity["offering_candidates"] >= 2].copy()
    multi_offering_texts = multi_offering_texts.sort_values(
        ["offering_candidates", "candidate_rows"], ascending=[False, False]
    )

    top_issuers = read_sql(conn, """
        SELECT issuer,
               COUNT(*) AS candidate_rows,
               COUNT(DISTINCT source_key) AS matched_texts,
               SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS explicit_rows,
               SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS inherited_rows,
               SUM(CASE WHEN entity_level='offering' THEN 1 ELSE 0 END) AS offering_rows,
               SUM(CASE WHEN entity_level='issuer' THEN 1 ELSE 0 END) AS issuer_rows,
               COUNT(DISTINCT offering_id) AS distinct_offerings
        FROM telegram_matches
        GROUP BY issuer
        ORDER BY candidate_rows DESC
        LIMIT 150
    """)

    top_offerings = read_sql(conn, """
        SELECT COALESCE(bond_name, CAST(offering_id AS TEXT)) AS offering,
               issuer,
               offering_id,
               COUNT(*) AS candidate_rows,
               COUNT(DISTINCT source_key) AS matched_texts,
               SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS explicit_rows,
               SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS inherited_rows
        FROM telegram_matches
        WHERE offering_id IS NOT NULL
        GROUP BY offering_id, bond_name, issuer
        ORDER BY candidate_rows DESC
        LIMIT 150
    """)

    inherited_flags = read_sql(conn, """
        SELECT issuer,
               channel,
               COUNT(*) AS candidate_rows,
               COUNT(DISTINCT source_key) AS matched_texts,
               SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS explicit_rows,
               SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS inherited_rows,
               ROUND(1.0 * SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END)
                     / NULLIF(COUNT(*), 0), 4) AS inherited_row_share,
               ROUND(AVG(CASE WHEN match_origin='inherited' THEN text_len END), 1) AS inherited_mean_text_len
        FROM telegram_matches
        GROUP BY issuer, channel
        HAVING COUNT(*) >= 20
        ORDER BY inherited_row_share DESC, candidate_rows DESC
    """)
    if not inherited_flags.empty:
        inherited_flags["flag_inherited_dominates"] = (
            (inherited_flags["inherited_row_share"] >= 0.75) & (inherited_flags["candidate_rows"] >= 30)
        )

    samples_tg_explicit = read_sql(conn, f"""
        SELECT r.channel, r.message_id, r.parent_id, r.issuer, r.bond_name, r.entity_level,
               r.match_origin, r.matched_aliases, r.alias_types, r.text_len, m.text
        FROM telegram_matches r
        JOIN messages m ON r.channel = m.channel AND r.message_id = m.id
        WHERE r.match_origin = 'explicit'
        ORDER BY RANDOM()
        LIMIT {int(sample_size)}
    """)

    samples_tg_inherited = read_sql(conn, f"""
        SELECT r.channel, r.message_id, r.parent_id, r.issuer, r.bond_name, r.entity_level,
               r.match_origin, r.inherit_anchor_key, r.matched_aliases, r.alias_types,
               r.text_len, m.text AS reply_text, p.text AS parent_text
        FROM telegram_matches r
        JOIN messages m ON r.channel = m.channel AND r.message_id = m.id
        LEFT JOIN messages p ON r.channel = p.channel AND r.parent_id = p.id
        WHERE r.match_origin = 'inherited'
        ORDER BY RANDOM()
        LIMIT {int(sample_size)}
    """)

    samples_tg_multi_issuer = pd.DataFrame()
    if not multi_issuer_texts.empty:
        sampled_keys = sample_df(multi_issuer_texts, sample_size, seed)["source_key"].tolist()
        placeholders = ",".join(["?"] * len(sampled_keys))
        samples_tg_multi_issuer = read_sql(conn, f"""
            SELECT x.source_key, x.channel, x.message_id, x.candidate_rows,
                   x.issuer_candidates, x.candidate_issuers, m.text
            FROM (
                SELECT source_key, channel, message_id,
                       COUNT(*) AS candidate_rows,
                       COUNT(DISTINCT issuer) AS issuer_candidates,
                       GROUP_CONCAT(DISTINCT issuer) AS candidate_issuers
                FROM telegram_matches
                WHERE source_key IN ({placeholders})
                GROUP BY source_key, channel, message_id
            ) x
            JOIN messages m ON x.channel = m.channel AND x.message_id = m.id
            ORDER BY x.issuer_candidates DESC, x.candidate_rows DESC
        """, sampled_keys)

    samples_tg_unmatched_bond_like = read_sql(conn, f"""
        SELECT m.channel, m.id AS message_id, m.date AS published_at, m.text
        FROM messages m
        WHERE ({sql_like_any("m.text", BOND_KEYWORDS)})
          AND length(m.text) >= 50
          AND NOT EXISTS (
              SELECT 1
              FROM telegram_matches r
              WHERE r.source_key = m.channel || ':' || m.id
          )
        LIMIT {int(sample_size)}
    """)

    alias_counts = flatten_alias_counts(
        read_sql(conn, "SELECT source_key, issuer, offering_id, matched_aliases FROM telegram_matches"),
        "telegram",
    )

    return {
        "stats": stats,
        "top_issuers": top_issuers,
        "top_offerings": top_offerings,
        "multi_issuer_texts": multi_issuer_texts.head(500),
        "multi_offering_texts": multi_offering_texts.head(500),
        "inherited_flags": inherited_flags,
        "samples_tg_explicit": samples_tg_explicit,
        "samples_tg_inherited": samples_tg_inherited,
        "samples_tg_multi_issuer": samples_tg_multi_issuer,
        "samples_tg_unmatched_bond_like": samples_tg_unmatched_bond_like,
        "alias_counts": alias_counts,
        "char_budget": char_budget,
    }


# -----------------------------
# Smart-Lab
# -----------------------------

def audit_smartlab(conn: sqlite3.Connection, sample_size: int, seed: int) -> dict:
    out: dict = {}
    alias_parts = []

    if table_exists(conn, "smartlab_post_matches") and table_exists(conn, "posts"):
        post_total = read_sql(conn, "SELECT COUNT(*) AS n FROM posts").iloc[0]["n"]
        post_stats = read_sql(conn, """
            SELECT COUNT(*) AS n_candidate_rows,
                   COUNT(DISTINCT source_key) AS n_matched_texts,
                   COUNT(DISTINCT issuer) AS n_issuers,
                   COUNT(DISTINCT offering_id) AS n_offerings,
                   SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS n_explicit_rows,
                   SUM(CASE WHEN entity_level='issuer' THEN 1 ELSE 0 END) AS n_issuer_rows,
                   SUM(CASE WHEN entity_level='offering' THEN 1 ELSE 0 END) AS n_offering_rows,
                   AVG(text_len) AS mean_text_len
            FROM smartlab_post_matches
        """).iloc[0].to_dict()
        post_stats["n_total_texts"] = int(post_total)
        post_stats["n_unmatched_texts"] = int(post_total - post_stats["n_matched_texts"])
        out["post_stats"] = post_stats

        out["post_char_budget"] = read_sql(conn, """
            WITH per_text AS (
                SELECT source_key, MAX(text_len) AS chars, COUNT(*) AS candidate_rows
                FROM smartlab_post_matches
                GROUP BY source_key
            )
            SELECT
                'smartlab_posts' AS source,
                COUNT(*) AS matched_texts,
                SUM(candidate_rows) AS candidate_rows,
                SUM(chars) AS unique_chars,
                SUM(chars * candidate_rows) AS candidate_weighted_chars,
                ROUND(1.0 * SUM(chars * candidate_rows) / NULLIF(SUM(chars), 0), 4) AS duplication_factor
            FROM per_text
        """)

        out["samples_sl_posts"] = read_sql(conn, f"""
            SELECT r.post_id, r.issuer, r.bond_name, r.entity_level, r.match_origin,
                   r.matched_aliases, r.alias_types, r.text_len, p.title, p.tags, p.text
            FROM smartlab_post_matches r
            JOIN posts p ON r.post_id = p.post_id
            ORDER BY RANDOM()
            LIMIT {int(sample_size)}
        """)

        post_multi = read_sql(conn, """
            SELECT source_key, post_id,
                   COUNT(*) AS candidate_rows,
                   COUNT(DISTINCT issuer) AS issuer_candidates,
                   COUNT(DISTINCT offering_id) AS offering_candidates,
                   GROUP_CONCAT(DISTINCT issuer) AS candidate_issuers
            FROM smartlab_post_matches
            GROUP BY source_key, post_id
            HAVING COUNT(DISTINCT issuer) >= 2
            ORDER BY issuer_candidates DESC, candidate_rows DESC
            LIMIT 500
        """)
        out["sl_post_multi_issuer"] = post_multi

        alias_parts.append(flatten_alias_counts(
            read_sql(conn, "SELECT source_key, issuer, offering_id, matched_aliases FROM smartlab_post_matches"),
            "smartlab_posts",
        ))
    else:
        out["post_missing"] = True

    if table_exists(conn, "smartlab_comment_matches") and table_exists(conn, "comments"):
        comment_total = read_sql(conn, "SELECT COUNT(*) AS n FROM comments").iloc[0]["n"]
        comment_stats = read_sql(conn, """
            SELECT COUNT(*) AS n_candidate_rows,
                   COUNT(DISTINCT source_key) AS n_matched_texts,
                   COUNT(DISTINCT issuer) AS n_issuers,
                   COUNT(DISTINCT offering_id) AS n_offerings,
                   SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS n_explicit_rows,
                   SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS n_inherited_rows,
                   COUNT(DISTINCT CASE WHEN match_origin='inherited' THEN source_key END) AS n_inherited_texts,
                   SUM(CASE WHEN entity_level='issuer' THEN 1 ELSE 0 END) AS n_issuer_rows,
                   SUM(CASE WHEN entity_level='offering' THEN 1 ELSE 0 END) AS n_offering_rows,
                   AVG(text_len) AS mean_text_len
            FROM smartlab_comment_matches
        """).iloc[0].to_dict()
        comment_stats["n_total_texts"] = int(comment_total)
        comment_stats["n_unmatched_texts"] = int(comment_total - comment_stats["n_matched_texts"])
        out["comment_stats"] = comment_stats

        out["comment_char_budget"] = read_sql(conn, """
            WITH per_text AS (
                SELECT source_key, MAX(text_len) AS chars, COUNT(*) AS candidate_rows
                FROM smartlab_comment_matches
                GROUP BY source_key
            )
            SELECT
                'smartlab_comments' AS source,
                COUNT(*) AS matched_texts,
                SUM(candidate_rows) AS candidate_rows,
                SUM(chars) AS unique_chars,
                SUM(chars * candidate_rows) AS candidate_weighted_chars,
                ROUND(1.0 * SUM(chars * candidate_rows) / NULLIF(SUM(chars), 0), 4) AS duplication_factor
            FROM per_text
        """)

        out["samples_sl_comments"] = read_sql(conn, f"""
            SELECT r.comment_id, r.post_id, r.issuer, r.bond_name, r.entity_level,
                   r.match_origin, r.inherit_anchor_key, r.matched_aliases, r.alias_types,
                   r.text_len, c.text AS comment_text, p.title AS post_title
            FROM smartlab_comment_matches r
            JOIN comments c ON r.comment_id = c.comment_id
            LEFT JOIN posts p ON r.post_id = p.post_id
            ORDER BY RANDOM()
            LIMIT {int(sample_size)}
        """)

        comment_multi = read_sql(conn, """
            SELECT source_key, comment_id, post_id,
                   COUNT(*) AS candidate_rows,
                   COUNT(DISTINCT issuer) AS issuer_candidates,
                   COUNT(DISTINCT offering_id) AS offering_candidates,
                   GROUP_CONCAT(DISTINCT issuer) AS candidate_issuers
            FROM smartlab_comment_matches
            GROUP BY source_key, comment_id, post_id
            HAVING COUNT(DISTINCT issuer) >= 2
            ORDER BY issuer_candidates DESC, candidate_rows DESC
            LIMIT 500
        """)
        out["sl_comment_multi_issuer"] = comment_multi

        if not comment_multi.empty:
            sampled = sample_df(comment_multi, sample_size, seed)
            keys = sampled["source_key"].tolist()
            placeholders = ",".join(["?"] * len(keys))
            out["samples_sl_multi_issuer"] = read_sql(conn, f"""
                SELECT x.source_key, x.comment_id, x.post_id, x.candidate_rows,
                       x.issuer_candidates, x.candidate_issuers,
                       c.text AS comment_text, p.title AS post_title
                FROM (
                    SELECT source_key, comment_id, post_id,
                           COUNT(*) AS candidate_rows,
                           COUNT(DISTINCT issuer) AS issuer_candidates,
                           GROUP_CONCAT(DISTINCT issuer) AS candidate_issuers
                    FROM smartlab_comment_matches
                    WHERE source_key IN ({placeholders})
                    GROUP BY source_key, comment_id, post_id
                ) x
                JOIN comments c ON x.comment_id = c.comment_id
                LEFT JOIN posts p ON x.post_id = p.post_id
                ORDER BY x.issuer_candidates DESC, x.candidate_rows DESC
            """, keys)
        else:
            out["samples_sl_multi_issuer"] = pd.DataFrame()

        alias_parts.append(flatten_alias_counts(
            read_sql(conn, "SELECT source_key, issuer, offering_id, matched_aliases FROM smartlab_comment_matches"),
            "smartlab_comments",
        ))
    else:
        out["comment_missing"] = True

    out["alias_counts"] = pd.concat(alias_parts, ignore_index=True) if alias_parts else pd.DataFrame()
    return out


# -----------------------------
# Coverage
# -----------------------------

def audit_coverage(coverage_df: Optional[pd.DataFrame]) -> dict:
    if coverage_df is None or coverage_df.empty or "all_mentions" not in coverage_df.columns:
        return {"missing": True}

    df = coverage_df.copy()
    mention_cols = [c for c in ["tg_mentions", "sl_post_mentions", "sl_comment_mentions", "all_mentions"] if c in df.columns]
    for col in mention_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    summary = {
        "n_offerings": int(len(df)),
        "n_with_any_mention": int((df["all_mentions"] > 0).sum()),
        "n_zero_coverage": int((df["all_mentions"] == 0).sum()),
        "mean_all_mentions": float(df["all_mentions"].mean()),
        "median_all_mentions": float(df["all_mentions"].median()),
        "p90_all_mentions": float(df["all_mentions"].quantile(0.90)),
        "p95_all_mentions": float(df["all_mentions"].quantile(0.95)),
    }
    for col in mention_cols:
        summary[f"mean_{col}"] = float(df[col].mean())
        summary[f"n_with_{col}"] = int((df[col] > 0).sum())

    dist = pd.DataFrame([{"metric": k, "value": v} for k, v in summary.items()])
    id_cols = [c for c in ["offering_id", "issuer", "bond_name", "ISIN", "book_date", "placement_date", "window_days", "all_mentions"] if c in df.columns]
    zero_cov = df.loc[df["all_mentions"] == 0, id_cols].copy()
    high_cov = df.sort_values("all_mentions", ascending=False).head(100)

    return {"summary": summary, "distribution": dist, "zero_coverage": zero_cov, "high_coverage": high_cov}


# -----------------------------
# Report
# -----------------------------

def build_report(tg: dict, sl: dict, cov: dict) -> str:
    lines: list[str] = []
    lines.append("BOND MATCH AUDIT REPORT (v10 — candidate-generator audit)")
    lines.append("=" * 80)
    lines.append("")
    lines.append("Matcher v10 does not resolve relevance, topic, sentiment, deposits vs bonds,")
    lines.append("or agent-vs-issuer roles. Treat every row here as an LLM candidate.")
    lines.append("")

    lines.append("1) TELEGRAM")
    lines.append("-" * 80)
    if "missing" in tg:
        lines.append(f"Missing tables: {', '.join(tg['missing'])}")
    else:
        s = tg["stats"]
        total = int(s["n_total_texts"])
        matched = int(s["n_matched_texts"])
        rows = int(s["n_candidate_rows"])
        lines.append(f"Total messages:       {total:,}")
        lines.append(f"Matched texts:        {matched:,} ({fmt_pct(pct(matched, total))})")
        lines.append(f"Unmatched texts:      {int(s['n_unmatched_texts']):,}")
        lines.append(f"Candidate rows:       {rows:,}")
        lines.append(f"Rows per matched text:{pct(rows, matched):.2f}")
        lines.append(f"Issuers covered:      {int(s['n_issuers']):,}")
        lines.append(f"Offerings covered:    {int(s['n_offerings']):,}")
        lines.append(f"Explicit rows:        {int(s['n_explicit_rows']):,}")
        lines.append(f"Inherited rows:       {int(s['n_inherited_rows']):,}")
        lines.append(f"Offering-level rows:  {int(s['n_offering_rows']):,}")
        lines.append(f"Issuer-level rows:    {int(s['n_issuer_rows']):,}")
        lines.append("")
        lines.append("QA priorities:")
        lines.append("  1. samples_tg_explicit.csv — precision of raw alias candidates.")
        lines.append("  2. samples_tg_multi_issuer.csv — LLM disambiguation workload and context needs.")
        lines.append("  3. samples_tg_inherited.csv — inherited candidates that may over-expand context.")
        lines.append("  4. samples_tg_unmatched_bond_like.csv — likely missing aliases.")
        if not tg["multi_issuer_texts"].empty:
            lines.append(f"Multi-issuer matched texts: {len(tg['multi_issuer_texts']):,} shown in multi_issuer_texts.csv")
        if not tg["multi_offering_texts"].empty:
            lines.append(f"Multi-offering matched texts: {len(tg['multi_offering_texts']):,} shown in multi_offering_texts.csv")
        if not tg["inherited_flags"].empty:
            flagged = int(tg["inherited_flags"]["flag_inherited_dominates"].sum())
            lines.append(f"Inheritance flags: {flagged:,} issuer/channel groups where inherited rows dominate.")

    lines.append("")
    lines.append("2) SMART-LAB")
    lines.append("-" * 80)
    if "post_stats" in sl:
        s = sl["post_stats"]
        lines.append(f"Posts total:     {int(s['n_total_texts']):,}")
        lines.append(f"Posts matched:   {int(s['n_matched_texts']):,} ({fmt_pct(pct(s['n_matched_texts'], s['n_total_texts']))})")
        lines.append(f"Post candidates: {int(s['n_candidate_rows']):,}")
    else:
        lines.append("Post match table missing.")

    if "comment_stats" in sl:
        s = sl["comment_stats"]
        lines.append(f"Comments total:     {int(s['n_total_texts']):,}")
        lines.append(f"Comments matched:   {int(s['n_matched_texts']):,} ({fmt_pct(pct(s['n_matched_texts'], s['n_total_texts']))})")
        lines.append(f"Comment candidates: {int(s['n_candidate_rows']):,}")
        lines.append(f"Comment inherited:  {int(s.get('n_inherited_rows', 0)):,}")
    else:
        lines.append("Comment match table missing.")

    lines.append("")
    lines.append("3) COVERAGE AROUND OFFERINGS")
    lines.append("-" * 80)
    if cov.get("missing"):
        lines.append("Coverage file not found or not in v10 all_mentions format.")
    else:
        s = cov["summary"]
        lines.append(f"Offerings:         {s['n_offerings']:,}")
        lines.append(f"With any mention:  {s['n_with_any_mention']:,} ({fmt_pct(pct(s['n_with_any_mention'], s['n_offerings']))})")
        lines.append(f"Zero coverage:     {s['n_zero_coverage']:,} ({fmt_pct(pct(s['n_zero_coverage'], s['n_offerings']))})")
        lines.append(f"Mean mentions:     {s['mean_all_mentions']:.2f}")
        lines.append(f"Median mentions:   {s['median_all_mentions']:.2f}")
        lines.append(f"P90/P95 mentions:  {s['p90_all_mentions']:.2f} / {s['p95_all_mentions']:.2f}")

    lines.append("")
    lines.append("4) LLM CHARACTER BUDGET")
    lines.append("-" * 80)
    char_parts = []
    if isinstance(tg.get("char_budget"), pd.DataFrame) and not tg["char_budget"].empty:
        char_parts.append(tg["char_budget"])
    for key in ["post_char_budget", "comment_char_budget"]:
        if isinstance(sl.get(key), pd.DataFrame) and not sl[key].empty:
            char_parts.append(sl[key])
    if char_parts:
        char_df = pd.concat(char_parts, ignore_index=True)
        for _, row in char_df.iterrows():
            unique_chars = float(row["unique_chars"] or 0)
            weighted_chars = float(row["candidate_weighted_chars"] or 0)
            lines.append(
                f"{row['source']}: unique={unique_chars:,.0f} chars "
                f"(~{token_range(unique_chars)} tokens); per-candidate={weighted_chars:,.0f} chars "
                f"(~{token_range(weighted_chars)} tokens); x{float(row['duplication_factor']):.2f}"
            )
        unique_total = float(char_df["unique_chars"].fillna(0).sum())
        weighted_total = float(char_df["candidate_weighted_chars"].fillna(0).sum())
        lines.append(
            f"TOTAL: unique={unique_total:,.0f} chars (~{token_range(unique_total)} tokens); "
            f"per-candidate={weighted_total:,.0f} chars (~{token_range(weighted_total)} tokens)."
        )
    else:
        lines.append("No match tables available for character budget.")

    lines.append("")
    lines.append("5) WHAT THIS AUDIT NO LONGER DOES")
    lines.append("-" * 80)
    lines.append("No ambiguous/suppressed/resolved diagnostics are produced, because v10")
    lines.append("intentionally stopped making those decisions. Use the sample files to tune")
    lines.append("alias dictionaries and to design the downstream LLM classifier prompt/schema.")
    return "\n".join(lines)


# -----------------------------
# Main
# -----------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Audit bond_entity_matcher_v10 candidate outputs.")
    parser.add_argument("--telegram-db", type=Path, default=Path("data.db"))
    parser.add_argument("--smartlab-db", type=Path, default=Path("smartlab_bonds.db"))
    parser.add_argument("--coverage", type=Path, default=Path("mention_coverage_window7.csv"))
    parser.add_argument("--out-dir", type=Path, default=Path("match_audit_output"))
    parser.add_argument("--sample-size", type=int, default=50)
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    ensure_dir(args.out_dir)

    tg_conn = sqlite3.connect(args.telegram_db)
    tg = audit_telegram(tg_conn, sample_size=args.sample_size, seed=args.seed)
    tg_conn.close()

    sl_conn = sqlite3.connect(args.smartlab_db)
    sl = audit_smartlab(sl_conn, sample_size=args.sample_size, seed=args.seed)
    sl_conn.close()

    cov = audit_coverage(load_coverage(args.coverage))

    for key in [
        "top_issuers",
        "top_offerings",
        "multi_issuer_texts",
        "multi_offering_texts",
        "inherited_flags",
        "samples_tg_explicit",
        "samples_tg_inherited",
        "samples_tg_multi_issuer",
        "samples_tg_unmatched_bond_like",
    ]:
        df = tg.get(key) if isinstance(tg, dict) else None
        if isinstance(df, pd.DataFrame):
            df.to_csv(args.out_dir / f"{key}.csv", index=False, encoding="utf-8-sig")

    for key in [
        "samples_sl_posts",
        "samples_sl_comments",
        "sl_post_multi_issuer",
        "sl_comment_multi_issuer",
        "samples_sl_multi_issuer",
    ]:
        df = sl.get(key) if isinstance(sl, dict) else None
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
                distinct_texts=("distinct_texts", "sum"),
                distinct_issuers=("distinct_issuers", "max"),
                distinct_offerings=("distinct_offerings", "max"),
                sources=("source", lambda s: ", ".join(sorted(set(s)))),
            )
            .sort_values(["distinct_issuers", "distinct_offerings", "rows"], ascending=[False, False, False])
        )
        alias_ambiguity["alias_len"] = alias_ambiguity["alias"].astype(str).str.len()
        alias_ambiguity["flag_suspicious"] = (
            (alias_ambiguity["distinct_issuers"] >= 3)
            | (alias_ambiguity["distinct_offerings"] >= 5)
            | ((alias_ambiguity["alias_len"] <= 4) & (alias_ambiguity["rows"] >= 20))
        )
        alias_ambiguity.to_csv(args.out_dir / "alias_ambiguity.csv", index=False, encoding="utf-8-sig")

    if isinstance(cov, dict) and not cov.get("missing"):
        cov["distribution"].to_csv(args.out_dir / "coverage_summary.csv", index=False, encoding="utf-8-sig")
        cov["zero_coverage"].to_csv(args.out_dir / "coverage_zero_cases.csv", index=False, encoding="utf-8-sig")
        cov["high_coverage"].to_csv(args.out_dir / "coverage_high_cases.csv", index=False, encoding="utf-8-sig")

    char_parts = []
    if isinstance(tg.get("char_budget"), pd.DataFrame) and not tg["char_budget"].empty:
        char_parts.append(tg["char_budget"])
    for key in ["post_char_budget", "comment_char_budget"]:
        if isinstance(sl.get(key), pd.DataFrame) and not sl[key].empty:
            char_parts.append(sl[key])
    if char_parts:
        char_budget = pd.concat(char_parts, ignore_index=True)
        total = pd.DataFrame([{
            "source": "total",
            "matched_texts": int(char_budget["matched_texts"].fillna(0).sum()),
            "candidate_rows": int(char_budget["candidate_rows"].fillna(0).sum()),
            "unique_chars": int(char_budget["unique_chars"].fillna(0).sum()),
            "candidate_weighted_chars": int(char_budget["candidate_weighted_chars"].fillna(0).sum()),
            "duplication_factor": (
                float(char_budget["candidate_weighted_chars"].fillna(0).sum())
                / float(char_budget["unique_chars"].fillna(0).sum())
            ),
        }])
        char_budget = pd.concat([char_budget, total], ignore_index=True)
        char_budget["unique_tokens_est_low_chars_per_4"] = char_budget["unique_chars"] / 4.0
        char_budget["unique_tokens_est_high_chars_per_2_5"] = char_budget["unique_chars"] / 2.5
        char_budget["candidate_tokens_est_low_chars_per_4"] = char_budget["candidate_weighted_chars"] / 4.0
        char_budget["candidate_tokens_est_high_chars_per_2_5"] = char_budget["candidate_weighted_chars"] / 2.5
        char_budget.to_csv(args.out_dir / "char_budget.csv", index=False, encoding="utf-8-sig")

    report = build_report(tg, sl, cov)
    (args.out_dir / "audit_report.txt").write_text(report, encoding="utf-8")
    print(report)
    print(f"\nSaved outputs to: {args.out_dir.resolve()}")


if __name__ == "__main__":
    main()
