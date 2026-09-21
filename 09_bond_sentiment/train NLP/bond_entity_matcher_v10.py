#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bond Entity Matcher v10 — multi-match, LLM-ready.

Philosophy
----------
This matcher does ONE thing: find entity mentions in text.
It saves ALL matches — no disambiguation, no filtering, no relevance checks.

A text mentioning "ВТБ" and "Сбер" produces two rows.
A text about "депозит ВТБ" still produces a row for ВТБ.
A text where "ВТБ is agent for Самолёт placement" produces rows for both.

The LLM pipeline downstream will determine for each (text, issuer) pair:
  1. Is the text relevant to this issuer's bonds?
  2. What is the sentiment?
  3. Is this about primary placement?

What the matcher handles
------------------------
- Alias lookup with consumed-position tracking (longest match wins)
- Russian morphology via pymorphy3 lemmatization (сбера→сбер)
- Telegram reply inheritance (reply gets all parent matches)
- Smart-Lab comment inheritance (optional)
- Coverage aggregation per offering

What the matcher does NOT handle (deferred to LLM)
---------------------------------------------------
- Is this about bonds or about banking services?
- Is the issuer the subject or just mentioned in passing?
- Is this about primary placement or secondary market?
- What is the sentiment?
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sqlite3
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

try:
    from tqdm import tqdm
except Exception:
    tqdm = None

# ---------------------------------------------------------------------------
# Lemmatizer: resolves Russian morphology (сбера→сбер, газпрома→газпром)
# Install: pip install pymorphy3
# If not installed, matcher works but won't handle case forms.
# ---------------------------------------------------------------------------
try:
    import pymorphy3
    _morph = pymorphy3.MorphAnalyzer()

    def lemmatize(token: str) -> str:
        return _morph.parse(token)[0].normal_form

    HAS_LEMMATIZER = True
    print("[INFO] pymorphy3 loaded — lemmatization enabled")
except ImportError:
    def lemmatize(token: str) -> str:
        return token

    HAS_LEMMATIZER = False
    print("[WARN] pymorphy3 not installed — no lemmatization (pip install pymorphy3)")


# ============================================================================
#  Configuration
# ============================================================================

SHORT_REPLY_STOPWORDS = {
    "+", "++", "ага", "угу", "ок", "окей", "да", "нет", "согласен",
    "не согласен", "норм", "лол", "жесть", "имхо", "ясно", "понятно",
    "спасибо", "thanks",
}

INHERIT_SCORE_DECAY = 0.70
MAX_TIME_GAP_HOURS = 72
CHUNK_SIZE = 5000
MAX_ALIAS_WORDS = 5
DEFAULT_WINDOW_DAYS = 7

STRONG_ISSUE_TYPES = {"isin", "bond_name_full", "issuer_plus_code", "issuer_plus_code_variant"}
CODE_ONLY_ISSUE_TYPES = {"code_exact", "code_variant"}


# ============================================================================
#  Text processing
# ============================================================================

TOKEN_RE = re.compile(r"[a-zа-я0-9]+(?:[-.][a-zа-я0-9]+)*", re.IGNORECASE)


def normalize_text(text: Any) -> str:
    if text is None:
        return ""
    s = str(text)
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", s)
    s = s.lower().replace("ё", "е")
    s = s.replace("—", "-").replace("–", "-")
    s = s.replace("«", '"').replace("»", '"').replace("\u201c", '"').replace("\u201d", '"')
    s = re.sub(r"[^a-zа-я0-9\-\.]+", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def tokenize(text_norm: str) -> List[str]:
    """Tokenize and optionally lemmatize.

    Applied symmetrically: both alias tokens (at dict build time) and
    text tokens (at match time) go through this function.  If pymorphy3
    is installed, "сбера" and "сбер" both become "сбер" → automatic match.
    """
    tokens = TOKEN_RE.findall(text_norm)
    if HAS_LEMMATIZER:
        tokens = [lemmatize(t) for t in tokens]
    return tokens



def parse_dt(value: Any) -> Optional[datetime]:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    try:
        ts = pd.to_datetime(str(value).strip(), errors="coerce", utc=True)
        if pd.isna(ts):
            return None
        return ts.tz_convert("UTC").tz_localize(None).to_pydatetime()
    except Exception:
        return None


def hours_diff(dt1: Optional[datetime], dt2: Optional[datetime]) -> Optional[float]:
    if dt1 is None or dt2 is None:
        return None
    return abs((dt1 - dt2).total_seconds()) / 3600.0


def is_short_low_info(text: str) -> bool:
    t = normalize_text(text)
    if not t or t in SHORT_REPLY_STOPWORDS:
        return True
    return len(tokenize(t)) <= 2


# ============================================================================
#  Alias loading
# ============================================================================

def read_alias_table(path: Path) -> pd.DataFrame:
    return pd.read_excel(path) if path.suffix.lower() in {".xlsx", ".xls"} else pd.read_csv(path)


def load_offerings(panel_path: Path) -> pd.DataFrame:
    df = pd.read_excel(panel_path).copy()
    df["offering_id"] = range(1, len(df) + 1)
    df["issuer_norm"] = df["issuer"].map(normalize_text)
    df["book_date"] = pd.to_datetime(df["book_date"], errors="coerce")
    df["placement_date"] = pd.to_datetime(df["placement_date"], errors="coerce")
    return df


def build_alias_dicts(
    panel_df: pd.DataFrame,
    issuer_alias_path: Path,
    issue_alias_path: Path,
) -> Tuple[Dict, Dict]:
    """Build {n_tokens: {gram: [meta, ...]}} dicts for issuers and issues."""
    issuer_df = read_alias_table(issuer_alias_path)
    issue_df = read_alias_table(issue_alias_path)

    # Defaults
    for df, col, default in [
        (issuer_df, "risk_flag", "low"),
        (issue_df, "risk_flag", "low"),
        (issue_df, "use_rule", "ok"),
    ]:
        if col not in df.columns:
            df[col] = default
        df[col] = df[col].fillna(default).astype(str).str.strip().str.lower()

    # Derive use_rule for issuers from risk_flag
    if "use_rule" not in issuer_df.columns:
        issuer_df["use_rule"] = issuer_df["risk_flag"].map(
            lambda r: "verify_nonlow" if r in {"medium", "high"} else "ok"
        )
    issuer_df["use_rule"] = issuer_df["use_rule"].fillna("ok").astype(str).str.strip().str.lower()

    # Map issue aliases → offering_id
    merge_cols = ["issuer", "bond_name", "ISIN"]
    panel_key = panel_df[merge_cols + ["offering_id"]].drop_duplicates()
    issue_df = issue_df.merge(panel_key, how="left", on=merge_cols)
    issue_df = issue_df[issue_df["offering_id"].notna()].copy()
    issue_df["offering_id"] = issue_df["offering_id"].astype(int)
    issue_df = issue_df[issue_df["use_rule"].isin({"ok", "require_issuer_or_window"})].copy()

    issuer_by_len: Dict[int, Dict[str, List[Dict]]] = defaultdict(lambda: defaultdict(list))
    issue_by_len: Dict[int, Dict[str, List[Dict]]] = defaultdict(lambda: defaultdict(list))

    for _, row in issuer_df.iterrows():
        an = normalize_text(row["alias"])
        if not an:
            continue
        n = len(an.split())
        if n == 0 or n > MAX_ALIAS_WORDS:
            continue
        issuer_by_len[n][an].append({
            "issuer": row["issuer"], "alias": row["alias"], "alias_norm": an,
            "alias_type": row.get("alias_type", "existing"),
            "risk_flag": str(row["risk_flag"]).lower(),
            "use_rule": row["use_rule"],
        })

    for _, row in issue_df.iterrows():
        an = normalize_text(row["alias"])
        if not an:
            continue
        n = len(an.split())
        if n == 0 or n > MAX_ALIAS_WORDS:
            continue
        issue_by_len[n][an].append({
            "issuer": row["issuer"], "bond_name": row["bond_name"],
            "isin": row["ISIN"], "offering_id": int(row["offering_id"]),
            "alias": row["alias"], "alias_norm": an,
            "alias_type": row.get("alias_type", ""),
            "risk_flag": str(row["risk_flag"]).lower(),
            "use_rule": row["use_rule"],
        })

    return issuer_by_len, issue_by_len


# ============================================================================
#  Core matching — returns ALL hits, no disambiguation
# ============================================================================

def match_document(
    raw_text: str,
    issuer_by_len: Dict,
    issue_by_len: Dict,
) -> Tuple[Dict[str, Dict], Dict[int, Dict]]:
    """Find all issuer and issue mentions in text.

    Uses consumed-position tracking: longest alias match wins at each
    position.  "ФПК" won't match inside "ФПК Гарант-Инвест".

    Returns:
      issuer_hits: {issuer_name: {issuer, matched_aliases, alias_types, ...}}
      issue_hits:  {offering_id: {issuer, offering_id, bond_name, ...}}
    """
    text_norm = normalize_text(raw_text)
    tokens = tokenize(text_norm)

    issuer_hits: Dict[str, Dict] = {}
    issue_hits: Dict[int, Dict] = {}
    pending_issues: List[Dict] = []

    # --- Position-aware matching with consumed-token tracking ---
    consumed: set = set()
    num_tokens = len(tokens)
    max_n = min(MAX_ALIAS_WORDS, num_tokens)

    for n in range(max_n, 0, -1):
        i_dict = issuer_by_len.get(n, {})
        s_dict = issue_by_len.get(n, {})
        if not i_dict and not s_dict:
            continue
        for i in range(num_tokens - n + 1):
            if any(p in consumed for p in range(i, i + n)):
                continue
            gram = " ".join(tokens[i:i + n])

            i_metas = i_dict.get(gram, [])
            s_metas = s_dict.get(gram, [])
            if not i_metas and not s_metas:
                continue

            matched = False

            for meta in i_metas:
                # Relaxed verification for medium/high risk:
                # include if alias is >= 4 chars or has latin/digits
                risk = meta["risk_flag"]
                if risk in ("medium", "high"):
                    an = meta["alias_norm"]
                    if len(an) < 4 and not re.search(r"[a-z0-9]", an):
                        continue

                matched = True
                issuer = meta["issuer"]
                hit = issuer_hits.setdefault(issuer, {
                    "issuer": issuer, "matched_aliases": set(),
                    "alias_types": set(), "entity_level": "issuer",
                })
                hit["matched_aliases"].add(meta["alias"])
                hit["alias_types"].add(meta["alias_type"])

            for meta in s_metas:
                matched = True
                pending_issues.append(meta)

            if matched:
                consumed.update(range(i, i + n))

    # Process issue aliases — require_issuer_or_window needs issuer in same text
    for meta in pending_issues:
        if meta["use_rule"] == "require_issuer_or_window" and meta["issuer"] not in issuer_hits:
            continue
        oid = int(meta["offering_id"])
        hit = issue_hits.setdefault(oid, {
            "issuer": meta["issuer"], "offering_id": oid,
            "bond_name": meta["bond_name"], "isin": meta["isin"],
            "matched_aliases": set(), "alias_types": set(),
            "entity_level": "offering",
        })
        hit["matched_aliases"].add(meta["alias"])
        hit["alias_types"].add(meta["alias_type"])

    return issuer_hits, issue_hits


def _hits_to_rows(
    issuer_hits: Dict, issue_hits: Dict,
    source_key: str, channel_or_id: str, msg_id: Any,
    parent_id: Any, published_at: Any, text_len: int,
    match_origin: str = "explicit",
    inherit_anchor: str = None,
) -> List[Tuple]:
    """Convert match dicts into flat row tuples for SQL INSERT."""
    rows = []

    # Issue-level hits (higher priority — include the issuer even if also in issuer_hits)
    issuers_with_issue = set()
    for hit in issue_hits.values():
        issuers_with_issue.add(hit["issuer"])
        rows.append((
            source_key, channel_or_id, msg_id, parent_id, published_at,
            hit["issuer"], "offering", hit["offering_id"], hit["bond_name"], hit["isin"],
            match_origin,
            json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
            json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
            text_len, inherit_anchor,
        ))

    # Issuer-level hits (skip if already covered by issue-level for same issuer)
    for hit in issuer_hits.values():
        if hit["issuer"] in issuers_with_issue:
            continue
        rows.append((
            source_key, channel_or_id, msg_id, parent_id, published_at,
            hit["issuer"], "issuer", None, None, None,
            match_origin,
            json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
            json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
            text_len, inherit_anchor,
        ))

    return rows


# ============================================================================
#  SQLite schema
# ============================================================================

def create_telegram_tables(conn: sqlite3.Connection) -> None:
    conn.execute("DROP TABLE IF EXISTS telegram_matches")
    conn.execute("""
        CREATE TABLE telegram_matches (
            source_key TEXT,
            channel TEXT,
            message_id INTEGER,
            parent_id INTEGER,
            published_at TEXT,
            issuer TEXT NOT NULL,
            entity_level TEXT,
            offering_id INTEGER,
            bond_name TEXT,
            isin TEXT,
            match_origin TEXT,
            matched_aliases TEXT,
            alias_types TEXT,
            text_len INTEGER,
            inherit_anchor_key TEXT
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_sk ON telegram_matches(source_key)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_issuer ON telegram_matches(issuer)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_tg_oid ON telegram_matches(offering_id)")
    conn.commit()


def create_smartlab_tables(conn: sqlite3.Connection) -> None:
    for suffix in ["post", "comment"]:
        conn.execute(f"DROP TABLE IF EXISTS smartlab_{suffix}_matches")
        id_col = "post_id TEXT" if suffix == "post" else "comment_id TEXT, post_id TEXT"
        conn.execute(f"""
            CREATE TABLE smartlab_{suffix}_matches (
                source_key TEXT,
                {id_col},
                published_at TEXT,
                issuer TEXT NOT NULL,
                entity_level TEXT,
                offering_id INTEGER,
                bond_name TEXT,
                isin TEXT,
                match_origin TEXT,
                matched_aliases TEXT,
                alias_types TEXT,
                text_len INTEGER,
                inherit_anchor_key TEXT
            )
        """)
        conn.execute(f"CREATE INDEX IF NOT EXISTS idx_sl_{suffix}_sk ON smartlab_{suffix}_matches(source_key)")
        conn.execute(f"CREATE INDEX IF NOT EXISTS idx_sl_{suffix}_issuer ON smartlab_{suffix}_matches(issuer)")
    conn.commit()


# ============================================================================
#  Process Telegram
# ============================================================================

def process_telegram(
    db_path: Path,
    issuer_by_len: Dict,
    issue_by_len: Dict,
) -> None:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    create_telegram_tables(conn)

    # Pass 1: explicit matching — store meta for inheritance
    meta: Dict[str, Dict] = {}  # source_key → {dt, parent_key, hits...}
    batch: List[Tuple] = []
    total = conn.execute("SELECT COUNT(*) FROM messages").fetchone()[0]
    query = "SELECT id, channel, text, date, parent_id FROM messages ORDER BY channel, id"

    iterator = pd.read_sql_query(query, conn, chunksize=CHUNK_SIZE)
    if tqdm is not None:
        iterator = tqdm(iterator, total=max(1, math.ceil(total / CHUNK_SIZE)), desc="Telegram match")

    for chunk in iterator:
        for _, row in chunk.iterrows():
            channel = row["channel"]
            msg_id = int(row["id"])
            source_key = f"{channel}:{msg_id}"
            parent_id = None if pd.isna(row["parent_id"]) else int(row["parent_id"])
            parent_key = f"{channel}:{parent_id}" if parent_id is not None else None
            published_at = row["date"]
            dt = parse_dt(published_at)
            text = row["text"] if isinstance(row["text"], str) else ""
            text_len = len(text)

            issuer_hits, issue_hits = match_document(text, issuer_by_len, issue_by_len)

            rows = _hits_to_rows(
                issuer_hits, issue_hits,
                source_key, channel, msg_id, parent_id, published_at,
                text_len,
            )
            batch.extend(rows)

            meta[source_key] = {
                "parent_key": parent_key,
                "dt": dt,
                "published_at": published_at,
                "channel": channel,
                "msg_id": msg_id,
                "parent_id": parent_id,
                "text_len": text_len,
                "has_explicit": len(rows) > 0,
                "issuer_hits": issuer_hits,
                "issue_hits": issue_hits,
                "is_short": is_short_low_info(text),
            }

        if len(batch) >= 5000:
            conn.executemany(
                "INSERT INTO telegram_matches VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                batch,
            )
            conn.commit()
            batch.clear()

    if batch:
        conn.executemany(
            "INSERT INTO telegram_matches VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            batch,
        )
        conn.commit()
        batch.clear()

    # Pass 2: inheritance — replies inherit ALL parent matches
    inherit_batch: List[Tuple] = []
    keys = list(meta.keys())
    if tqdm is not None:
        keys = tqdm(keys, desc="Telegram inherit")

    for source_key in keys:
        m = meta[source_key]
        if m["has_explicit"]:
            continue  # already has own matches
        parent_key = m["parent_key"]
        if not parent_key or parent_key not in meta:
            continue
        parent = meta[parent_key]
        if not parent["has_explicit"]:
            continue

        # Time gap check
        gap = hours_diff(m["dt"], parent["dt"])
        if gap is not None and gap > MAX_TIME_GAP_HOURS:
            continue

        # Short low-info replies only inherit if parent has offering-level match
        if m["is_short"] and not parent["issue_hits"]:
            continue

        rows = _hits_to_rows(
            parent["issuer_hits"], parent["issue_hits"],
            source_key, m["channel"], m["msg_id"], m["parent_id"],
            m["published_at"], m["text_len"],
            match_origin="inherited", inherit_anchor=parent_key,
        )
        inherit_batch.extend(rows)

        if len(inherit_batch) >= 5000:
            conn.executemany(
                "INSERT INTO telegram_matches VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                inherit_batch,
            )
            conn.commit()
            inherit_batch.clear()

    if inherit_batch:
        conn.executemany(
            "INSERT INTO telegram_matches VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            inherit_batch,
        )
        conn.commit()

    stats = conn.execute("""
        SELECT COUNT(*) n_rows,
               COUNT(DISTINCT source_key) n_texts,
               SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) n_explicit,
               SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) n_inherited,
               COUNT(DISTINCT issuer) n_issuers
        FROM telegram_matches
    """).fetchone()
    print(f"[Telegram] {dict(stats)}")
    conn.close()


# ============================================================================
#  Process Smart-Lab
# ============================================================================

def process_smartlab(
    db_path: Path,
    issuer_by_len: Dict,
    issue_by_len: Dict,
    enable_comment_inheritance: bool = False,
) -> None:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    create_smartlab_tables(conn)

    # --- Posts ---
    post_meta: Dict[str, Dict] = {}
    batch: List[Tuple] = []
    total_posts = conn.execute("SELECT COUNT(*) FROM posts").fetchone()[0]
    iterator = pd.read_sql_query(
        "SELECT post_id, title, tags, text, published_at FROM posts ORDER BY post_id",
        conn, chunksize=CHUNK_SIZE,
    )
    if tqdm is not None:
        iterator = tqdm(iterator, total=max(1, math.ceil(total_posts / CHUNK_SIZE)), desc="SL posts")

    for chunk in iterator:
        for _, row in chunk.iterrows():
            post_id = str(row["post_id"])
            source_key = f"post:{post_id}"
            full_text = " ".join(filter(None, [
                row["title"] if isinstance(row["title"], str) else "",
                row["tags"] if isinstance(row["tags"], str) else "",
                row["text"] if isinstance(row["text"], str) else "",
            ]))
            published_at = row["published_at"]
            text_len = len(full_text)

            issuer_hits, issue_hits = match_document(full_text, issuer_by_len, issue_by_len)

            for hit in list(issue_hits.values()) + [h for h in issuer_hits.values() if h["issuer"] not in {ih["issuer"] for ih in issue_hits.values()}]:
                oid = hit.get("offering_id")
                batch.append((
                    source_key, post_id, published_at,
                    hit["issuer"], hit["entity_level"], oid,
                    hit.get("bond_name"), hit.get("isin"),
                    "explicit",
                    json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                    json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
                    text_len, None,
                ))

            post_meta[source_key] = {
                "issuer_hits": issuer_hits, "issue_hits": issue_hits,
                "has_matches": bool(issuer_hits or issue_hits),
            }

        if len(batch) >= 5000:
            conn.executemany("INSERT INTO smartlab_post_matches VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", batch)
            conn.commit()
            batch.clear()

    if batch:
        conn.executemany("INSERT INTO smartlab_post_matches VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", batch)
        conn.commit()
        batch.clear()

    # --- Comments ---
    total_comments = conn.execute("SELECT COUNT(*) FROM comments").fetchone()[0]
    iterator2 = pd.read_sql_query(
        "SELECT comment_id, post_id, published_at, text FROM comments ORDER BY comment_id",
        conn, chunksize=CHUNK_SIZE,
    )
    if tqdm is not None:
        iterator2 = tqdm(iterator2, total=max(1, math.ceil(total_comments / CHUNK_SIZE)), desc="SL comments")

    for chunk in iterator2:
        for _, row in chunk.iterrows():
            comment_id = str(row["comment_id"])
            post_id = str(row["post_id"])
            source_key = f"comment:{comment_id}"
            post_key = f"post:{post_id}"
            published_at = row["published_at"]
            text = row["text"] if isinstance(row["text"], str) else ""
            text_len = len(text)

            issuer_hits, issue_hits = match_document(text, issuer_by_len, issue_by_len)

            has_explicit = bool(issuer_hits or issue_hits)
            parent_meta = post_meta.get(post_key)

            # Save explicit matches
            for hit in list(issue_hits.values()) + [h for h in issuer_hits.values() if h["issuer"] not in {ih["issuer"] for ih in issue_hits.values()}]:
                oid = hit.get("offering_id")
                batch.append((
                    source_key, comment_id, post_id, published_at,
                    hit["issuer"], hit["entity_level"], oid,
                    hit.get("bond_name"), hit.get("isin"),
                    "explicit",
                    json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                    json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
                    text_len, None,
                ))

            # Inheritance from post (if enabled and no explicit match)
            if enable_comment_inheritance and not has_explicit and parent_meta and parent_meta["has_matches"]:
                if text_len >= 40:
                    for hit in list(parent_meta["issue_hits"].values()) + list(parent_meta["issuer_hits"].values()):
                        oid = hit.get("offering_id")
                        batch.append((
                            source_key, comment_id, post_id, published_at,
                            hit["issuer"], hit["entity_level"], oid,
                            hit.get("bond_name"), hit.get("isin"),
                            "inherited",
                            json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                            json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
                            text_len, post_key,
                        ))

        if len(batch) >= 5000:
            conn.executemany("INSERT INTO smartlab_comment_matches VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", batch)
            conn.commit()
            batch.clear()

    if batch:
        conn.executemany("INSERT INTO smartlab_comment_matches VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", batch)
        conn.commit()
        batch.clear()

    for tbl in ["smartlab_post_matches", "smartlab_comment_matches"]:
        stats = conn.execute(f"SELECT COUNT(*) n_rows, COUNT(DISTINCT source_key) n_texts, COUNT(DISTINCT issuer) n_issuers FROM {tbl}").fetchone()
        print(f"[{tbl}] {dict(stats)}")
    conn.close()


# ============================================================================
#  Coverage aggregation
# ============================================================================

def aggregate_mention_coverage(
    panel_df: pd.DataFrame,
    telegram_db: Path,
    smartlab_db: Path,
    out_csv: Path,
    window_days: int = DEFAULT_WINDOW_DAYS,
) -> None:
    tg = sqlite3.connect(telegram_db)
    sl = sqlite3.connect(smartlab_db)

    cols = "published_at, issuer, offering_id, entity_level, match_origin"
    tg_df = pd.read_sql_query(f"SELECT {cols} FROM telegram_matches", tg)
    sl_posts = pd.read_sql_query(f"SELECT {cols} FROM smartlab_post_matches", sl)
    sl_comments = pd.read_sql_query(f"SELECT {cols} FROM smartlab_comment_matches", sl)
    tg.close()
    sl.close()

    for df in [tg_df, sl_posts, sl_comments]:
        df["published_at"] = pd.to_datetime(df["published_at"], errors="coerce", utc=True).dt.tz_localize(None)

    rows = []
    iterator = panel_df.itertuples(index=False)
    if tqdm is not None:
        iterator = tqdm(list(iterator), desc="Coverage")

    for row in iterator:
        ref = row.book_date if pd.notna(row.book_date) else row.placement_date
        if pd.isna(ref):
            continue
        ref = pd.Timestamp(ref)
        if ref.tzinfo is not None:
            ref = ref.tz_localize(None)
        start, end = ref - pd.Timedelta(days=window_days), ref

        def count_matches(df):
            win = df[(df["published_at"] >= start) & (df["published_at"] <= end)]
            return len(win[(win["offering_id"] == row.offering_id) | ((win["offering_id"].isna()) & (win["issuer"] == row.issuer))])

        tg_n = count_matches(tg_df)
        slp_n = count_matches(sl_posts)
        slc_n = count_matches(sl_comments)

        rows.append({
            "offering_id": row.offering_id, "issuer": row.issuer,
            "bond_name": row.bond_name, "ISIN": row.ISIN,
            "book_date": row.book_date, "placement_date": row.placement_date,
            "window_days": window_days,
            "tg_mentions": tg_n,
            "sl_post_mentions": slp_n,
            "sl_comment_mentions": slc_n,
            "all_mentions": tg_n + slp_n + slc_n,
        })

    pd.DataFrame(rows).to_csv(out_csv, index=False, encoding="utf-8-sig")
    print(f"[Coverage] saved to {out_csv}")


# ============================================================================
#  Main
# ============================================================================

def main() -> None:
    p = argparse.ArgumentParser(description="Bond entity matcher v10 — multi-match edition")
    p.add_argument("--telegram-db", default="data.db")
    p.add_argument("--smartlab-db", default="smartlab_bonds.db")
    p.add_argument("--panel", default="panel_final_v3.xlsx")
    p.add_argument("--issuer-aliases", default="issuer_aliases.xlsx")
    p.add_argument("--issue-aliases", default="issue_aliases_candidates_fixed.xlsx")
    p.add_argument("--window-days", type=int, default=DEFAULT_WINDOW_DAYS)
    p.add_argument("--enable-smartlab-comment-inherit", action="store_true")
    p.add_argument("--skip-coverage", action="store_true")
    p.add_argument("--coverage-out", default="mention_coverage_window7.csv")
    args = p.parse_args()

    paths = {
        "telegram": Path(args.telegram_db), "smartlab": Path(args.smartlab_db),
        "panel": Path(args.panel), "issuer_aliases": Path(args.issuer_aliases),
        "issue_aliases": Path(args.issue_aliases),
    }
    for name, path in paths.items():
        if not path.exists():
            raise FileNotFoundError(f"{name}: {path}")

    panel_df = load_offerings(paths["panel"])
    issuer_by_len, issue_by_len = build_alias_dicts(
        panel_df, paths["issuer_aliases"], paths["issue_aliases"],
    )

    print("[INFO] Telegram...")
    process_telegram(paths["telegram"], issuer_by_len, issue_by_len)

    print("[INFO] Smart-Lab...")
    process_smartlab(
        paths["smartlab"], issuer_by_len, issue_by_len,
        enable_comment_inheritance=args.enable_smartlab_comment_inherit,
    )

    if not args.skip_coverage:
        aggregate_mention_coverage(
            panel_df, paths["telegram"], paths["smartlab"],
            Path(args.coverage_out), args.window_days,
        )

    print("[DONE]")


if __name__ == "__main__":
    main()
