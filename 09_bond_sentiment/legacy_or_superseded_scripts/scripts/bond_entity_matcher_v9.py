#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bond Entity Matcher v9 — match social media texts to bond issuers and offerings.

Research context
----------------
This script is part of a study on how investor sentiment affects the success of
corporate bond primary placements in Russia.  It links texts from Telegram channels
and Smart-Lab (posts + comments) to specific bond issuers and offerings so that
sentiment scores can later be computed per issuer/offering around the placement date.

What the script does
--------------------
1. Loads alias dictionaries for issuers and offerings (from xlsx files).
2. For each text, finds entity mentions using n-gram matching with consumed-position
   tracking (longest match wins; prevents "ФПК" false-matching inside "ФПК Гарант-Инвест").
3. Classifies each text's topical context:
   - bond_context:    mentions bonds, coupons, yields, spreads, etc.
   - primary_context: mentions primary placement specifics (bookbuilding, ориентир, etc.)
   - broker_noise:    mentions brokerage services (tariffs, IIS, deposits, etc.)
4. Suppresses false matches for dual-role issuers (banks that are also brokers)
   when text is about banking services, not about bonds.
5. Resolves one final entity per text. Telegram replies can inherit from parent.
6. Saves results to SQLite with topical flags for three-layer analysis:
   - primary_relevant: main regression specification
   - bond_relevant:    robustness check
   - entity_only:      baseline

Input files
-----------
- data.db:          Telegram messages (table: messages)
- smartlab_bonds.db: Smart-Lab posts and comments (tables: posts, comments)
- panel_final_v2.xlsx:  bond offerings panel
- issuer_aliases.xlsx:  issuer → alias mapping
- issue_aliases_candidates_fixed.xlsx: offering → alias mapping

Output tables (in SQLite)
-------------------------
- telegram_matches_resolved:         one row per Telegram message
- smartlab_post_matches_resolved:    one row per Smart-Lab post
- smartlab_comment_matches_resolved: one row per Smart-Lab comment
- mention_coverage CSV:              per-offering mention counts in 3 layers
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sqlite3
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd

try:
    from tqdm import tqdm
except Exception:  # pragma: no cover
    tqdm = None


# =========================
# Configuration
# =========================

# ---------------------------------------------------------------------------
# Three-tier topical keyword system
# ---------------------------------------------------------------------------
# BOND_KEYWORDS: broad bond-market context (secondary + primary)
# PRIMARY_KEYWORDS: narrow primary-placement context (subset of bond)
# BROKER_NOISE_KEYWORDS: brokerage/banking service mentions (for dual-role filter)

BOND_KEYWORDS = [
    # instrument
    "облигац", "бонд", "bond",
    # cash flows
    "купон", "погашен", "амортиз", "оферт", "дюрац",
    # valuation
    "доходност", "спред", "кредитн", "рейтинг",
    # structure
    "эмитент", "эмиссия", "бумага",
    # process (broad — includes secondary market)
    "размещ", "выпуск",
    # codes / identifiers (broad)
    "isin", "001p", "001р", "серия",
]

PRIMARY_KEYWORDS = [
    # NOTE: "размещ" intentionally excluded — too ambiguous
    # ("нашёл размещение на ИИС" ≠ bond primary placement)
    "первичк", "первичн",
    "книг",                             # stem: книга/книгу/книги (bookbuilding)
    "букбилд", "бук-билд", "бук билд",
    "сбор заяв",
    "ориентир",                         # ориентир купона — specific to primary
    "маркетинг",                        # маркетинговый период
    "новый выпуск", "новая серия",
]

BROKER_NOISE_KEYWORDS = [
    "брокер", "тариф", "комисси",
    "приложен", "терминал",
    "иис",
    "вклад", "депозит",
    "карт", "кешбэк", "кэшбэк",
    "маржа", "маржинал",
    "перевод", "вывод средств",
    "налог", "ндфл",
    "сертификат",
    "ипотек",
]

# Issuers that double as banks/brokers — broker_noise filter is especially
# important for these because many texts mention them in a service context
DUAL_ROLE_ISSUERS = {
    "Банк ВТБ (ПАО)", "Банк ПСБ", "Альфа-Банк", "Совкомбанк",
    "Газпромбанк", "Банк ДОМ.РФ", "ДОМ.РФ", "МТС-Банк",
    "Банк Открытие", "Банк ЗЕНИТ", "ТКБ", "Тинькофф Банк",
    "Сбербанк", "Сбербанк России", "РСХБ", "Россельхозбанк",
    "МКБ", "Московский кредитный банк", "Т-Банк",
}


SHORT_REPLY_STOPWORDS = {
    "+", "++", "ага", "угу", "ок", "окей", "да", "нет", "согласен",
    "не согласен", "норм", "лол", "жесть", "имхо", "ясно", "понятно",
    "спасибо", "thanks"
}

# Inherited match score decay (only depth-1 inheritance is used)
INHERIT_SCORE_DECAY = 0.70
INHERIT_SHORT_REPLY_PENALTY = 0.5
DEFAULT_MAX_TIME_GAP_HOURS = 72
DEFAULT_CHUNK_SIZE = 5000
DEFAULT_WINDOW_DAYS = 7
DEFAULT_MAX_ALIAS_WORDS = 5

STRONG_ISSUE_ALIAS_TYPES = {"isin", "bond_name_full", "issuer_plus_code", "issuer_plus_code_variant"}
CODE_ONLY_ISSUE_ALIAS_TYPES = {"code_exact", "code_variant"}


# =========================
# Data classes
# =========================

@dataclass
class ResolvedEntity:
    issuer: str
    entity_level: str  # issuer | offering
    offering_id: Optional[int] = None
    bond_name: Optional[str] = None
    isin: Optional[str] = None
    match_origin: str = "explicit"  # explicit | inherited
    matched_aliases: str = "[]"
    alias_types: str = "[]"
    match_score: float = 0.0
    resolution_note: str = ""
    inherit_anchor_key: Optional[str] = None
    inherit_depth: int = 0


@dataclass
class TopicFlags:
    """Topical classification of a text, independent of entity matching.

    Three tiers:
    - has_bond_context: text mentions bonds / coupons / yields etc.
    - has_primary_context: text mentions primary placement specifics
    - has_broker_noise: text mentions brokerage/banking services

    has_primary_context=True always implies has_bond_context=True.
    """
    has_bond_context: bool = False
    has_primary_context: bool = False
    has_broker_noise: bool = False
    bond_keywords_found: List[str] = field(default_factory=list)
    primary_keywords_found: List[str] = field(default_factory=list)
    noise_keywords_found: List[str] = field(default_factory=list)


# =========================
# Text normalization
# =========================

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
    s = s.replace("«", '"').replace("»", '"')
    s = s.replace("“", '"').replace("”", '"')
    s = s.replace("’", "'")
    # Replace punctuation with space, but preserve hyphen/dot inside codes.
    s = re.sub(r"[^a-zа-я0-9\-\.]+", " ", s, flags=re.IGNORECASE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def tokenize(text_norm: str) -> List[str]:
    return TOKEN_RE.findall(text_norm)



def classify_topic(text_norm: str) -> TopicFlags:
    """Classify topical context of *normalized* text.

    Independent of entity matching — this is a property of the text itself.
    """
    flags = TopicFlags()

    for kw in BOND_KEYWORDS:
        if kw in text_norm:
            flags.has_bond_context = True
            flags.bond_keywords_found.append(kw)

    for kw in PRIMARY_KEYWORDS:
        if kw in text_norm:
            flags.has_primary_context = True
            flags.primary_keywords_found.append(kw)
            # Primary implies bond context
            flags.has_bond_context = True

    for kw in BROKER_NOISE_KEYWORDS:
        if kw in text_norm:
            flags.has_broker_noise = True
            flags.noise_keywords_found.append(kw)

    return flags


def should_suppress_dual_role_match(
    entity: ResolvedEntity,
    topic: TopicFlags,
    issue_hits: Dict[int, Dict[str, Any]],
) -> bool:
    issuer = entity.issuer

    if issuer not in DUAL_ROLE_ISSUERS:
        return False

    # offering-level explicit evidence не suppressим
    if entity.entity_level == "offering":
        return False

    # если есть strong issue evidence по тому же issuer — не suppressим
    has_strong_issue = any(
        hit["issuer"] == issuer and hit.get("has_strong_issue_evidence", False)
        for hit in issue_hits.values()
    )
    if has_strong_issue:
        return False

    # broker-noise без primary-context -> suppress
    if topic.has_broker_noise and not topic.has_primary_context:
        return True

    return False


def parse_dt(value: Any) -> Optional[datetime]:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    s = str(value).strip()
    if not s:
        return None
    try:
        ts = pd.to_datetime(s, errors="coerce", utc=True)
        if pd.isna(ts):
            return None
        return ts.tz_convert("UTC").tz_localize(None).to_pydatetime()
    except Exception:
        return None


def hours_diff(dt1: Optional[datetime], dt2: Optional[datetime]) -> Optional[float]:
    if dt1 is None or dt2 is None:
        return None
    return abs((dt1 - dt2).total_seconds()) / 3600.0


def is_short_low_info_reply(text: str) -> bool:
    t = normalize_text(text)
    if not t:
        return True
    if t in SHORT_REPLY_STOPWORDS:
        return True
    toks = tokenize(t)
    return len(toks) <= 2


# =========================
# Alias loading
# =========================


def issuer_base_score(alias_type: str, risk_flag: str, use_rule: str) -> float:
    """Score an issuer alias based on its type and risk level.

    Actual alias_type values in our data: 'existing' (manually curated)
    and 'auto' (auto-generated from issuer name normalization).
    """
    base = {
        "existing": 0.72,      # manually curated alias
        "auto": 0.60,          # auto-generated (normalization variants)
    }.get(alias_type, 0.55)
    if risk_flag == "medium":
        base -= 0.05
    elif risk_flag == "high":
        base -= 0.10
    if use_rule == "require_context":
        base -= 0.03
    return max(base, 0.30)


def issue_base_score(alias_type: str, risk_flag: str, use_rule: str) -> float:
    base = {
        "isin": 1.00,
        "bond_name_full": 0.98,
        "issuer_plus_code": 0.92,
        "issuer_plus_code_variant": 0.88,
        "code_exact": 0.64,
        "code_variant": 0.58,
    }.get(alias_type, 0.55)
    if risk_flag == "medium":
        base -= 0.04
    elif risk_flag == "high":
        base -= 0.08
    if use_rule == "review":
        base -= 0.15
    return max(base, 0.25)




def issue_alias_is_strong(alias_type: str) -> bool:
    return alias_type in STRONG_ISSUE_ALIAS_TYPES


def issue_alias_is_code_only(alias_type: str) -> bool:
    return alias_type in CODE_ONLY_ISSUE_ALIAS_TYPES


def load_offerings(panel_path: Path) -> pd.DataFrame:
    df = pd.read_excel(panel_path)
    df = df.copy()
    df["offering_id"] = range(1, len(df) + 1)
    df["issuer_norm"] = df["issuer"].map(normalize_text)
    df["bond_name_norm"] = df["bond_name"].map(normalize_text)
    df["ISIN_norm"] = df["ISIN"].astype(str).map(normalize_text)
    df["book_date"] = pd.to_datetime(df["book_date"], errors="coerce")
    df["placement_date"] = pd.to_datetime(df["placement_date"], errors="coerce")
    return df



def alias_has_latin_or_digit(alias_norm: str) -> bool:
    return bool(re.search(r"[a-z0-9]", alias_norm))


def verify_nonlow_issuer_alias(
    meta: Dict[str, Any],
    context_flag: bool,
    issuer_has_low_risk_alias: bool,
    issuer_has_issue_hit: bool,
    total_candidate_issuers: int,
) -> bool:
    """
    Conservative verifier for medium/high-risk issuer aliases.

    Verification passes if at least one strong corroborator exists:
    - same issuer also matched a low-risk alias in the same text
    - same issuer also has an issue-level hit in the same text

    Otherwise, allow only limited context-based fallback:
    - medium risk: unique issuer candidate + bond context + non-trivial alias
    - high risk: unique issuer candidate + bond context + alias is either
      alnum/ticker-like OR comes from a more semantic alias type
      (official/manual_common/drop_group/drop_bank)
    """
    if issuer_has_low_risk_alias or issuer_has_issue_hit:
        return True

    if not context_flag or total_candidate_issuers != 1:
        return False

    alias_norm = meta["alias_norm"]
    alias_type = str(meta.get("alias_type", ""))
    risk_flag = str(meta.get("risk_flag", "low")).lower()
    long_enough = len(alias_norm) >= 4
    lat_or_digit = alias_has_latin_or_digit(alias_norm)

    if risk_flag == "medium":
        return long_enough or lat_or_digit

    if risk_flag == "high":
        safer_context_types = {
            "official",
            "manual_common",
            "drop_group",
            "drop_group_companies",
            "drop_bank",
        }
        return (lat_or_digit and len(alias_norm) >= 2) or (alias_type in safer_context_types and long_enough)

    return True



def read_alias_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(path)
    return pd.read_csv(path)


def default_issuer_use_rule(risk_flag: str) -> str:
    risk = str(risk_flag).strip().lower()
    return "verify_nonlow" if risk in {"medium", "high"} else "ok"


def build_alias_dicts(
    panel_df: pd.DataFrame,
    issuer_alias_path: Path,
    issue_alias_path: Path,
    max_alias_words: int = DEFAULT_MAX_ALIAS_WORDS,
) -> Tuple[Dict[int, Dict[str, List[Dict[str, Any]]]], Dict[int, Dict[str, List[Dict[str, Any]]]], Dict[str, List[int]]]:
    issuer_df = read_alias_table(issuer_alias_path)
    issue_df = read_alias_table(issue_alias_path)

    if "risk_flag" not in issuer_df.columns:
        issuer_df["risk_flag"] = "low"
    issuer_df["risk_flag"] = issuer_df["risk_flag"].fillna("low").astype(str).str.strip().str.lower()

    if "use_rule" not in issuer_df.columns:
        issuer_df["use_rule"] = issuer_df["risk_flag"].map(default_issuer_use_rule)
    else:
        issuer_df["use_rule"] = issuer_df["use_rule"].fillna(
            issuer_df["risk_flag"].map(default_issuer_use_rule)
        )
    issuer_df["use_rule"] = issuer_df["use_rule"].fillna("ok").astype(str).str.strip().str.lower()

    if "risk_flag" not in issue_df.columns:
        issue_df["risk_flag"] = "low"
    issue_df["risk_flag"] = issue_df["risk_flag"].fillna("low").astype(str).str.strip().str.lower()

    if "use_rule" not in issue_df.columns:
        issue_df["use_rule"] = "ok"
    issue_df["use_rule"] = issue_df["use_rule"].fillna("ok").astype(str).str.strip().str.lower()

    # Map issue aliases to offering_id from panel.
    merge_cols = ["issuer", "bond_name", "ISIN"]
    panel_key = panel_df[merge_cols + ["offering_id"]].drop_duplicates()
    issue_df = issue_df.merge(panel_key, how="left", on=merge_cols)

    missing = issue_df["offering_id"].isna().sum()
    if missing:
        print(f"[WARN] {missing} issue alias rows could not be mapped to offering_id", file=sys.stderr)
    issue_df = issue_df[issue_df["offering_id"].notna()].copy()
    issue_df["offering_id"] = issue_df["offering_id"].astype(int)

    allowed_issue_rules = {"ok", "require_issuer_or_window"}
    issue_df = issue_df[issue_df["use_rule"].isin(allowed_issue_rules)].copy()

    issuer_aliases_by_len: Dict[int, Dict[str, List[Dict[str, Any]]]] = defaultdict(lambda: defaultdict(list))
    issue_aliases_by_len: Dict[int, Dict[str, List[Dict[str, Any]]]] = defaultdict(lambda: defaultdict(list))
    issuer_to_offerings: Dict[str, List[int]] = defaultdict(list)

    for _, row in panel_df[["issuer", "offering_id"]].drop_duplicates().iterrows():
        issuer_to_offerings[row["issuer"]].append(int(row["offering_id"]))

    for _, row in issuer_df.iterrows():
        alias_norm = normalize_text(row["alias"])
        if not alias_norm:
            continue
        n = len(alias_norm.split())
        if n == 0 or n > max_alias_words:
            continue
        meta = {
            "issuer": row["issuer"],
            "alias": row["alias"],
            "alias_norm": alias_norm,
            "alias_type": row["alias_type"],
            "risk_flag": str(row["risk_flag"]).lower(),
            "use_rule": row["use_rule"],
            "base_score": issuer_base_score(row["alias_type"], str(row["risk_flag"]).lower(), row["use_rule"]),
        }
        issuer_aliases_by_len[n][alias_norm].append(meta)

    for _, row in issue_df.iterrows():
        alias_norm = normalize_text(row["alias"])
        if not alias_norm:
            continue
        n = len(alias_norm.split())
        if n == 0 or n > max_alias_words:
            continue
        meta = {
            "issuer": row["issuer"],
            "bond_name": row["bond_name"],
            "isin": row["ISIN"],
            "offering_id": int(row["offering_id"]),
            "alias": row["alias"],
            "alias_norm": alias_norm,
            "alias_type": row["alias_type"],
            "risk_flag": str(row["risk_flag"]).lower(),
            "use_rule": row["use_rule"],
            "base_score": issue_base_score(row["alias_type"], str(row["risk_flag"]).lower(), row["use_rule"]),
        }
        issue_aliases_by_len[n][alias_norm].append(meta)

    return issuer_aliases_by_len, issue_aliases_by_len, issuer_to_offerings



# =========================
# Matching logic
# =========================



def match_document(
    raw_text: str,
    issuer_aliases_by_len: Dict[int, Dict[str, List[Dict[str, Any]]]],
    issue_aliases_by_len: Dict[int, Dict[str, List[Dict[str, Any]]]],
    max_alias_words: int = DEFAULT_MAX_ALIAS_WORDS,
) -> Tuple[Dict[str, Dict[str, Any]], Dict[int, Dict[str, Any]], TopicFlags]:
    """
    Returns:
      issuer_hits: dict[issuer] -> aggregated metadata
      issue_hits:  dict[offering_id] -> aggregated metadata
      topic: TopicFlags (bond_context, primary_context, broker_noise)

    Uses position-aware consumed-token tracking: when an n-gram at token
    positions [i, i+n) matches an alias, those positions are marked as
    consumed.  Shorter n-grams that overlap any consumed position are
    skipped.  This prevents false positives where a short alias (e.g.
    "ФПК") is a substring of a longer entity name that already matched
    (e.g. "ФПК Гарант-Инвест").  If the short alias appears at a
    *different* position in the text, it still matches correctly.
    """
    text_norm = normalize_text(raw_text)
    tokens = tokenize(text_norm)
    topic = classify_topic(text_norm)
    context_flag = topic.has_bond_context  # used for require_context aliases

    issuer_hits: Dict[str, Dict[str, Any]] = {}
    issue_hits: Dict[int, Dict[str, Any]] = {}
    pending_issue_rows: List[Dict[str, Any]] = []
    tentative_nonlow_by_issuer: Dict[str, List[Dict[str, Any]]] = defaultdict(list)

    # --- Position-aware matching with consumed-token tracking ---
    consumed: set = set()          # token indices already claimed by a longer match
    num_tokens = len(tokens)
    max_n = min(max_alias_words, num_tokens)

    for n in range(max_n, 0, -1):
        issuer_dict_n = issuer_aliases_by_len.get(n, {})
        issue_dict_n = issue_aliases_by_len.get(n, {})
        if not issuer_dict_n and not issue_dict_n:
            continue
        for i in range(num_tokens - n + 1):
            # Skip if any token in this span is already consumed
            span = range(i, i + n)
            if any(p in consumed for p in span):
                continue
            gram = " ".join(tokens[i:i + n])

            issuer_metas = issuer_dict_n.get(gram, [])
            issue_metas = issue_dict_n.get(gram, [])
            if not issuer_metas and not issue_metas:
                continue

            # At least one alias matched at this span → consume the positions
            matched_any = False

            for meta in issuer_metas:
                if meta["use_rule"] == "require_context" and not context_flag:
                    continue
                matched_any = True
                issuer = meta["issuer"]
                hit = issuer_hits.setdefault(
                    issuer,
                    {
                        "issuer": issuer,
                        "matched_aliases": set(),
                        "alias_types": set(),
                        "max_score": 0.0,
                        "has_low_risk_alias": False,
                    },
                )
                if str(meta.get("risk_flag", "low")).lower() == "low":
                    hit["matched_aliases"].add(meta["alias"])
                    hit["alias_types"].add(meta["alias_type"])
                    hit["max_score"] = max(hit["max_score"], meta["base_score"])
                    hit["has_low_risk_alias"] = True
                else:
                    tentative_nonlow_by_issuer[issuer].append(meta)
                    matched_any = True  # tentative still consumes positions

            for meta in issue_metas:
                matched_any = True
                pending_issue_rows.append(meta)

            if matched_any:
                consumed.update(span)

    for meta in pending_issue_rows:
        # Unsafe short codes require issuer presence in the same text.
        if meta["use_rule"] == "require_issuer_or_window" and meta["issuer"] not in issuer_hits:
            continue
        oid = int(meta["offering_id"])
        hit = issue_hits.setdefault(
            oid,
            {
                "issuer": meta["issuer"],
                "offering_id": oid,
                "bond_name": meta["bond_name"],
                "isin": meta["isin"],
                "matched_aliases": set(),
                "alias_types": set(),
                "max_score": 0.0,
                "has_strong_issue_evidence": False,
                "has_code_only_issue_evidence": False,
            },
        )
        hit["matched_aliases"].add(meta["alias"])
        hit["alias_types"].add(meta["alias_type"])
        hit["max_score"] = max(hit["max_score"], meta["base_score"])
        if issue_alias_is_strong(meta["alias_type"]):
            hit["has_strong_issue_evidence"] = True
        if issue_alias_is_code_only(meta["alias_type"]):
            hit["has_code_only_issue_evidence"] = True

    total_candidate_issuers = len(issuer_hits)

    # Verify medium/high-risk issuer aliases only after seeing the full document.
    for issuer, metas in tentative_nonlow_by_issuer.items():
        hit = issuer_hits.setdefault(
            issuer,
            {
                "issuer": issuer,
                "matched_aliases": set(),
                "alias_types": set(),
                "max_score": 0.0,
                "has_low_risk_alias": False,
            },
        )
        issuer_has_issue_hit = any(ih["issuer"] == issuer for ih in issue_hits.values())
        for meta in metas:
            if verify_nonlow_issuer_alias(
                meta=meta,
                context_flag=context_flag,
                issuer_has_low_risk_alias=bool(hit.get("has_low_risk_alias")),
                issuer_has_issue_hit=issuer_has_issue_hit,
                total_candidate_issuers=total_candidate_issuers,
            ):
                hit["matched_aliases"].add(meta["alias"])
                hit["alias_types"].add(meta["alias_type"])
                hit["max_score"] = max(hit["max_score"], meta["base_score"])

    # Drop empty issuer hits that were created only by unverified tentative aliases.
    issuer_hits = {
        issuer: {
            "issuer": hit["issuer"],
            "matched_aliases": hit["matched_aliases"],
            "alias_types": hit["alias_types"],
            "max_score": hit["max_score"],
        }
        for issuer, hit in issuer_hits.items()
        if hit["matched_aliases"]
    }

    return issuer_hits, issue_hits, topic




def resolve_explicit_entity(
    issuer_hits: Dict[str, Dict[str, Any]],
    issue_hits: Dict[int, Dict[str, Any]],
) -> Tuple[str, Optional[ResolvedEntity], str]:
    """
    Returns: (status, resolved_entity, note)
      status in {'resolved','ambiguous','none'}
    """
    if issue_hits:
        if len(issue_hits) == 1:
            hit = next(iter(issue_hits.values()))
            return (
                "resolved",
                ResolvedEntity(
                    issuer=hit["issuer"],
                    entity_level="offering",
                    offering_id=hit["offering_id"],
                    bond_name=hit["bond_name"],
                    isin=hit["isin"],
                    match_origin="explicit",
                    matched_aliases=json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                    alias_types=json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
                    match_score=round(hit["max_score"], 4),
                    resolution_note="unique_issue_match",
                ),
                "unique_issue_match",
            )
        return "ambiguous", None, f"ambiguous_issue_matches:{len(issue_hits)}"

    if issuer_hits:
        if len(issuer_hits) == 1:
            hit = next(iter(issuer_hits.values()))
            return (
                "resolved",
                ResolvedEntity(
                    issuer=hit["issuer"],
                    entity_level="issuer",
                    match_origin="explicit",
                    matched_aliases=json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                    alias_types=json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
                    match_score=round(hit["max_score"], 4),
                    resolution_note="unique_issuer_match",
                ),
                "unique_issuer_match",
            )
        return "ambiguous", None, f"ambiguous_issuer_matches:{len(issuer_hits)}"

    return "none", None, "no_match"


# =========================
# SQLite helpers
# =========================


def drop_output_tables(conn: sqlite3.Connection, tables: Iterable[str]) -> None:
    cur = conn.cursor()
    for tbl in tables:
        cur.execute(f"DROP TABLE IF EXISTS {tbl}")
    conn.commit()



def create_output_tables_telegram(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS telegram_match_candidates (
            source_key TEXT,
            channel TEXT,
            message_id INTEGER,
            parent_id INTEGER,
            published_at TEXT,
            entity_type TEXT,
            issuer TEXT,
            offering_id INTEGER,
            bond_name TEXT,
            isin TEXT,
            matched_aliases TEXT,
            alias_types TEXT,
            match_score REAL,
            note TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS telegram_matches_resolved (
            source_key TEXT PRIMARY KEY,
            channel TEXT,
            message_id INTEGER,
            parent_id INTEGER,
            published_at TEXT,
            issuer TEXT,
            entity_level TEXT,
            offering_id INTEGER,
            bond_name TEXT,
            isin TEXT,
            match_origin TEXT,
            matched_aliases TEXT,
            alias_types TEXT,
            match_score REAL,
            resolution_note TEXT,
            inherit_anchor_key TEXT,
            inherit_depth INTEGER,
            text_len INTEGER,
            is_short_reply INTEGER,
            explicit_status TEXT,
            bond_context INTEGER DEFAULT 0,
            primary_context INTEGER DEFAULT 0,
            broker_noise INTEGER DEFAULT 0
        )
        """
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_tg_resolved_issuer ON telegram_matches_resolved(issuer)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_tg_resolved_offering ON telegram_matches_resolved(offering_id)")
    conn.commit()



def create_output_tables_smartlab(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS smartlab_post_match_candidates (
            source_key TEXT,
            post_id TEXT,
            published_at TEXT,
            entity_type TEXT,
            issuer TEXT,
            offering_id INTEGER,
            bond_name TEXT,
            isin TEXT,
            matched_aliases TEXT,
            alias_types TEXT,
            match_score REAL,
            note TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS smartlab_post_matches_resolved (
            source_key TEXT PRIMARY KEY,
            post_id TEXT,
            published_at TEXT,
            issuer TEXT,
            entity_level TEXT,
            offering_id INTEGER,
            bond_name TEXT,
            isin TEXT,
            match_origin TEXT,
            matched_aliases TEXT,
            alias_types TEXT,
            match_score REAL,
            resolution_note TEXT,
            inherit_anchor_key TEXT,
            inherit_depth INTEGER,
            explicit_status TEXT,
            bond_context INTEGER DEFAULT 0,
            primary_context INTEGER DEFAULT 0,
            broker_noise INTEGER DEFAULT 0
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS smartlab_comment_match_candidates (
            source_key TEXT,
            comment_id TEXT,
            post_id TEXT,
            published_at TEXT,
            entity_type TEXT,
            issuer TEXT,
            offering_id INTEGER,
            bond_name TEXT,
            isin TEXT,
            matched_aliases TEXT,
            alias_types TEXT,
            match_score REAL,
            note TEXT
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS smartlab_comment_matches_resolved (
            source_key TEXT PRIMARY KEY,
            comment_id TEXT,
            post_id TEXT,
            published_at TEXT,
            issuer TEXT,
            entity_level TEXT,
            offering_id INTEGER,
            bond_name TEXT,
            isin TEXT,
            match_origin TEXT,
            matched_aliases TEXT,
            alias_types TEXT,
            match_score REAL,
            resolution_note TEXT,
            inherit_anchor_key TEXT,
            inherit_depth INTEGER,
            explicit_status TEXT,
            bond_context INTEGER DEFAULT 0,
            primary_context INTEGER DEFAULT 0,
            broker_noise INTEGER DEFAULT 0
        )
        """
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_sl_post_resolved_issuer ON smartlab_post_matches_resolved(issuer)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_sl_post_resolved_offering ON smartlab_post_matches_resolved(offering_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_sl_comment_resolved_issuer ON smartlab_comment_matches_resolved(issuer)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_sl_comment_resolved_offering ON smartlab_comment_matches_resolved(offering_id)")
    conn.commit()


# =========================
# Processing: Telegram
# =========================


def process_telegram(
    db_path: Path,
    issuer_aliases_by_len: Dict[int, Dict[str, List[Dict[str, Any]]]],
    issue_aliases_by_len: Dict[int, Dict[str, List[Dict[str, Any]]]],
    chunk_size: int = DEFAULT_CHUNK_SIZE,
    max_gap_hours: int = DEFAULT_MAX_TIME_GAP_HOURS,
    max_alias_words: int = DEFAULT_MAX_ALIAS_WORDS,
) -> None:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    drop_output_tables(conn, ["telegram_match_candidates", "telegram_matches_resolved"])
    create_output_tables_telegram(conn)

    meta: Dict[str, Dict[str, Any]] = {}
    candidate_rows: List[Tuple[Any, ...]] = []
    resolved_rows: List[Tuple[Any, ...]] = []

    total = conn.execute("SELECT COUNT(*) FROM messages").fetchone()[0]
    query = "SELECT id, channel, text, date, parent_id FROM messages ORDER BY channel, id"

    iterator = pd.read_sql_query(query, conn, chunksize=chunk_size)
    if tqdm is not None:
        iterator = tqdm(iterator, total=max(1, math.ceil(total / chunk_size)), desc="Telegram explicit")

    for chunk in iterator:
        for _, row in chunk.iterrows():
            channel = row["channel"]
            message_id = int(row["id"])
            source_key = f"{channel}:{message_id}"
            parent_id = None if pd.isna(row["parent_id"]) else int(row["parent_id"])
            parent_key = f"{channel}:{parent_id}" if parent_id is not None else None
            published_at = row["date"]
            dt = parse_dt(published_at)
            text = row["text"] if isinstance(row["text"], str) else ""
            text_len = len(text)
            short_reply_flag = 1 if is_short_low_info_reply(text) else 0

            issuer_hits, issue_hits, topic = match_document(
                text,
                issuer_aliases_by_len,
                issue_aliases_by_len,
                max_alias_words=max_alias_words,
            )
            explicit_status, explicit_entity, note = resolve_explicit_entity(issuer_hits, issue_hits)

            # Suppress dual-role issuer matches in broker-noise texts
            if explicit_entity and should_suppress_dual_role_match(explicit_entity, topic, issue_hits):
                explicit_entity = ResolvedEntity(
                    issuer=explicit_entity.issuer,
                    entity_level=explicit_entity.entity_level,
                    offering_id=explicit_entity.offering_id,
                    bond_name=explicit_entity.bond_name,
                    isin=explicit_entity.isin,
                    match_origin=explicit_entity.match_origin,
                    matched_aliases=explicit_entity.matched_aliases,
                    alias_types=explicit_entity.alias_types,
                    match_score=explicit_entity.match_score,
                    resolution_note="suppressed_broker_noise",
                )
                note = "suppressed_broker_noise"
                explicit_status = "suppressed"
                explicit_entity = None

            # candidate rows
            for hit in issue_hits.values():
                candidate_rows.append(
                    (
                        source_key, channel, message_id, parent_id, published_at,
                        "offering", hit["issuer"], hit["offering_id"], hit["bond_name"], hit["isin"],
                        json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                        json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
                        round(hit["max_score"], 4), "explicit_issue_candidate"
                    )
                )
            for hit in issuer_hits.values():
                candidate_rows.append(
                    (
                        source_key, channel, message_id, parent_id, published_at,
                        "issuer", hit["issuer"], None, None, None,
                        json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                        json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
                        round(hit["max_score"], 4), "explicit_issuer_candidate"
                    )
                )

            meta[source_key] = {
                "channel": channel,
                "message_id": message_id,
                "parent_id": parent_id,
                "parent_key": parent_key,
                "published_at": published_at,
                "dt": dt,
                "text_len": text_len,
                "is_short_reply": short_reply_flag,
                "explicit_status": explicit_status,
                "explicit_entity": explicit_entity,
                "bond_context": int(topic.has_bond_context),
                "primary_context": int(topic.has_primary_context),
                "broker_noise": int(topic.has_broker_noise),
            }

        if candidate_rows:
            conn.executemany(
                "INSERT INTO telegram_match_candidates VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                candidate_rows,
            )
            conn.commit()
            candidate_rows.clear()

    # Resolve inheritance: simple direct-parent lookup (depth=1 only).
    def resolve_final(source_key: str) -> Optional[ResolvedEntity]:
        row = meta.get(source_key)
        if row is None:
            return None

        # Already resolved explicitly → use it
        if row["explicit_status"] == "resolved":
            return row["explicit_entity"]

        # Ambiguous → no resolution
        if row["explicit_status"] != "none":
            return None

        # Try inheriting from direct parent (Telegram reply → parent message)
        parent_key = row["parent_key"]
        if not parent_key or parent_key not in meta:
            return None

        parent_row = meta[parent_key]
        if parent_row["explicit_status"] != "resolved" or parent_row["explicit_entity"] is None:
            return None

        parent_ent = parent_row["explicit_entity"]

        # Time gap check
        gap = hours_diff(row["dt"], parent_row["dt"])
        if gap is not None and gap > max_gap_hours:
            return None

        # Short low-info replies only inherit from offering-level anchors
        if row["is_short_reply"] and parent_ent.entity_level != "offering":
            return None

        decay = INHERIT_SCORE_DECAY
        if row["is_short_reply"]:
            decay *= INHERIT_SHORT_REPLY_PENALTY

        return ResolvedEntity(
            issuer=parent_ent.issuer,
            entity_level=parent_ent.entity_level,
            offering_id=parent_ent.offering_id,
            bond_name=parent_ent.bond_name,
            isin=parent_ent.isin,
            match_origin="inherited",
            matched_aliases=parent_ent.matched_aliases,
            alias_types=parent_ent.alias_types,
            match_score=round(parent_ent.match_score * decay, 4),
            resolution_note="inherited_from_direct_explicit_parent",
            inherit_anchor_key=parent_key,
            inherit_depth=1,
        )

    keys = list(meta.keys())
    iterator2 = keys if tqdm is None else tqdm(keys, desc="Telegram resolve")
    for source_key in iterator2:
        ent = resolve_final(source_key)
        row = meta[source_key]
        topic_cols = (row["bond_context"], row["primary_context"], row["broker_noise"])
        if ent is None:
            resolved_rows.append(
                (
                    source_key, row["channel"], row["message_id"], row["parent_id"], row["published_at"],
                    None, None, None, None, None, None, "[]", "[]", None,
                    "unresolved", None, 0, row["text_len"], row["is_short_reply"], row["explicit_status"],
                    *topic_cols
                )
            )
        else:
            resolved_rows.append(
                (
                    source_key, row["channel"], row["message_id"], row["parent_id"], row["published_at"],
                    ent.issuer, ent.entity_level, ent.offering_id, ent.bond_name, ent.isin, ent.match_origin,
                    ent.matched_aliases, ent.alias_types, ent.match_score, ent.resolution_note,
                    ent.inherit_anchor_key, ent.inherit_depth, row["text_len"], row["is_short_reply"], row["explicit_status"],
                    *topic_cols
                )
            )
        if len(resolved_rows) >= 5000:
            conn.executemany(
                "INSERT OR REPLACE INTO telegram_matches_resolved VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                resolved_rows,
            )
            conn.commit()
            resolved_rows.clear()

    if resolved_rows:
        conn.executemany(
            "INSERT OR REPLACE INTO telegram_matches_resolved VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            resolved_rows,
        )
        conn.commit()

    # Print a short summary.
    stats = conn.execute(
        """
        SELECT
            COUNT(*) AS n_total,
            SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) AS n_explicit,
            SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) AS n_inherited,
            SUM(CASE WHEN entity_level='offering' THEN 1 ELSE 0 END) AS n_offering,
            SUM(CASE WHEN entity_level='issuer' THEN 1 ELSE 0 END) AS n_issuer,
            SUM(CASE WHEN issuer IS NULL THEN 1 ELSE 0 END) AS n_unresolved
        FROM telegram_matches_resolved
        """
    ).fetchone()
    print("[Telegram]", dict(stats))
    conn.close()


# =========================
# Processing: Smart-Lab
# =========================


def process_smartlab(
    db_path: Path,
    issuer_aliases_by_len: Dict[int, Dict[str, List[Dict[str, Any]]]],
    issue_aliases_by_len: Dict[int, Dict[str, List[Dict[str, Any]]]],
    chunk_size: int = DEFAULT_CHUNK_SIZE,
    max_alias_words: int = DEFAULT_MAX_ALIAS_WORDS,
    enable_comment_inheritance: bool = False,
) -> None:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    drop_output_tables(
        conn,
        [
            "smartlab_post_match_candidates",
            "smartlab_post_matches_resolved",
            "smartlab_comment_match_candidates",
            "smartlab_comment_matches_resolved",
        ],
    )
    create_output_tables_smartlab(conn)

    post_meta: Dict[str, Dict[str, Any]] = {}
    post_candidate_rows: List[Tuple[Any, ...]] = []
    post_resolved_rows: List[Tuple[Any, ...]] = []

    total_posts = conn.execute("SELECT COUNT(*) FROM posts").fetchone()[0]
    post_query = "SELECT post_id, title, tags, text, published_at FROM posts ORDER BY post_id"
    iterator = pd.read_sql_query(post_query, conn, chunksize=chunk_size)
    if tqdm is not None:
        iterator = tqdm(iterator, total=max(1, math.ceil(total_posts / chunk_size)), desc="Smart-Lab posts")

    for chunk in iterator:
        for _, row in chunk.iterrows():
            post_id = str(row["post_id"])
            source_key = f"post:{post_id}"
            full_text = " ".join(
                [
                    row["title"] if isinstance(row["title"], str) else "",
                    row["tags"] if isinstance(row["tags"], str) else "",
                    row["text"] if isinstance(row["text"], str) else "",
                ]
            )
            published_at = row["published_at"]

            issuer_hits, issue_hits, topic = match_document(
                full_text,
                issuer_aliases_by_len,
                issue_aliases_by_len,
                max_alias_words=max_alias_words,
            )
            explicit_status, explicit_entity, note = resolve_explicit_entity(issuer_hits, issue_hits)

            # Suppress dual-role issuer matches in broker-noise texts
            if explicit_entity and should_suppress_dual_role_match(explicit_entity, topic, issue_hits):
                note = "suppressed_broker_noise"
                explicit_status = "suppressed"
                explicit_entity = None

            for hit in issue_hits.values():
                post_candidate_rows.append(
                    (
                        source_key, post_id, published_at, "offering", hit["issuer"], hit["offering_id"],
                        hit["bond_name"], hit["isin"],
                        json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                        json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
                        round(hit["max_score"], 4), "explicit_issue_candidate"
                    )
                )
            for hit in issuer_hits.values():
                post_candidate_rows.append(
                    (
                        source_key, post_id, published_at, "issuer", hit["issuer"], None, None, None,
                        json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                        json.dumps(sorted(hit["alias_types"]), ensure_ascii=False),
                        round(hit["max_score"], 4), "explicit_issuer_candidate"
                    )
                )

            post_meta[source_key] = {
                "post_id": post_id,
                "published_at": published_at,
                "explicit_status": explicit_status,
                "explicit_entity": explicit_entity,
                "topic": topic,
            }

            if explicit_status == "resolved" and explicit_entity is not None:
                t = post_meta[source_key]["topic"]
                topic_cols = (int(t.has_bond_context), int(t.has_primary_context), int(t.has_broker_noise))
                post_resolved_rows.append(
                    (
                        source_key, post_id, published_at, explicit_entity.issuer, explicit_entity.entity_level,
                        explicit_entity.offering_id, explicit_entity.bond_name, explicit_entity.isin,
                        explicit_entity.match_origin, explicit_entity.matched_aliases, explicit_entity.alias_types,
                        explicit_entity.match_score, explicit_entity.resolution_note,
                        None, 0, explicit_status, *topic_cols
                    )
                )
            else:
                t = post_meta[source_key]["topic"]
                topic_cols = (int(t.has_bond_context), int(t.has_primary_context), int(t.has_broker_noise))
                post_resolved_rows.append(
                    (source_key, post_id, published_at, None, None, None, None, None, None, "[]", "[]", None, note, None, 0, explicit_status, *topic_cols)
                )

        if post_candidate_rows:
            conn.executemany("INSERT INTO smartlab_post_match_candidates VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", post_candidate_rows)
            conn.commit()
            post_candidate_rows.clear()
        if post_resolved_rows:
            conn.executemany("INSERT OR REPLACE INTO smartlab_post_matches_resolved VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", post_resolved_rows)
            conn.commit()
            post_resolved_rows.clear()

    # Comments
    comment_candidate_rows: List[Tuple[Any, ...]] = []
    comment_resolved_rows: List[Tuple[Any, ...]] = []
    total_comments = conn.execute("SELECT COUNT(*) FROM comments").fetchone()[0]
    comment_query = "SELECT comment_id, post_id, published_at, text FROM comments ORDER BY comment_id"
    iterator2 = pd.read_sql_query(comment_query, conn, chunksize=chunk_size)
    if tqdm is not None:
        iterator2 = tqdm(iterator2, total=max(1, math.ceil(total_comments / chunk_size)), desc="Smart-Lab comments")

    for chunk in iterator2:
        for _, row in chunk.iterrows():
            comment_id = str(row["comment_id"])
            post_id = str(row["post_id"])
            source_key = f"comment:{comment_id}"
            post_key = f"post:{post_id}"
            published_at = row["published_at"]
            text = row["text"] if isinstance(row["text"], str) else ""

            issuer_hits, issue_hits, comment_topic = match_document(
                text,
                issuer_aliases_by_len,
                issue_aliases_by_len,
                max_alias_words=max_alias_words,
            )
            explicit_status, explicit_entity, note = resolve_explicit_entity(issuer_hits, issue_hits)

            # --- Topical context inheritance from parent post ---
            # If the comment itself lacks bond/primary context but its parent
            # post has it, inherit the topical flags.  This preserves short
            # sentiment comments ("слишком дорого") under relevant posts.
            parent_topic = post_meta.get(post_key, {}).get("topic")
            if parent_topic is not None:
                if not comment_topic.has_bond_context and parent_topic.has_bond_context:
                    comment_topic.has_bond_context = True
                if not comment_topic.has_primary_context and parent_topic.has_primary_context:
                    comment_topic.has_primary_context = True

            # Suppress dual-role issuer matches in broker-noise texts
            if explicit_entity and should_suppress_dual_role_match(explicit_entity, comment_topic, issue_hits):
                note = "suppressed_broker_noise"
                explicit_status = "suppressed"
                explicit_entity = None

            topic_cols = (int(comment_topic.has_bond_context), int(comment_topic.has_primary_context), int(comment_topic.has_broker_noise))

            for hit in issue_hits.values():
                comment_candidate_rows.append(
                    (
                        source_key, comment_id, post_id, published_at, "offering", hit["issuer"], hit["offering_id"],
                        hit["bond_name"], hit["isin"], json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                        json.dumps(sorted(hit["alias_types"]), ensure_ascii=False), round(hit["max_score"], 4),
                        "explicit_issue_candidate"
                    )
                )
            for hit in issuer_hits.values():
                comment_candidate_rows.append(
                    (
                        source_key, comment_id, post_id, published_at, "issuer", hit["issuer"], None, None, None,
                        json.dumps(sorted(hit["matched_aliases"]), ensure_ascii=False),
                        json.dumps(sorted(hit["alias_types"]), ensure_ascii=False), round(hit["max_score"], 4),
                        "explicit_issuer_candidate"
                    )
                )

            if explicit_status == "resolved" and explicit_entity is not None:
                ent = explicit_entity
                comment_resolved_rows.append(
                    (
                        source_key, comment_id, post_id, published_at, ent.issuer, ent.entity_level, ent.offering_id,
                        ent.bond_name, ent.isin, ent.match_origin, ent.matched_aliases, ent.alias_types,
                        ent.match_score, ent.resolution_note, None, 0, explicit_status, *topic_cols
                    )
                )
            elif (
                enable_comment_inheritance
                and explicit_status == "none"
                and post_key in post_meta
                and post_meta[post_key]["explicit_status"] == "resolved"
            ):
                parent_ent = post_meta[post_key]["explicit_entity"]
                if parent_ent is not None and parent_ent.entity_level == "offering" and comment_topic.has_bond_context and len(text) >= 40:
                    inherited = ResolvedEntity(
                        issuer=parent_ent.issuer,
                        entity_level=parent_ent.entity_level,
                        offering_id=parent_ent.offering_id,
                        bond_name=parent_ent.bond_name,
                        isin=parent_ent.isin,
                        match_origin="inherited",
                        matched_aliases=parent_ent.matched_aliases,
                        alias_types=parent_ent.alias_types,
                        match_score=round(parent_ent.match_score * 0.55, 4),
                        resolution_note="inherited_from_post_offering_context",
                        inherit_anchor_key=post_key,
                        inherit_depth=1,
                    )
                    comment_resolved_rows.append(
                        (
                            source_key, comment_id, post_id, published_at, inherited.issuer, inherited.entity_level,
                            inherited.offering_id, inherited.bond_name, inherited.isin, inherited.match_origin,
                            inherited.matched_aliases, inherited.alias_types, inherited.match_score,
                            inherited.resolution_note, inherited.inherit_anchor_key, inherited.inherit_depth, explicit_status,
                            *topic_cols
                        )
                    )
                else:
                    comment_resolved_rows.append(
                        (source_key, comment_id, post_id, published_at, None, None, None, None, None, None, "[]", "[]", None, "comment_inheritance_disabled_or_failed_guardrails", None, 0, explicit_status, *topic_cols)
                    )
            else:
                comment_resolved_rows.append(
                    (source_key, comment_id, post_id, published_at, None, None, None, None, None, None, "[]", "[]", None, note, None, 0, explicit_status, *topic_cols)
                )

        if comment_candidate_rows:
            conn.executemany("INSERT INTO smartlab_comment_match_candidates VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", comment_candidate_rows)
            conn.commit()
            comment_candidate_rows.clear()
        if comment_resolved_rows:
            conn.executemany("INSERT OR REPLACE INTO smartlab_comment_matches_resolved VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", comment_resolved_rows)
            conn.commit()
            comment_resolved_rows.clear()

    post_stats = conn.execute(
        "SELECT COUNT(*) n_total, SUM(CASE WHEN issuer IS NULL THEN 1 ELSE 0 END) n_unresolved, SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) n_explicit FROM smartlab_post_matches_resolved"
    ).fetchone()
    comment_stats = conn.execute(
        "SELECT COUNT(*) n_total, SUM(CASE WHEN issuer IS NULL THEN 1 ELSE 0 END) n_unresolved, SUM(CASE WHEN match_origin='explicit' THEN 1 ELSE 0 END) n_explicit, SUM(CASE WHEN match_origin='inherited' THEN 1 ELSE 0 END) n_inherited FROM smartlab_comment_matches_resolved"
    ).fetchone()
    print("[Smart-Lab posts]", dict(post_stats))
    print("[Smart-Lab comments]", dict(comment_stats))
    conn.close()


# =========================
# Coverage aggregation
# =========================


def aggregate_mention_coverage(
    panel_df: pd.DataFrame,
    telegram_db: Path,
    smartlab_db: Path,
    out_csv: Path,
    window_days: int = DEFAULT_WINDOW_DAYS,
) -> None:
    tg = sqlite3.connect(telegram_db)
    sl = sqlite3.connect(smartlab_db)

    base_cols = "published_at, issuer, offering_id, entity_level, match_origin, bond_context, primary_context"

    tg_df = pd.read_sql_query(
        f"SELECT {base_cols} FROM telegram_matches_resolved WHERE issuer IS NOT NULL",
        tg,
    )
    sl_posts = pd.read_sql_query(
        f"SELECT {base_cols} FROM smartlab_post_matches_resolved WHERE issuer IS NOT NULL",
        sl,
    )
    sl_comments = pd.read_sql_query(
        f"SELECT {base_cols} FROM smartlab_comment_matches_resolved WHERE issuer IS NOT NULL",
        sl,
    )
    tg.close()
    sl.close()

    for df in [tg_df, sl_posts, sl_comments]:
        df["published_at"] = pd.to_datetime(df["published_at"], errors="coerce", utc=True).dt.tz_localize(None)
        # Ensure int columns (may be read as float if NULLs present)
        for col in ["bond_context", "primary_context"]:
            df[col] = df[col].fillna(0).astype(int)

    def _count_layer(df_slice: pd.DataFrame) -> Dict[str, int]:
        """Count mentions at three topical layers for a given entity/window slice."""
        n = len(df_slice)
        n_bond = int(df_slice["bond_context"].sum()) if n else 0
        n_primary = int(df_slice["primary_context"].sum()) if n else 0
        return {"entity_only": n, "bond": n_bond, "primary": n_primary}

    rows = []
    iterator = panel_df.itertuples(index=False)
    if tqdm is not None:
        iterator = tqdm(list(panel_df.itertuples(index=False)), desc="Coverage aggregation")

    for row in iterator:
        ref_date = row.book_date if pd.notna(row.book_date) else row.placement_date
        if pd.isna(ref_date):
            continue
        ref_date = pd.Timestamp(ref_date).tz_localize(None) if getattr(pd.Timestamp(ref_date), "tzinfo", None) is not None else pd.Timestamp(ref_date)
        start = ref_date - pd.Timedelta(days=window_days)
        end = ref_date

        tg_win = tg_df[(tg_df["published_at"] >= start) & (tg_df["published_at"] <= end)]
        slp_win = sl_posts[(sl_posts["published_at"] >= start) & (sl_posts["published_at"] <= end)]
        slc_win = sl_comments[(sl_comments["published_at"] >= start) & (sl_comments["published_at"] <= end)]

        tg_issue = tg_win[(tg_win["offering_id"] == row.offering_id)]
        tg_issuer = tg_win[(tg_win["offering_id"].isna()) & (tg_win["issuer"] == row.issuer)]
        slp_issue = slp_win[(slp_win["offering_id"] == row.offering_id)]
        slp_issuer = slp_win[(slp_win["offering_id"].isna()) & (slp_win["issuer"] == row.issuer)]
        slc_issue = slc_win[(slc_win["offering_id"] == row.offering_id)]
        slc_issuer = slc_win[(slc_win["offering_id"].isna()) & (slc_win["issuer"] == row.issuer)]

        # Three-layer counts per source
        tg_c = _count_layer(pd.concat([tg_issue, tg_issuer], ignore_index=True))
        slp_c = _count_layer(pd.concat([slp_issue, slp_issuer], ignore_index=True))
        slc_c = _count_layer(pd.concat([slc_issue, slc_issuer], ignore_index=True))

        rec = {
            "offering_id": row.offering_id,
            "issuer": row.issuer,
            "bond_name": row.bond_name,
            "ISIN": row.ISIN,
            "book_date": row.book_date,
            "placement_date": row.placement_date,
            "window_days": window_days,
        }

        # Telegram
        rec["tg_entity_only"] = tg_c["entity_only"]
        rec["tg_bond"] = tg_c["bond"]
        rec["tg_primary"] = tg_c["primary"]

        # Smart-Lab posts
        rec["sl_post_entity_only"] = slp_c["entity_only"]
        rec["sl_post_bond"] = slp_c["bond"]
        rec["sl_post_primary"] = slp_c["primary"]

        # Smart-Lab comments
        rec["sl_comment_entity_only"] = slc_c["entity_only"]
        rec["sl_comment_bond"] = slc_c["bond"]
        rec["sl_comment_primary"] = slc_c["primary"]

        # Aggregates across all sources
        rec["all_entity_only"] = tg_c["entity_only"] + slp_c["entity_only"] + slc_c["entity_only"]
        rec["all_bond"] = tg_c["bond"] + slp_c["bond"] + slc_c["bond"]
        rec["all_primary"] = tg_c["primary"] + slp_c["primary"] + slc_c["primary"]

        rows.append(rec)

    out_df = pd.DataFrame(rows)
    out_df.to_csv(out_csv, index=False, encoding="utf-8-sig")
    print(f"[Coverage] saved to {out_csv}")


# =========================
# Main
# =========================


def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Match Telegram/Smart-Lab texts to bond issuers and offerings.",
        epilog="Outputs: resolved match tables in SQLite + mention_coverage CSV.",
    )
    p.add_argument("--telegram-db", default="data.db", help="Path to Telegram DB")
    p.add_argument("--smartlab-db", default="smartlab_bonds.db", help="Path to Smart-Lab DB")
    p.add_argument("--panel", default="panel_final_v2.xlsx", help="Path to panel data")
    p.add_argument("--issuer-aliases", default="issuer_aliases.xlsx", help="Path to issuer aliases")
    p.add_argument("--issue-aliases", default="issue_aliases_candidates_fixed.xlsx", help="Path to issue aliases")
    p.add_argument("--window-days", type=int, default=DEFAULT_WINDOW_DAYS, help="Coverage window in days before book_date")
    p.add_argument("--enable-smartlab-comment-inherit", action="store_true", help="Enable inheritance from post to comment on Smart-Lab")
    p.add_argument("--skip-coverage", action="store_true", help="Skip building mention coverage CSV")
    p.add_argument("--coverage-out", default="mention_coverage_window7.csv", help="Output CSV for coverage")
    return p



def main() -> None:
    args = build_arg_parser().parse_args()

    telegram_db = Path(args.telegram_db)
    smartlab_db = Path(args.smartlab_db)
    panel_path = Path(args.panel)
    issuer_aliases_path = Path(args.issuer_aliases)
    issue_aliases_path = Path(args.issue_aliases)

    for path in [telegram_db, smartlab_db, panel_path, issuer_aliases_path, issue_aliases_path]:
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")

    panel_df = load_offerings(panel_path)
    issuer_aliases_by_len, issue_aliases_by_len, _ = build_alias_dicts(
        panel_df=panel_df,
        issuer_alias_path=issuer_aliases_path,
        issue_alias_path=issue_aliases_path,
    )

    print("[INFO] Running Telegram matcher...")
    process_telegram(
        telegram_db,
        issuer_aliases_by_len,
        issue_aliases_by_len,
    )

    print("[INFO] Running Smart-Lab matcher...")
    process_smartlab(
        smartlab_db,
        issuer_aliases_by_len,
        issue_aliases_by_len,
        enable_comment_inheritance=args.enable_smartlab_comment_inherit,
    )

    if not args.skip_coverage:
        aggregate_mention_coverage(
            panel_df=panel_df,
            telegram_db=telegram_db,
            smartlab_db=smartlab_db,
            out_csv=Path(args.coverage_out),
            window_days=args.window_days,
        )

    print("[DONE]")


if __name__ == "__main__":
    main()
