#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mine colloquial issuer aliases from already-resolved explicit matches.

Idea
----
1. Use only documents that were explicitly matched to an issuer.
2. Within those documents, look around anchor aliases that already matched.
3. Collect short candidate n-grams (1-3 tokens) that are:
   - frequent inside the issuer's explicit corpus
   - relatively rare across other issuers
4. Save candidates with examples for manual review.

This does not auto-append new aliases into the production dictionary.
It creates a review file first.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sqlite3
from collections import Counter, defaultdict
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd

TOKEN_RE = re.compile(r"[a-zа-я0-9]+(?:[-.][a-zа-я0-9]+)*", re.IGNORECASE)
GENERIC_STOP = {
    "облигац", "облигации", "выпуск", "размещ", "размещение", "первич", "книга",
    "купон", "доход", "доходность", "эмитент", "эмитента", "эмитенту", "эмитентом",
    "бумага", "бумаги", "серия", "сбор", "заявок", "заявка", "рынок", "рынке",
    "оферта", "амортизац", "рейтинг", "компания", "компании", "группа", "пао",
    "ооо", "ао", "зао", "руб", "рублей", "млрд", "млн", "банк", "капитал",
    "bo", "бо", "p", "р", "001p", "001р", "002p", "002р"
}
RU_STOP = {
    "и","в","во","на","по","с","со","к","ко","у","о","об","от","до","за","из","под","над",
    "не","ни","но","а","или","ли","же","это","тот","та","те","как","что","кто","где","когда",
    "уже","еще","очень","все","всё","там","тут","для","про","под","без","при","если","то","из-за",
    "надо","будет","был","была","были","быть","есть","нет","да","ну","мы","вы","они","он","она",
    "его","ее","её","их","наш","ваш","мне","тебе","себе","меня","тебя"
}
CODE_PAT = re.compile(r"^(?:\d{2,4}[pр]-?\d{2,3}[a-zа-я]*|бо-?\d{2,3}[pр]?-?\d{0,3}|пбо-?\d{2,3}|[a-z]{2}\d+)$", re.IGNORECASE)
ISIN_PAT = re.compile(r"^[A-Z]{2}[A-Z0-9]{10}$")
CYR_OR_LAT = re.compile(r"[a-zа-я]", re.IGNORECASE)


def normalize_text(text: object) -> str:
    if text is None:
        return ""
    s = str(text).lower().replace("ё", "е").replace("—", "-").replace("–", "-")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def tokenize(text: str) -> List[str]:
    return TOKEN_RE.findall(text)


def alias_tokens(alias: str) -> Tuple[str, ...]:
    return tuple(tokenize(normalize_text(alias)))


def looks_like_candidate(gram: str) -> bool:
    g = gram.strip()
    if not g:
        return False
    if len(g) < 3:
        return False
    if ISIN_PAT.match(g.upper()):
        return False
    if CODE_PAT.match(g):
        return False
    if not CYR_OR_LAT.search(g):
        return False
    toks = g.split()
    if all(t in RU_STOP for t in toks):
        return False
    if all(t in GENERIC_STOP or t in RU_STOP for t in toks):
        return False
    if any(len(t) == 1 for t in toks):
        return False
    return True


def iter_ngrams(tokens: List[str], max_n: int = 3) -> Iterable[str]:
    m = min(max_n, len(tokens))
    for n in range(1, m + 1):
        for i in range(len(tokens) - n + 1):
            yield " ".join(tokens[i:i+n])


def find_anchor_spans(tokens: List[str], alias_token_sets: List[Tuple[str, ...]]) -> List[Tuple[int, int]]:
    spans = []
    for alias_toks in alias_token_sets:
        n = len(alias_toks)
        if n == 0 or n > len(tokens):
            continue
        for i in range(len(tokens) - n + 1):
            if tuple(tokens[i:i+n]) == alias_toks:
                spans.append((i, i+n))
    # deduplicate
    return sorted(set(spans))


def collect_candidates_from_doc(tokens: List[str], anchor_spans: List[Tuple[int, int]], window: int = 6) -> List[str]:
    if not tokens:
        return []
    cand = []
    if not anchor_spans:
        # Fallback: use whole doc, but cap length
        anchor_spans = [(0, min(len(tokens), 12))]
    for i, j in anchor_spans:
        left = max(0, i - window)
        right = min(len(tokens), j + window)
        window_tokens = tokens[left:right]
        cand.extend(iter_ngrams(window_tokens, max_n=3))
    return cand


def load_aliases(alias_csv: Path) -> pd.DataFrame:
    df = pd.read_csv(alias_csv)
    df["alias_norm"] = df["alias"].map(normalize_text)
    return df


def fetch_explicit_docs(db_path: Path, kind: str) -> pd.DataFrame:
    conn = sqlite3.connect(db_path)
    if kind == "telegram":
        q = """
        SELECT r.issuer, r.entity_level, r.offering_id, r.bond_name,
               m.channel || ':' || m.id AS source_key,
               'telegram' AS source,
               m.text AS text
        FROM telegram_matches_resolved r
        JOIN messages m
          ON r.channel = m.channel AND r.message_id = m.id
        WHERE r.match_origin = 'explicit'
          AND r.issuer IS NOT NULL
          AND m.text IS NOT NULL
          AND LENGTH(TRIM(m.text)) > 0
        """
    elif kind == "smartlab_posts":
        q = """
        SELECT r.issuer, r.entity_level, r.offering_id, r.bond_name,
               'post:' || p.post_id AS source_key,
               'smartlab_post' AS source,
               COALESCE(p.title, '') || ' ' || COALESCE(p.tags, '') || ' ' || COALESCE(p.text, '') AS text
        FROM smartlab_post_matches_resolved r
        JOIN posts p ON r.post_id = p.post_id
        WHERE r.match_origin = 'explicit'
          AND r.issuer IS NOT NULL
        """
    elif kind == "smartlab_comments":
        q = """
        SELECT r.issuer, r.entity_level, r.offering_id, r.bond_name,
               'comment:' || c.comment_id AS source_key,
               'smartlab_comment' AS source,
               c.text AS text
        FROM smartlab_comment_matches_resolved r
        JOIN comments c ON r.comment_id = c.comment_id
        WHERE r.match_origin = 'explicit'
          AND r.issuer IS NOT NULL
          AND c.text IS NOT NULL
          AND LENGTH(TRIM(c.text)) > 0
        """
    else:
        raise ValueError(kind)
    out = pd.read_sql_query(q, conn)
    conn.close()
    out["text_norm"] = out["text"].map(normalize_text)
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description="Mine colloquial issuer aliases from explicit matches")
    ap.add_argument("--telegram-db", default="data.db")
    ap.add_argument("--smartlab-db", default="smartlab_bonds.db")
    ap.add_argument("--issuer-aliases", default="issuer_aliases_candidates.csv")
    ap.add_argument("--out-csv", default="colloquial_alias_candidates.csv")
    ap.add_argument("--min-docs-in-issuer", type=int, default=4)
    ap.add_argument("--max-docs-outside", type=int, default=3)
    ap.add_argument("--top-k-per-issuer", type=int, default=25)
    args = ap.parse_args()

    alias_df = load_aliases(Path(args.issuer_aliases))
    issuer_alias_tokens: Dict[str, List[Tuple[str, ...]]] = defaultdict(list)
    issuer_alias_norms: Dict[str, set] = defaultdict(set)
    for _, r in alias_df.iterrows():
        issuer = r["issuer"]
        an = r["alias_norm"]
        toks = alias_tokens(an)
        if toks:
            issuer_alias_tokens[issuer].append(toks)
            issuer_alias_norms[issuer].add(an)

    docs = pd.concat(
        [
            fetch_explicit_docs(Path(args.telegram_db), "telegram"),
            fetch_explicit_docs(Path(args.smartlab_db), "smartlab_posts"),
            fetch_explicit_docs(Path(args.smartlab_db), "smartlab_comments"),
        ],
        ignore_index=True,
    )
    docs = docs.drop_duplicates(subset=["source_key", "issuer"]).copy()
    docs["tokens"] = docs["text_norm"].map(tokenize)

    cand_docs_by_issuer: Dict[str, Counter] = defaultdict(Counter)
    cand_sources_by_issuer: Dict[str, Counter] = defaultdict(Counter)
    cand_examples: Dict[Tuple[str, str], List[str]] = defaultdict(list)
    candidate_to_issuers: Dict[str, set] = defaultdict(set)

    for _, row in docs.iterrows():
        issuer = row["issuer"]
        tokens = row["tokens"]
        if not tokens:
            continue
        spans = find_anchor_spans(tokens, issuer_alias_tokens.get(issuer, []))
        cands = set()
        for gram in collect_candidates_from_doc(tokens, spans, window=6):
            gram = normalize_text(gram)
            if not looks_like_candidate(gram):
                continue
            if gram in issuer_alias_norms[issuer]:
                continue
            # exclude ngrams fully covered by generic vocabulary
            toks = gram.split()
            if any(tok in GENERIC_STOP for tok in toks):
                continue
            cands.add(gram)
        for cand in cands:
            cand_docs_by_issuer[issuer][cand] += 1
            cand_sources_by_issuer[issuer][cand] += 1
            candidate_to_issuers[cand].add(issuer)
            key = (issuer, cand)
            if len(cand_examples[key]) < 3:
                snippet = row["text_norm"][:220]
                cand_examples[key].append(snippet)

    rows = []
    for issuer, counter in cand_docs_by_issuer.items():
        for cand, docs_in in counter.items():
            docs_out = sum(
                cand_docs_by_issuer[other].get(cand, 0)
                for other in candidate_to_issuers[cand]
                if other != issuer
            )
            if docs_in < args.min_docs_in_issuer:
                continue
            if docs_out > args.max_docs_outside:
                continue
            score = docs_in / (1.0 + docs_out)
            score *= math.log1p(docs_in)
            rows.append(
                {
                    "issuer": issuer,
                    "candidate_alias": cand,
                    "docs_in_issuer": docs_in,
                    "docs_outside_issuer": docs_out,
                    "distinct_issuers_with_candidate": len(candidate_to_issuers[cand]),
                    "score": round(score, 4),
                    "examples": " || ".join(cand_examples[(issuer, cand)]),
                }
            )

    out = pd.DataFrame(rows)
    if out.empty:
        print("No candidates survived filters.")
        out.to_csv(args.out_csv, index=False, encoding="utf-8-sig")
        return

    out = out.sort_values(["issuer", "score", "docs_in_issuer"], ascending=[True, False, False]).copy()
    out["rank_within_issuer"] = out.groupby("issuer").cumcount() + 1
    out = out[out["rank_within_issuer"] <= args.top_k_per_issuer].copy()
    out.to_csv(args.out_csv, index=False, encoding="utf-8-sig")
    print(f"Saved {len(out):,} colloquial alias candidates to {args.out_csv}")


if __name__ == "__main__":
    main()
