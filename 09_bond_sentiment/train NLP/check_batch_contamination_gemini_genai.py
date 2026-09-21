#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
import os
import random
import re
import time
from pathlib import Path
from typing import Dict, List, Any

import pandas as pd
from google import genai
from google.genai import types


SYSTEM_PROMPT = """Ты аналитик российского долгового рынка.

Тебе дан набор независимых сообщений из Telegram-каналов, форумов и комментариев,
связанных с первичными размещениями корпоративных облигаций на российском рынке.

ВАЖНО:
- Каждый элемент классифицируй СТРОГО независимо от других элементов батча.
- Не сравнивай сообщения между собой.
- Не переноси выводы из одного элемента в другой.
- Для каждого id используй только его собственный текст и его собственный контекст.

Для каждого элемента определи:
1) relevance_to_placement:
   - yes: текст относится к кредитному качеству эмитента, спросу на выпуск,
          успеху/неуспеху размещения, купону, переподписке, рискам дефолта,
          целесообразности участия в размещении
   - no: текст не даёт релевантного сигнала для первичного размещения

2) sentiment:
   - positive: сильный спрос, хорошее качество эмитента, переподписка, снижение купона,
               позитивный прогноз, рекомендация участвовать / покупать
   - negative: слабый спрос, проблемы эмитента, высокая долговая нагрузка, опасения дефолта,
               рекомендация не участвовать / не покупать, риск провала размещения
   - neutral: сухой факт без явной оценки или недостаточно сигнала

3) confidence: число от 0 до 1
4) reason_short: очень короткое объяснение (до 12 слов)

Отвечай СТРОГО JSON-массивом без markdown и без дополнительного текста.
"""

RESULTS_JSON_SCHEMA: Dict[str, Any] = {
    "type": "array",
    "items": {
        "type": "object",
        "properties": {
            "id": {"type": "string"},
            "relevance_to_placement": {
                "type": "string",
                "enum": ["yes", "no"],
            },
            "sentiment": {
                "type": "string",
                "enum": ["positive", "negative", "neutral"],
            },
            "confidence": {"type": "number", "minimum": 0, "maximum": 1},
            "reason_short": {"type": "string"},
        },
        "required": [
            "id",
            "relevance_to_placement",
            "sentiment",
            "confidence",
            "reason_short",
        ],
        "additionalProperties": False,
    },
}

ALLOWED_SENTIMENTS = {"positive", "negative", "neutral"}
ALLOWED_RELEVANCE = {"yes", "no"}


def clean_text(s: str) -> str:
    s = str(s or "")
    s = s.replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def safe_json_loads(raw: str):
    raw = (raw or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-zA-Z0-9_-]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw)
    return json.loads(raw.strip())


def extract_retry_delay_seconds(exc: Exception) -> float | None:
    text = str(exc)
    patterns = [
        r"retry in\s+([0-9]+(?:\.[0-9]+)?)s",
        r"retry_delay\s*\{\s*seconds:\s*([0-9]+)",
        r"Retry in\s+([0-9]+(?:\.[0-9]+)?)s",
    ]
    for pattern in patterns:
        m = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
        if m:
            try:
                return float(m.group(1))
            except ValueError:
                return None
    return None


def build_item(row: pd.Series) -> str:
    issuer = clean_text(row.get("issuer", ""))
    bond_name = clean_text(row.get("bond_name", ""))
    isin = clean_text(row.get("ISIN", ""))
    source = clean_text(row.get("source", ""))
    channel = clean_text(row.get("channel", ""))
    text = clean_text(row.get("text", ""))
    parent_text = clean_text(row.get("parent_text", ""))
    uid = str(row["mention_uid"])

    parts = [
        "<ITEM>",
        f"id: {uid}",
        f"issuer: {issuer}",
        f"bond_name: {bond_name}",
        f"ISIN: {isin}",
        f"source: {source}",
        f"channel: {channel}",
        f"text: {text}",
    ]
    if parent_text:
        parts.append(f"parent_text: {parent_text}")
    parts.append("</ITEM>")
    return "\n".join(parts)


def parse_results_payload(payload: Any) -> Dict[str, dict]:
    if isinstance(payload, str):
        arr = safe_json_loads(payload)
    else:
        arr = payload
    out: Dict[str, dict] = {}
    if not isinstance(arr, list):
        raise ValueError("Model output is not a JSON array")
    for item in arr:
        uid = str(item.get("id", "")).strip()
        if not uid:
            continue
        sentiment = str(item.get("sentiment", "neutral")).strip().lower()
        relevance = str(item.get("relevance_to_placement", "no")).strip().lower()
        reason = clean_text(item.get("reason_short", ""))[:200]
        try:
            confidence = float(item.get("confidence", 0.0))
        except Exception:
            confidence = 0.0
        confidence = min(max(confidence, 0.0), 1.0)
        if sentiment not in ALLOWED_SENTIMENTS:
            sentiment = "neutral"
        if relevance not in ALLOWED_RELEVANCE:
            relevance = "no"
        if relevance == "no":
            sentiment = "neutral"
        out[uid] = {
            "sentiment": sentiment,
            "relevance_to_placement": relevance,
            "confidence": confidence,
            "reason_short": reason,
        }
    return out


def generate_with_retry(client: genai.Client, model_name: str, prompt_text: str, max_retries: int = 4, sleep_base: float = 5.0) -> Any:
    last_err = None
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=prompt_text,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    temperature=0,
                    max_output_tokens=4096,
                    response_mime_type="application/json",
                    response_json_schema=RESULTS_JSON_SCHEMA,
                ),
            )
            parsed = getattr(response, "parsed", None)
            if parsed is not None:
                return parsed
            return response.text
        except Exception as e:
            last_err = e
            retry_after = extract_retry_delay_seconds(e)
            wait = retry_after if retry_after is not None else sleep_base * (2 ** attempt)
            wait = min(max(wait, sleep_base), 90.0)
            print(f"  retry {attempt + 1}/{max_retries} after error: {e}")
            if attempt < max_retries - 1:
                time.sleep(wait)
    raise RuntimeError(f"API failed after retries: {last_err}")


def classify_rows(client: genai.Client, model_name: str, rows: List[pd.Series]) -> Dict[str, dict]:
    user_prompt = "\n\n".join(build_item(r) for r in rows)
    payload = generate_with_retry(client, model_name, user_prompt)
    return parse_results_payload(payload)


def batched(iterable: List[pd.Series], batch_size: int) -> List[List[pd.Series]]:
    return [iterable[i:i + batch_size] for i in range(0, len(iterable), batch_size)]


def compare_cols(df: pd.DataFrame, left: str, right: str) -> dict:
    valid = df[left].notna() & df[right].notna()
    sub = df.loc[valid].copy()
    n = len(sub)
    if n == 0:
        return {"n": 0, "agreement": None}
    agreement = (sub[left] == sub[right]).mean()
    return {"n": int(n), "agreement": float(agreement)}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output-dir", default="gemini_contamination_check", type=Path)
    parser.add_argument("--api-key", default=os.getenv("GEMINI_API_KEY"))
    parser.add_argument("--model", default="gemini-3.1-flash-lite-preview")
    parser.add_argument("--sample-size", type=int, default=100)
    parser.add_argument("--batch-size", type=int, default=20)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--min-text-len", type=int, default=10)
    parser.add_argument("--delay-single", type=float, default=5.0)
    parser.add_argument("--delay-batch", type=float, default=5.0)
    args = parser.parse_args()

    if not args.api_key:
        raise ValueError("Передай API key через --api-key или GEMINI_API_KEY")

    args.output_dir.mkdir(parents=True, exist_ok=True)

    client = genai.Client(api_key=args.api_key)

    df = pd.read_csv(args.input, low_memory=False)
    df["text"] = df["text"].fillna("").astype(str)
    df = df[df["text"].str.len() >= args.min_text_len].copy()
    df = df.drop_duplicates(subset=["mention_uid"]).copy()

    if len(df) < args.sample_size:
        raise ValueError(f"В файле только {len(df)} уникальных сообщений, sample_size={args.sample_size} слишком большой")

    sampled = df.sample(n=args.sample_size, random_state=args.seed).copy().reset_index(drop=True)
    sampled.to_csv(args.output_dir / "sample_used.csv", index=False, encoding="utf-8-sig")

    rows = [row for _, row in sampled.iterrows()]

    single_results = {}
    print(f"[1/3] Single-item classification: {len(rows)} requests")
    for i, row in enumerate(rows, 1):
        uid = str(row["mention_uid"])
        try:
            res = classify_rows(client, args.model, [row])
            single_results[uid] = res.get(uid, {})
        except Exception as e:
            single_results[uid] = {
                "sentiment": None,
                "relevance_to_placement": None,
                "confidence": None,
                "reason_short": None,
                "error": str(e),
            }
        if args.delay_single > 0 and i < len(rows):
            time.sleep(args.delay_single)

    batch_results = {}
    print(f"[2/3] Batched classification: batch_size={args.batch_size}")
    original_batches = batched(rows, args.batch_size)
    for j, chunk in enumerate(original_batches, 1):
        try:
            res = classify_rows(client, args.model, chunk)
            batch_results.update(res)
        except Exception as e:
            for row in chunk:
                uid = str(row["mention_uid"])
                batch_results[uid] = {
                    "sentiment": None,
                    "relevance_to_placement": None,
                    "confidence": None,
                    "reason_short": None,
                    "error": str(e),
                }
        if args.delay_batch > 0 and j < len(original_batches):
            time.sleep(args.delay_batch)

    shuffled_rows = rows.copy()
    random.Random(args.seed + 1).shuffle(shuffled_rows)
    shuffled_results = {}
    print(f"[3/3] Batched + shuffled order classification")
    shuffled_batches = batched(shuffled_rows, args.batch_size)
    for j, chunk in enumerate(shuffled_batches, 1):
        try:
            res = classify_rows(client, args.model, chunk)
            shuffled_results.update(res)
        except Exception as e:
            for row in chunk:
                uid = str(row["mention_uid"])
                shuffled_results[uid] = {
                    "sentiment": None,
                    "relevance_to_placement": None,
                    "confidence": None,
                    "reason_short": None,
                    "error": str(e),
                }
        if args.delay_batch > 0 and j < len(shuffled_batches):
            time.sleep(args.delay_batch)

    out = sampled.copy()

    def map_field(results_dict, field):
        return out["mention_uid"].astype(str).map(lambda x: results_dict.get(str(x), {}).get(field))

    for prefix, resdict in [
        ("single", single_results),
        ("batch", batch_results),
        ("shuffle", shuffled_results),
    ]:
        out[f"{prefix}_sentiment"] = map_field(resdict, "sentiment")
        out[f"{prefix}_relevance"] = map_field(resdict, "relevance_to_placement")
        out[f"{prefix}_confidence"] = map_field(resdict, "confidence")
        out[f"{prefix}_reason"] = map_field(resdict, "reason_short")
        out[f"{prefix}_error"] = map_field(resdict, "error")

    summary = {
        "sample_size": int(len(out)),
        "batch_size": int(args.batch_size),
        "model": args.model,
        "single_vs_batch": compare_cols(out, "single_sentiment", "batch_sentiment"),
        "single_vs_shuffle": compare_cols(out, "single_sentiment", "shuffle_sentiment"),
        "batch_vs_shuffle": compare_cols(out, "batch_sentiment", "shuffle_sentiment"),
        "single_relevance_vs_batch": compare_cols(out, "single_relevance", "batch_relevance"),
        "single_relevance_vs_shuffle": compare_cols(out, "single_relevance", "shuffle_relevance"),
        "batch_relevance_vs_shuffle": compare_cols(out, "batch_relevance", "shuffle_relevance"),
    }

    sentiment_cm_single_batch = pd.crosstab(out["single_sentiment"], out["batch_sentiment"], dropna=False)
    sentiment_cm_batch_shuffle = pd.crosstab(out["batch_sentiment"], out["shuffle_sentiment"], dropna=False)

    diffs = out[
        (out["single_sentiment"] != out["batch_sentiment"]) |
        (out["single_sentiment"] != out["shuffle_sentiment"]) |
        (out["batch_sentiment"] != out["shuffle_sentiment"])
    ].copy()

    out.to_csv(args.output_dir / "contamination_check_results.csv", index=False, encoding="utf-8-sig")
    diffs.to_csv(args.output_dir / "contamination_check_differences.csv", index=False, encoding="utf-8-sig")
    sentiment_cm_single_batch.to_csv(args.output_dir / "cm_single_vs_batch.csv", encoding="utf-8-sig")
    sentiment_cm_batch_shuffle.to_csv(args.output_dir / "cm_batch_vs_shuffle.csv", encoding="utf-8-sig")

    with open(args.output_dir / "summary.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print("\n=== SUMMARY ===")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print("\nTopline:")
    svb = summary["single_vs_batch"]["agreement"]
    bvs = summary["batch_vs_shuffle"]["agreement"]
    if svb is not None:
        print(f"sentiment agreement single vs batch:   {svb:.3%}")
    if bvs is not None:
        print(f"sentiment agreement batch vs shuffled: {bvs:.3%}")
    print(f"differences saved to: {args.output_dir / 'contamination_check_differences.csv'}")


if __name__ == "__main__":
    main()
