#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gemini sentiment classification for sampled texts.

Classifies texts from samples_for_gemini.csv (output of sample_for_labeling.py).
Results are merged with existing training data for fine-tuning.

Usage:
  export GEMINI_API_KEY="your-key"
  python classify_samples_gemini.py --input samples_for_gemini.csv

Resume after interruption (reads already-classified rows from output CSV):
  python classify_samples_gemini.py --input samples_for_gemini.csv
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

# ============================================================================
#  Configuration
# ============================================================================

DEFAULT_MODEL = "gemini-2.5-flash"
DEFAULT_BATCH_ITEMS = 20
DEFAULT_MAX_BATCH_CHARS = 40000
DEFAULT_MAX_REQUESTS = 20
DEFAULT_DELAY = 15.0
DEFAULT_MIN_TEXT_LEN = 15
DEFAULT_MAX_RETRIES = 5
DEFAULT_MAX_OUTPUT_TOKENS = 8192

ALLOWED_SENTIMENTS = {"positive", "negative", "neutral"}
ALLOWED_RELEVANCE = {"yes", "no"}

SYSTEM_PROMPT = """Ты аналитик российского первичного долгового рынка.

Ты получаешь пакет сообщений из Telegram, Smart-Lab и других источников.
Для КАЖДОГО сообщения определи:
1) относится ли оно к кредитному качеству указанного эмитента (issuer)
   или перспективам его размещения облигаций;
2) если относится — какова тональность.

ВАЖНО:
- issuer — это TARGET ISSUER. Оценивай текст ТОЛЬКО относительно этого эмитента.
- Если issuer упомянут как брокер, банк, организатор, агент, площадка —
  это НЕ релевантно (relevance = "no").
- Если в тексте несколько эмитентов — оценивай ТОЛЬКО фрагмент про target issuer.
- Если текст про депозиты, тарифы, ИИС, карты, комиссии банка — это "no".
- Сухие параметры выпуска target issuer — это "yes" + "neutral".

relevance_to_placement:
- "yes" — текст связан с кредитным качеством, спросом на размещение, рисками,
  долговой нагрузкой, рекомендацией участвовать/не участвовать.
- "no" — не даёт информации про качество эмитента или перспективы размещения.

sentiment (только если relevance = "yes"):
- "positive" — сильный спрос, переподписка, доверие, качественная книга.
- "negative" — слабый спрос, сомнения, высокая нагрузка, риск дефолта.
- "neutral" — сухой факт, анонс, параметры без оценки.

Если relevance = "no", sentiment = "neutral".

confidence — число 0..1.
reason_short — максимум 15 слов.

Возвращай ТОЛЬКО JSON-массив, без markdown и комментариев.
"""

RESULTS_JSON_SCHEMA: Dict[str, Any] = {
    "type": "array",
    "items": {
        "type": "object",
        "properties": {
            "id": {"type": "integer", "minimum": 1},
            "relevance_to_placement": {"type": "string", "enum": ["yes", "no"]},
            "sentiment": {"type": "string", "enum": ["positive", "negative", "neutral"]},
            "confidence": {"type": "number", "minimum": 0, "maximum": 1},
            "reason_short": {"type": "string"},
        },
        "required": ["id", "relevance_to_placement", "sentiment", "confidence", "reason_short"],
        "additionalProperties": False,
    },
}

LAST_REQUEST_TS = 0.0


# ============================================================================
#  Helpers
# ============================================================================

def norm(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def safe_float(value: Any, default: float = 0.5) -> float:
    try:
        x = float(value)
        return max(0.0, min(1.0, x)) if not math.isnan(x) else default
    except (TypeError, ValueError):
        return default


def sleep_min_interval(min_interval: float):
    global LAST_REQUEST_TS
    now = time.time()
    if LAST_REQUEST_TS and (now - LAST_REQUEST_TS) < min_interval:
        time.sleep(min_interval - (now - LAST_REQUEST_TS))


def mark_sent():
    global LAST_REQUEST_TS
    LAST_REQUEST_TS = time.time()


# ============================================================================
#  Input preparation
# ============================================================================

def generate_uid(row: pd.Series, idx: int) -> str:
    """Generate unique ID from available columns."""
    if "mention_uid" in row.index and pd.notna(row.get("mention_uid")):
        return str(row["mention_uid"])
    # Build from source identifiers
    parts = []
    for col in ["channel", "message_id", "post_id", "comment_id"]:
        val = row.get(col)
        if pd.notna(val):
            parts.append(f"{col}:{val}")
    if parts:
        issuer = norm(row.get("issuer", ""))
        return "|".join(parts) + f"|{issuer}" if issuer else "|".join(parts)
    return f"row:{idx}"


def load_already_done(output_path: Path) -> set:
    """Load UIDs already classified from output CSV."""
    if not output_path.exists():
        return set()
    try:
        df = pd.read_csv(output_path, usecols=["uid", "sentiment_status"], low_memory=False)
        done = set(df.loc[df["sentiment_status"].isin(["classified", "skipped_short"]), "uid"].astype(str))
        return done
    except Exception:
        return set()


def build_prompt_item(row: pd.Series) -> Dict[str, str]:
    """Build payload dict for one text item."""
    issuer = norm(row.get("issuer", ""))
    text = norm(row.get("text", ""))
    category = norm(row.get("sample_category", ""))

    # For Smart-Lab posts, prepend title
    title = norm(row.get("title", ""))
    if title and title not in text:
        text = f"{title}\n{text}"

    # For Smart-Lab comments, add post title as context
    post_title = norm(row.get("post_title", ""))

    return {
        "issuer": issuer,
        "text": text,
        "post_title": post_title,
        "category": category,
    }


def build_user_prompt(batch: List[Dict]) -> str:
    """Build the user prompt from a batch of items."""
    parts = []
    for idx, item in enumerate(batch, start=1):
        lines = [
            "<ITEM>",
            f"id: {idx}",
            f"issuer: {item['issuer'] or 'UNKNOWN'}",
            "text:",
            item["text"] or "NA",
        ]
        if item.get("post_title"):
            lines.extend(["post_title (context):", item["post_title"]])
        lines.append("</ITEM>")
        parts.append("\n".join(lines))
    return "\n\n".join(parts)


# ============================================================================
#  Gemini API
# ============================================================================

def call_gemini(
    client, model_name: str, prompt: str,
    max_retries: int, max_output_tokens: int, min_interval: float,
) -> Dict:
    """Call Gemini with retries. Returns {raw_text, parsed, usage}."""
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            sleep_min_interval(min_interval)
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=types.GenerateContentConfig(
                    system_instruction=SYSTEM_PROMPT,
                    temperature=0,
                    max_output_tokens=max_output_tokens,
                    response_mime_type="application/json",
                    response_json_schema=RESULTS_JSON_SCHEMA,
                ),
            )
            mark_sent()
            return {
                "raw_text": getattr(response, "text", "") or "",
                "parsed": getattr(response, "parsed", None),
                "usage": getattr(response, "usage_metadata", None),
                "version": getattr(response, "model_version", None),
            }
        except Exception as e:
            last_exc = e
            # Extract retry delay from error message
            m = re.search(r"retry.*?(\d+(?:\.\d+)?)\s*s", str(e), re.IGNORECASE)
            wait = float(m.group(1)) if m else min(5 * 2 ** (attempt - 1), 60)
            if attempt < max_retries:
                print(f"  retry {attempt}/{max_retries} after {wait:.0f}s: {e}", flush=True)
                time.sleep(wait)
    raise RuntimeError(str(last_exc))


def parse_response(raw) -> List[Dict]:
    """Parse Gemini response into list of item dicts."""
    payload = raw["parsed"] if raw["parsed"] is not None else raw["raw_text"]
    if isinstance(payload, list):
        return payload
    if isinstance(payload, str):
        text = payload.strip()
        if text.startswith("```"):
            text = re.sub(r"^```(?:json)?\s*", "", text)
            text = re.sub(r"\s*```$", "", text)
        start, end = text.find("["), text.rfind("]")
        if start != -1 and end > start:
            return json.loads(text[start:end + 1])
    raise ValueError("Cannot parse response")


def normalize_item(item: Dict, batch_len: int) -> Optional[Dict]:
    """Validate and normalize one parsed item."""
    if not isinstance(item, dict):
        return None
    try:
        idx = int(item.get("id"))
    except (TypeError, ValueError):
        return None
    if not (1 <= idx <= batch_len):
        return None

    rel = str(item.get("relevance_to_placement", "")).strip().lower()
    if rel not in ALLOWED_RELEVANCE:
        rel = "no"

    sent = str(item.get("sentiment", "")).strip().lower()
    if sent not in ALLOWED_SENTIMENTS:
        sent = "neutral"
    if rel == "no":
        sent = "neutral"

    conf = safe_float(item.get("confidence"))
    reason = norm(item.get("reason_short", ""))[:200] or None

    return {"id": idx, "relevance": rel, "sentiment": sent, "confidence": conf, "reason": reason}


# ============================================================================
#  Main loop
# ============================================================================

def main():
    p = argparse.ArgumentParser(description="Classify sampled texts with Gemini")
    p.add_argument("--input", type=Path, default=Path("samples_for_gemini.csv"))
    p.add_argument("--output", type=Path, default=Path("samples_classified_gemini.csv"))
    p.add_argument("--model", default=DEFAULT_MODEL)
    p.add_argument("--max-batch-items", type=int, default=DEFAULT_BATCH_ITEMS)
    p.add_argument("--max-batch-chars", type=int, default=DEFAULT_MAX_BATCH_CHARS)
    p.add_argument("--max-requests", type=int, default=DEFAULT_MAX_REQUESTS)
    p.add_argument("--delay", type=float, default=DEFAULT_DELAY)
    p.add_argument("--min-text-len", type=int, default=DEFAULT_MIN_TEXT_LEN)
    p.add_argument("--max-retries", type=int, default=DEFAULT_MAX_RETRIES)
    p.add_argument("--api-key", default=os.getenv("GEMINI_API_KEY", ""))
    p.add_argument("--no-confirm", action="store_true")
    args = p.parse_args()

    if not args.api_key:
        raise ValueError("Set GEMINI_API_KEY or pass --api-key")
    if genai is None:
        raise ImportError("pip install google-genai")

    client = genai.Client(api_key=args.api_key)

    # Load input
    df = pd.read_csv(args.input, low_memory=False)
    df["uid"] = [generate_uid(row, i) for i, row in df.iterrows()]
    print(f"Input: {len(df)} rows from {args.input}")

    # Load progress
    done = load_already_done(args.output)
    print(f"Already classified: {len(done)}")

    # Prepare pending items
    pending = []
    results = []  # accumulate results for saving

    # Load existing results if resuming
    if args.output.exists():
        existing = pd.read_csv(args.output, low_memory=False)
        results = existing.to_dict("records")

    for _, row in df.iterrows():
        uid = str(row["uid"])
        if uid in done:
            continue
        text = norm(row.get("text", ""))
        if len(text) < args.min_text_len:
            results.append({
                "uid": uid, "issuer": norm(row.get("issuer")),
                "sample_category": norm(row.get("sample_category")),
                "text": text,
                "sentiment_status": "skipped_short",
                "relevance_to_placement": None, "sentiment": None,
                "sentiment_confidence": None, "reason_short": None,
            })
            done.add(uid)
            continue
        pending.append((uid, row, build_prompt_item(row)))

    print(f"Pending: {len(pending)}")

    # Build batches
    batches = []
    current_batch = []
    current_chars = 0
    for uid, row, payload in pending:
        chars = sum(len(str(v)) for v in payload.values()) + 200
        if current_batch and (len(current_batch) >= args.max_batch_items or current_chars + chars > args.max_batch_chars):
            batches.append(current_batch)
            current_batch = []
            current_chars = 0
        current_batch.append((uid, row, payload))
        current_chars += chars
    if current_batch:
        batches.append(current_batch)

    if len(batches) > args.max_requests:
        batches = batches[:args.max_requests]
        print(f"Limited to {args.max_requests} requests")

    print(f"Batches: {len(batches)}")
    if batches and not args.no_confirm:
        answer = input(f"Send {len(batches)} requests to Gemini? (y/n): ").strip().lower()
        if answer != "y":
            print("Cancelled.")
            return

    # Process batches
    total_classified = 0
    try:
        for batch_idx, batch in enumerate(batches, 1):
            print(f"\nBatch {batch_idx}/{len(batches)} | items={len(batch)}", flush=True)

            prompt = build_user_prompt([p for _, _, p in batch])
            try:
                raw = call_gemini(
                    client, args.model, prompt,
                    args.max_retries, DEFAULT_MAX_OUTPUT_TOKENS, args.delay,
                )
                items = parse_response(raw)
                normalized = {}
                for item in items:
                    n = normalize_item(item, len(batch))
                    if n:
                        normalized[n["id"]] = n

                for idx, (uid, row, payload) in enumerate(batch, 1):
                    if idx in normalized:
                        n = normalized[idx]
                        results.append({
                            "uid": uid,
                            "issuer": payload["issuer"],
                            "sample_category": payload["category"],
                            "text": payload["text"][:500],
                            "sentiment_status": "classified",
                            "relevance_to_placement": n["relevance"],
                            "sentiment": n["sentiment"],
                            "sentiment_confidence": n["confidence"],
                            "reason_short": n["reason"],
                        })
                        total_classified += 1
                    else:
                        results.append({
                            "uid": uid,
                            "issuer": payload["issuer"],
                            "sample_category": payload["category"],
                            "text": payload["text"][:500],
                            "sentiment_status": "missing_in_response",
                            "relevance_to_placement": None, "sentiment": None,
                            "sentiment_confidence": None, "reason_short": None,
                        })
                    done.add(uid)

                ok = sum(1 for i in range(1, len(batch)+1) if i in normalized)
                print(f"  classified: {ok}/{len(batch)}", flush=True)

            except Exception as e:
                print(f"  ERROR: {e}", flush=True)
                for uid, row, payload in batch:
                    results.append({
                        "uid": uid,
                        "issuer": payload["issuer"],
                        "sample_category": payload["category"],
                        "text": payload["text"][:500],
                        "sentiment_status": "api_error",
                        "relevance_to_placement": None, "sentiment": None,
                        "sentiment_confidence": None,
                        "reason_short": str(e)[:200],
                    })
                    done.add(uid)

            # Save after every batch
            pd.DataFrame(results).to_csv(args.output, index=False, encoding="utf-8-sig")

    except KeyboardInterrupt:
        print("\nInterrupted. Saving progress...")

    # Final save
    out_df = pd.DataFrame(results)
    out_df.to_csv(args.output, index=False, encoding="utf-8-sig")

    # Summary
    print(f"\n{'='*50}")
    print(f"Total results: {len(out_df)}")
    if "sentiment_status" in out_df.columns:
        print(out_df["sentiment_status"].value_counts().to_string())
    if "sample_category" in out_df.columns and "sentiment" in out_df.columns:
        classified = out_df[out_df["sentiment_status"] == "classified"]
        if not classified.empty:
            print(f"\nBy category:")
            ct = pd.crosstab(classified["sample_category"], classified["sentiment"])
            print(ct.to_string())
            print(f"\nRelevance:")
            print(classified["relevance_to_placement"].value_counts().to_string())

    print(f"\nSaved: {args.output}")
    print(f"Total classified this run: {total_classified}")


if __name__ == "__main__":
    main()
