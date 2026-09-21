#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import json
import math
import os
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

# ------------------------------------------------------------
# Defaults for final long run
# ------------------------------------------------------------
DEFAULT_MODEL = "gemini-3.1-flash-lite-preview"
DEFAULT_BATCH_ITEMS = 12
DEFAULT_MAX_BATCH_CHARS = 30000
DEFAULT_MAX_REQUESTS = 490          # keep headroom under 500 RPD
DEFAULT_DELAY = 5.0                 # ~12 RPM, safer than 15 RPM hard ceiling
DEFAULT_MIN_TEXT_LEN = 10
DEFAULT_SAVE_EVERY = 1              # save cache after every batch
DEFAULT_MAX_RETRIES = 5
DEFAULT_MAX_OUTPUT_TOKENS = 8192
DEFAULT_OUTPUT_WRITE_EVERY = 1      # write partial CSV after every batch

ALLOWED_SENTIMENTS = {"positive", "negative", "neutral"}
ALLOWED_RELEVANCE = {"yes", "no"}

SYSTEM_PROMPT = """Ты аналитик российского первичного долгового рынка.

Ты получаешь пакет сообщений из Telegram, Smart-Lab и других источников про корпоративные облигации РФ.
Для КАЖДОГО сообщения нужно отдельно определить:
1) относится ли оно к кредитному качеству target issuer или перспективам его первичного размещения облигаций;
2) если относится, какова тональность именно относительно качества target issuer и вероятного успеха размещения.

КРИТИЧЕСКИ ВАЖНО:
- Рассматривай каждый элемент батча строго независимо от других.
- Не сравнивай сообщения между собой.
- Не переноси тональность из одного сообщения на другое.
- Для текущего элемента используй только поля текущего элемента.
- Если соседние элементы противоречат друг другу, игнорируй это и классифицируй каждый элемент отдельно.
- issuer — это TARGET ISSUER: эмитент, для которого строится sentiment перед новым размещением.
- bond_name и isin — это TARGET EVENT: новый выпуск target issuer, перед закрытием книги которого измеряется sentiment.
- Текст не обязан упоминать именно target bond_name/isin. Он релевантен, если говорит о кредитном качестве,
  рисках, спросе на облигации или размещениях target issuer как эмитента облигаций.
- matched_* поля описывают, какой алиас/выпуск нашёл маппер в тексте. Это подсказка, но не доказательство релевантности.

Определи поля:
- relevance_to_placement = "yes", если сообщение связано с кредитным качеством target issuer,
  спросом на размещение, параметрами книги, вероятностью сильного/слабого размещения,
  рекомендацией участвовать/не участвовать, риском дефолта, долговой нагрузкой.
- relevance_to_placement = "no", если сообщение не даёт содержательной информации про качество target issuer
  или перспективы размещения.
- relevance_to_placement = "no", если целевой issuer упомянут только как брокер, банк,
  депозитарий, организатор, агент, площадка подачи заявки или место, где доступна чужая бумага.
- Если это дайджест с несколькими эмитентами, оценивай только фрагмент про target issuer.
  Если нельзя уверенно понять, какой фрагмент относится к target, ставь "no".
- Сухие параметры собственного выпуска target issuer — это "yes" + neutral.
- Комментарии про брокера/приложение/комиссии/ИИС/подачу заявки через банк — это "no",
  даже если в тексте есть слова "купон", "оферта", "размещение" про чужую бумагу.

Классы sentiment:
- positive: сильный спрос, переподписка, доверие к эмитенту, качественная книга,
  возможность снизить купон, увеличить объём, положительная оценка кредитного качества.
- negative: слабый спрос, риск плохой книги, необходимость дать премию, сомнения в кредитном качестве,
  высокая долговая нагрузка, риск дефолта, рекомендация не участвовать.
- neutral: сухой факт, анонс, параметры выпуска, новость без явной оценки,
  либо сообщение не относится к размещению/кредитному качеству.

Правила:
- Если relevance_to_placement = "no", то sentiment должен быть "neutral".
- confidence — число от 0 до 1.
- reason_short — очень короткая причина, максимум 20 слов.
- Возвращай только JSON-массив.
- Никакого markdown, никаких комментариев, никакого текста вне JSON.
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


@dataclass
class BatchItem:
    uid: str
    payload: Dict[str, Any]
    estimated_chars: int


@dataclass
class CallResult:
    parsed: Any
    raw_text: str
    response_model_version: Optional[str]
    usage_metadata: Any


LAST_REQUEST_TS = 0.0


def normalize_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def safe_float_01(value: Any, default: float = 0.5) -> float:
    try:
        x = float(value)
    except (TypeError, ValueError):
        return default
    if math.isnan(x):
        return default
    return max(0.0, min(1.0, x))


def load_json(path: Path) -> Dict[str, Any]:
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def none_if_nan(value: Any) -> Any:
    return None if pd.isna(value) else value


def load_progress_from_output_csv(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}

    prev = pd.read_csv(path, low_memory=False)
    required_cols = {"mention_uid", "sentiment_status"}
    if not required_cols.issubset(prev.columns):
        return {}

    prev["mention_uid"] = prev["mention_uid"].astype(str)
    prev = prev[prev["sentiment_status"].notna()].copy()

    progress: Dict[str, Any] = {}
    for _, row in prev.iterrows():
        uid = str(row["mention_uid"])
        progress[uid] = {
            "mention_uid": uid,
            "status": none_if_nan(row.get("sentiment_status")),
            "sentiment": none_if_nan(row.get("sentiment")),
            "relevance_to_placement": none_if_nan(row.get("relevance_to_placement")),
            "confidence": none_if_nan(row.get("sentiment_confidence")),
            "reason_short": none_if_nan(row.get("sentiment_reason_short")),
            "model_name": none_if_nan(row.get("sentiment_model_name")),
            "model_version": none_if_nan(row.get("sentiment_model_version")),
            "prompt_tokens": none_if_nan(row.get("sentiment_prompt_tokens")),
            "output_tokens": none_if_nan(row.get("sentiment_output_tokens")),
            "total_tokens": none_if_nan(row.get("sentiment_total_tokens")),
            "error_type": none_if_nan(row.get("sentiment_error_type")),
            "error_message": none_if_nan(row.get("sentiment_error_message")),
            "raw_response_excerpt": None,
        }
    return progress


def save_json(path: Path, payload: Dict[str, Any]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def extract_json_array(raw_text: str) -> str:
    raw_text = (raw_text or "").strip()
    if raw_text.startswith("```"):
        raw_text = re.sub(r"^```(?:json)?\s*", "", raw_text, flags=re.IGNORECASE)
        raw_text = re.sub(r"\s*```$", "", raw_text)
        raw_text = raw_text.strip()
    start = raw_text.find("[")
    end = raw_text.rfind("]")
    if start != -1 and end != -1 and end > start:
        return raw_text[start : end + 1]
    raise ValueError("JSON array not found in model output")


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


def sleep_to_respect_min_interval(min_interval: float) -> None:
    global LAST_REQUEST_TS
    now = time.time()
    elapsed = now - LAST_REQUEST_TS
    if LAST_REQUEST_TS and elapsed < min_interval:
        time.sleep(min_interval - elapsed)


def mark_request_sent() -> None:
    global LAST_REQUEST_TS
    LAST_REQUEST_TS = time.time()


def make_result(
    *,
    uid: str,
    status: str,
    sentiment: Optional[str] = None,
    relevance_to_placement: Optional[str] = None,
    confidence: Optional[float] = None,
    reason_short: Optional[str] = None,
    model_name: Optional[str] = None,
    model_version: Optional[str] = None,
    prompt_tokens: Optional[int] = None,
    output_tokens: Optional[int] = None,
    total_tokens: Optional[int] = None,
    error_type: Optional[str] = None,
    error_message: Optional[str] = None,
    raw_response_excerpt: Optional[str] = None,
) -> Dict[str, Any]:
    return {
        "mention_uid": str(uid),
        "status": status,
        "sentiment": sentiment,
        "relevance_to_placement": relevance_to_placement,
        "confidence": confidence,
        "reason_short": reason_short,
        "model_name": model_name,
        "model_version": model_version,
        "prompt_tokens": prompt_tokens,
        "output_tokens": output_tokens,
        "total_tokens": total_tokens,
        "error_type": error_type,
        "error_message": error_message,
        "raw_response_excerpt": raw_response_excerpt,
    }


def build_payload_from_row(row: pd.Series) -> Dict[str, Any]:
    return {
        "issuer": normalize_text(row.get("issuer", "")),
        "bond_name": normalize_text(row.get("bond_name", "")),
        "isin": normalize_text(row.get("ISIN", "")),
        "source": normalize_text(row.get("source", "")),
        "channel": normalize_text(row.get("channel", "")),
        "entity_level": normalize_text(row.get("entity_level", "")),
        "match_origin": normalize_text(row.get("match_origin", "")),
        "matched_offering_id": normalize_text(row.get("matched_offering_id", "")),
        "matched_bond_name": normalize_text(row.get("matched_bond_name", "")),
        "matched_isin": normalize_text(row.get("matched_isin", "")),
        "matched_aliases": normalize_text(row.get("matched_aliases", "")),
        "alias_types": normalize_text(row.get("alias_types", "")),
        "hours_to_event_end": normalize_text(row.get("hours_to_event_end", "")),
        "book_date": normalize_text(row.get("book_date", "")),
        "placement_date": normalize_text(row.get("placement_date", "")),
        "text": normalize_text(row.get("text", "")),
        "parent_text": normalize_text(row.get("parent_text", "")),
    }


def estimate_payload_chars(payload: Dict[str, Any]) -> int:
    total = sum(len(str(v)) for v in payload.values())
    return total + 400


def build_user_prompt(batch: List[BatchItem]) -> str:
    parts: List[str] = []
    for idx, item in enumerate(batch, start=1):
        p = item.payload
        block = [
            "<ITEM>",
            f"id: {idx}",
            f"issuer: {p.get('issuer') or 'NA'}",
            f"bond_name: {p.get('bond_name') or 'NA'}",
            f"isin: {p.get('isin') or 'NA'}",
            f"source: {p.get('source') or 'NA'}",
            f"channel: {p.get('channel') or 'NA'}",
            f"entity_level: {p.get('entity_level') or 'NA'}",
            f"match_origin: {p.get('match_origin') or 'NA'}",
            f"matched_offering_id: {p.get('matched_offering_id') or 'NA'}",
            f"matched_bond_name: {p.get('matched_bond_name') or 'NA'}",
            f"matched_isin: {p.get('matched_isin') or 'NA'}",
            f"matched_aliases: {p.get('matched_aliases') or 'NA'}",
            f"alias_types: {p.get('alias_types') or 'NA'}",
            f"hours_to_event_end: {p.get('hours_to_event_end') or 'NA'}",
            f"book_date: {p.get('book_date') or 'NA'}",
            f"placement_date: {p.get('placement_date') or 'NA'}",
            "text:",
            p.get("text") or "NA",
        ]
        if p.get("parent_text"):
            block.extend(["parent_text:", p["parent_text"]])
        block.append("</ITEM>")
        parts.append("\n".join(block))
    return "\n\n".join(parts)


def parse_model_items(raw_payload: Any) -> List[Dict[str, Any]]:
    if isinstance(raw_payload, list):
        parsed = raw_payload
    elif isinstance(raw_payload, str):
        parsed = json.loads(extract_json_array(raw_payload))
    else:
        raise ValueError("Unsupported model payload type")
    if not isinstance(parsed, list):
        raise ValueError("Model output is not a list")
    return parsed


def normalize_model_item(item: Dict[str, Any], batch_len: int) -> Optional[Dict[str, Any]]:
    if not isinstance(item, dict):
        return None
    try:
        idx = int(item.get("id"))
    except (TypeError, ValueError):
        return None
    if not (1 <= idx <= batch_len):
        return None

    relevance = str(item.get("relevance_to_placement", "")).strip().lower()
    if relevance not in ALLOWED_RELEVANCE:
        relevance = "no"

    sentiment = str(item.get("sentiment", "")).strip().lower()
    if sentiment not in ALLOWED_SENTIMENTS:
        sentiment = "neutral"
    if relevance == "no":
        sentiment = "neutral"

    confidence = safe_float_01(item.get("confidence"), default=0.5)
    reason_short = normalize_text(item.get("reason_short", ""))
    if len(reason_short.split()) > 20:
        reason_short = " ".join(reason_short.split()[:20])
    reason_short = reason_short[:220] or None

    return {
        "id": idx,
        "relevance_to_placement": relevance,
        "sentiment": sentiment,
        "confidence": confidence,
        "reason_short": reason_short,
    }


def generate_structured(
    client: genai.Client,
    model_name: str,
    prompt: str,
    max_retries: int,
    max_output_tokens: int,
    min_interval: float,
) -> CallResult:
    last_exc: Optional[Exception] = None
    for attempt in range(1, max_retries + 1):
        try:
            if types is None:
                raise ImportError("google-genai is not installed. Install it or run in an environment that provides google.genai.")
            sleep_to_respect_min_interval(min_interval)
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
            mark_request_sent()
            raw_text = getattr(response, "text", "") or ""
            parsed = getattr(response, "parsed", None)
            usage = getattr(response, "usage_metadata", None)
            version = getattr(response, "model_version", None)
            return CallResult(parsed=parsed, raw_text=raw_text, response_model_version=version, usage_metadata=usage)
        except Exception as e:
            last_exc = e
            retry_delay = extract_retry_delay_seconds(e)
            wait = retry_delay if retry_delay is not None else min(5 * (2 ** (attempt - 1)), 60)
            if attempt < max_retries:
                print(f"  retry {attempt}/{max_retries} after {wait:.1f}s: {e}")
                time.sleep(wait)
            else:
                break
    raise RuntimeError(str(last_exc) if last_exc else "Unknown API error")


def usage_counts(usage_metadata: Any) -> tuple[Optional[int], Optional[int], Optional[int]]:
    if usage_metadata is None:
        return None, None, None
    prompt_tokens = getattr(usage_metadata, "prompt_token_count", None)
    output_tokens = getattr(usage_metadata, "candidates_token_count", None)
    total_tokens = getattr(usage_metadata, "total_token_count", None)
    return prompt_tokens, output_tokens, total_tokens


def classify_batch(
    client: genai.Client,
    model_name: str,
    batch: List[BatchItem],
    max_retries: int,
    max_output_tokens: int,
    min_interval: float,
) -> Dict[str, Dict[str, Any]]:
    prompt = build_user_prompt(batch)
    last_error: Optional[str] = None
    last_raw_excerpt: Optional[str] = None

    try:
        call_result = generate_structured(
            client=client,
            model_name=model_name,
            prompt=prompt,
            max_retries=max_retries,
            max_output_tokens=max_output_tokens,
            min_interval=min_interval,
        )
        parsed_items = parse_model_items(call_result.parsed if call_result.parsed is not None else call_result.raw_text)
        last_raw_excerpt = (call_result.raw_text or "")[:1200] if call_result.raw_text else None
        prompt_tokens, output_tokens, total_tokens = usage_counts(call_result.usage_metadata)

        normalized_items: Dict[int, Dict[str, Any]] = {}
        for item in parsed_items:
            normalized = normalize_model_item(item, len(batch))
            if normalized is not None:
                normalized_items[normalized["id"]] = normalized

        results: Dict[str, Dict[str, Any]] = {}
        for idx, batch_item in enumerate(batch, start=1):
            uid = batch_item.uid
            if idx in normalized_items:
                obj = normalized_items[idx]
                results[uid] = make_result(
                    uid=uid,
                    status="classified",
                    sentiment=obj["sentiment"],
                    relevance_to_placement=obj["relevance_to_placement"],
                    confidence=obj["confidence"],
                    reason_short=obj["reason_short"],
                    model_name=model_name,
                    model_version=call_result.response_model_version,
                    prompt_tokens=prompt_tokens,
                    output_tokens=output_tokens,
                    total_tokens=total_tokens,
                )
            else:
                results[uid] = make_result(
                    uid=uid,
                    status="missing_item_in_response",
                    model_name=model_name,
                    model_version=call_result.response_model_version,
                    prompt_tokens=prompt_tokens,
                    output_tokens=output_tokens,
                    total_tokens=total_tokens,
                    error_type="missing_item_in_response",
                    raw_response_excerpt=last_raw_excerpt,
                )
        return results

    except Exception as e:
        last_error = f"{type(e).__name__}: {e}"

    return {
        item.uid: make_result(
            uid=item.uid,
            status="api_error",
            model_name=model_name,
            error_type="api_error",
            error_message=last_error,
            raw_response_excerpt=last_raw_excerpt,
        )
        for item in batch
    }


def prepare_pending_items(df: pd.DataFrame, cache: Dict[str, Any], min_text_len: int) -> List[BatchItem]:
    pending: List[BatchItem] = []
    for _, row in df.iterrows():
        uid = str(row["mention_uid"])
        if uid in cache:
            continue

        text = normalize_text(row.get("text", ""))
        if len(text) < min_text_len:
            cache[uid] = make_result(
                uid=uid,
                status="skipped_short",
                error_type="skipped_short",
                error_message=f"text shorter than {min_text_len} chars",
            )
            continue

        payload = build_payload_from_row(row)
        pending.append(BatchItem(uid=uid, payload=payload, estimated_chars=estimate_payload_chars(payload)))
    return pending


def make_batches(items: List[BatchItem], max_batch_items: int, max_batch_chars: int) -> List[List[BatchItem]]:
    batches: List[List[BatchItem]] = []
    current: List[BatchItem] = []
    current_chars = 0

    for item in items:
        item_chars = item.estimated_chars
        if current and (len(current) >= max_batch_items or current_chars + item_chars > max_batch_chars):
            batches.append(current)
            current = []
            current_chars = 0
        current.append(item)
        current_chars += item_chars

    if current:
        batches.append(current)
    return batches


def apply_results(df: pd.DataFrame, cache: Dict[str, Any]) -> pd.DataFrame:
    out = df.copy()
    out["mention_uid"] = out["mention_uid"].astype(str)

    def get_field(uid: str, field: str) -> Any:
        item = cache.get(uid)
        if isinstance(item, dict):
            return item.get(field)
        return None

    out["sentiment"] = out["mention_uid"].map(lambda uid: get_field(uid, "sentiment"))
    out["relevance_to_placement"] = out["mention_uid"].map(lambda uid: get_field(uid, "relevance_to_placement"))
    out["sentiment_confidence"] = out["mention_uid"].map(lambda uid: get_field(uid, "confidence"))
    out["sentiment_reason_short"] = out["mention_uid"].map(lambda uid: get_field(uid, "reason_short"))
    out["sentiment_status"] = out["mention_uid"].map(lambda uid: get_field(uid, "status"))
    out["sentiment_error_type"] = out["mention_uid"].map(lambda uid: get_field(uid, "error_type"))
    out["sentiment_error_message"] = out["mention_uid"].map(lambda uid: get_field(uid, "error_message"))
    out["sentiment_model_name"] = out["mention_uid"].map(lambda uid: get_field(uid, "model_name"))
    out["sentiment_model_version"] = out["mention_uid"].map(lambda uid: get_field(uid, "model_version"))
    out["sentiment_prompt_tokens"] = out["mention_uid"].map(lambda uid: get_field(uid, "prompt_tokens"))
    out["sentiment_output_tokens"] = out["mention_uid"].map(lambda uid: get_field(uid, "output_tokens"))
    out["sentiment_total_tokens"] = out["mention_uid"].map(lambda uid: get_field(uid, "total_tokens"))

    score_map = {"positive": 1, "neutral": 0, "negative": -1}
    out["sentiment_score"] = out["sentiment"].map(score_map)
    return out


def print_summary(df: pd.DataFrame) -> None:
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print("\nStatuses:")
    print(df["sentiment_status"].fillna("missing").value_counts(dropna=False))

    classified = df[df["sentiment_status"] == "classified"].copy()
    if classified.empty:
        print("\nНет успешно классифицированных записей.")
        return

    print("\nSentiment distribution (classified only):")
    print(classified["sentiment"].value_counts(dropna=False))
    print("\nRelevance distribution (classified only):")
    print(classified["relevance_to_placement"].value_counts(dropna=False))
    print("\nConfidence by sentiment (classified only):")
    conf_stats = (
        classified.groupby("sentiment")["sentiment_confidence"]
        .agg(["count", "mean", "median", "min", "max"])
        .round(3)
    )
    print(conf_stats)
    print(f"\nMean sentiment_score (classified only): {classified['sentiment_score'].mean():.3f}")


def write_partial_output(df: pd.DataFrame, cache: Dict[str, Any], output_path: Path) -> None:
    out = apply_results(df, cache)
    out.to_csv(output_path, index=False, encoding="utf-8-sig")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Batch sentiment classification with Gemini for primary bond messages.")
    parser.add_argument("--input", type=Path, default=Path("primary_sentiment_ready/sentiment_universe_primary_v10.csv"))
    parser.add_argument("--output", type=Path, default=Path("sentiment_scored_gemini_v10.csv"))
    parser.add_argument("--cache", type=Path, default=Path("sentiment_cache_gemini_v10.json"))
    parser.add_argument("--model", type=str, default=DEFAULT_MODEL)
    parser.add_argument("--max-batch-items", type=int, default=DEFAULT_BATCH_ITEMS)
    parser.add_argument("--max-batch-chars", type=int, default=DEFAULT_MAX_BATCH_CHARS)
    parser.add_argument("--max-requests", type=int, default=DEFAULT_MAX_REQUESTS)
    parser.add_argument("--delay", type=float, default=DEFAULT_DELAY)
    parser.add_argument("--min-text-len", type=int, default=DEFAULT_MIN_TEXT_LEN)
    parser.add_argument("--save-every", type=int, default=DEFAULT_SAVE_EVERY)
    parser.add_argument("--output-write-every", type=int, default=DEFAULT_OUTPUT_WRITE_EVERY)
    parser.add_argument("--max-retries", type=int, default=DEFAULT_MAX_RETRIES)
    parser.add_argument("--max-output-tokens", type=int, default=DEFAULT_MAX_OUTPUT_TOKENS)
    parser.add_argument("--api-key", type=str, default=os.getenv("GEMINI_API_KEY", ""))
    parser.add_argument("--no-confirm", action="store_true")
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    if not args.api_key:
        raise ValueError("API key not found. Pass --api-key or set GEMINI_API_KEY.")
    if genai is None:
        raise ImportError("google-genai is not installed. Install it with `pip install google-genai`.")

    client = genai.Client(api_key=args.api_key)

    df = pd.read_csv(args.input, low_memory=False)
    if "mention_uid" not in df.columns:
        raise KeyError("Input file must contain 'mention_uid'.")

    df["mention_uid"] = df["mention_uid"].astype(str)

    csv_progress = load_progress_from_output_csv(args.output)
    json_cache = load_json(args.cache)

    if csv_progress:
        cache = csv_progress
        progress_source = f"output CSV ({args.output})"
    else:
        cache = json_cache
        progress_source = f"JSON cache ({args.cache})"

    pending = prepare_pending_items(df, cache, min_text_len=args.min_text_len)
    batches = make_batches(pending, max_batch_items=args.max_batch_items, max_batch_chars=args.max_batch_chars)

    if len(batches) > args.max_requests:
        batches = batches[: args.max_requests]
        print(f"⚠ Limited to first {args.max_requests} requests due to max_requests.")

    print("=" * 70)
    print("Gemini sentiment classification for primary bond messages")
    print("=" * 70)
    print(f"Input: {args.input}")
    print(f"Output: {args.output}")
    print(f"Cache: {args.cache}")
    print(f"Model: {args.model}")
    print(f"Rows in input: {len(df):,}")
    print(f"Progress source: {progress_source}")
    print(f"Cached/skipped already present: {len(cache):,}")
    print(f"Pending items: {len(pending):,}")
    print(f"Planned batches this run: {len(batches):,}")
    print(f"Delay between requests: {args.delay:.1f}s")

    if batches:
        batch_sizes = [len(b) for b in batches]
        batch_chars = [sum(x.estimated_chars for x in b) for b in batches]
        print(f"Batch size stats (items): min={min(batch_sizes)}, median={pd.Series(batch_sizes).median():.0f}, max={max(batch_sizes)}")
        print(f"Batch char stats:         min={min(batch_chars)}, median={pd.Series(batch_chars).median():.0f}, max={max(batch_chars)}")

    if batches and not args.no_confirm:
        answer = input(f"Send {len(batches)} requests to Gemini? (y/n): ").strip().lower()
        if answer != "y":
            print("Cancelled by user.")
            return

    try:
        for batch_idx, batch in enumerate(batches, start=1):
            chars_now = sum(x.estimated_chars for x in batch)
            print(f"\nBatch {batch_idx}/{len(batches)} | items={len(batch)} | est_chars={chars_now}")

            batch_results = classify_batch(
                client=client,
                model_name=args.model,
                batch=batch,
                max_retries=args.max_retries,
                max_output_tokens=args.max_output_tokens,
                min_interval=args.delay,
            )
            cache.update(batch_results)

            classified_now = sum(1 for x in batch_results.values() if x.get("status") == "classified")
            failed_now = len(batch_results) - classified_now
            print(f"  classified: {classified_now} | non-classified: {failed_now}")

            if batch_idx % args.save_every == 0:
                save_json(args.cache, cache)
                print(f"  cache saved: {args.cache}")

            if batch_idx % args.output_write_every == 0:
                write_partial_output(df, cache, args.output)
                print(f"  partial CSV saved: {args.output}")

    except KeyboardInterrupt:
        print("\nInterrupted by user. Saving current progress...")

    finally:
        save_json(args.cache, cache)
        write_partial_output(df, cache, args.output)
        print(f"\nProgress saved to cache: {args.cache}")
        print(f"Progress saved to CSV:   {args.output}")

    out = apply_results(df, cache)
    print_summary(out)


if __name__ == "__main__":
    main()
