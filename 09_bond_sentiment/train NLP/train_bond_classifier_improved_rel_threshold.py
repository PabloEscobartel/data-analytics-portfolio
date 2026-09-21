#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Fine-tune RuBERT / DAPT-RuBERT for Russian bond-market relevance and sentiment classification.

Main changes vs. the previous version:
  1) Correct text-pair training: issuer and text are separated with token_type_ids.
  2) truncation="only_second": the issuer is not truncated, long message text is truncated.
  3) Softer class weights for sentiment to reduce false positive/negative overprediction.
  4) Lower label smoothing for sentiment.
  5) Validation-time neutral-zone tuning for sentiment:
        if positive/negative is weak or too close to neutral -> neutral.
  6) Validation-time relevance threshold tuning:
        classify as yes only if P(yes) >= tuned threshold.
  7) Saved inference_config.json with best neutral-zone / relevance-threshold params.
  8) Confusion matrix printed for both tasks.

Usage:
  python train_bond_classifier_improved.py --train \
    --training-data training_data_combined.csv \
    --base-model model_dapt_clean_from118k

  python train_bond_classifier_improved.py --predict \
    --input sentiment_universe.csv \
    --output sentiment_classified.csv

Requirements:
  pip install torch transformers scikit-learn pandas numpy
"""

import argparse
import json
import os
import random
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import torch
from sklearn.metrics import classification_report, confusion_matrix, f1_score
from sklearn.model_selection import train_test_split
from sklearn.utils.class_weight import compute_class_weight
from torch.utils.data import DataLoader, Dataset
from transformers import (
    AutoModelForSequenceClassification,
    AutoTokenizer,
    get_linear_schedule_with_warmup,
)


# ============================================================================
# Configuration
# ============================================================================

BASE_MODEL = "DeepPavlov/rubert-base-cased-conversational"
DAPT_MODEL_DIR = "model_dapt_clean_from118k"
DEFAULT_FINE_TUNE_BASE_MODEL = DAPT_MODEL_DIR if Path(DAPT_MODEL_DIR).exists() else BASE_MODEL

MODEL_BUNDLE_DIR = "bond_classifier_models_f1_0753"
RELEVANCE_MODEL_DIR = str(Path(MODEL_BUNDLE_DIR) / "model_relevance")
SENTIMENT_MODEL_DIR = str(Path(MODEL_BUNDLE_DIR) / "model_sentiment")

RELEVANCE_LABELS = {"no": 0, "yes": 1}
SENTIMENT_LABELS = {"negative": 0, "neutral": 1, "positive": 2}

# For Smart-Lab posts 384 is usually safer than 256.
MAX_LEN = 384
BATCH_SIZE = 16
EPOCHS = 6
LR = 2e-5
WARMUP_RATIO = 0.10
WEIGHT_DECAY = 0.01
GRAD_CLIP = 1.0
SEED = 42
LOG_EVERY_STEPS = 25

# The CSV only has sentiment_confidence. For relevance we keep the old 0.80 filter.
# For sentiment we use a cleaner subset by default.
RELEVANCE_MIN_CONFIDENCE = 0.80
SENTIMENT_MIN_CONFIDENCE = 0.90

# Full balanced weights often make the model overpredict minority classes.
# For sentiment, 0.0 = no weights, 0.5 = sqrt-balanced, 1.0 = fully balanced.
RELEVANCE_CLASS_WEIGHT_POWER = 1.0
SENTIMENT_CLASS_WEIGHT_POWER = 0.5

RELEVANCE_LABEL_SMOOTHING = 0.05
SENTIMENT_LABEL_SMOOTHING = 0.03

# Relevance inference threshold is tuned on validation.
# Higher threshold -> fewer false positives entering sentiment stage.
# Lower threshold -> fewer missed relevant texts.
RELEVANCE_THRESHOLD_GRID = [
    0.20, 0.25, 0.30, 0.35, 0.40, 0.45, 0.50,
    0.55, 0.60, 0.65, 0.70, 0.75, 0.80,
]

DEVICE = torch.device(
    "cuda" if torch.cuda.is_available()
    else "mps" if torch.backends.mps.is_available()
    else "cpu"
)


# ============================================================================
# Utilities
# ============================================================================

def set_seed(seed: int = SEED) -> None:
    random.seed(seed)
    np.random.seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(seed)


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def label_names_from_map(label_map: Dict[str, int]) -> List[str]:
    return [k for k, _ in sorted(label_map.items(), key=lambda x: x[1])]


def make_class_weights(
    labels: List[int],
    n_classes: int,
    power: float = 1.0,
) -> Optional[torch.Tensor]:
    """
    power=0.0 -> no class weights
    power=0.5 -> softened/sqrt-balanced weights
    power=1.0 -> fully balanced weights
    """
    if power <= 0:
        return None

    labels_np = np.asarray(labels)
    classes = np.arange(n_classes)
    weights = compute_class_weight("balanced", classes=classes, y=labels_np)

    if power != 1.0:
        weights = np.power(weights, power)
        weights = weights / weights.mean()

    return torch.tensor(weights, dtype=torch.float32).to(DEVICE)


def print_confusion(y_true: Iterable[int], y_pred: Iterable[int], label_names: List[str]) -> None:
    cm = confusion_matrix(y_true, y_pred, labels=list(range(len(label_names))))
    cm_df = pd.DataFrame(
        cm,
        index=[f"true_{x}" for x in label_names],
        columns=[f"pred_{x}" for x in label_names],
    )
    print("Confusion matrix:")
    print(cm_df.to_string())


# ============================================================================
# Data
# ============================================================================

class TextPairDataset(Dataset):
    """BERT text-pair dataset: [CLS] issuer [SEP] text [SEP]."""

    def __init__(
        self,
        texts: List[str],
        issuers: List[str],
        labels: List[int],
        tokenizer,
        max_len: int = MAX_LEN,
    ):
        self.texts = texts
        self.issuers = issuers
        self.labels = labels
        self.tokenizer = tokenizer
        self.max_len = max_len

    def __len__(self) -> int:
        return len(self.texts)

    def __getitem__(self, idx: int) -> Dict[str, torch.Tensor]:
        encoding = self.tokenizer(
            self.issuers[idx],
            self.texts[idx],
            max_length=self.max_len,
            padding="max_length",
            truncation="only_second",
            return_tensors="pt",
            return_token_type_ids=True,
        )

        item = {
            "input_ids": encoding["input_ids"].squeeze(0),
            "attention_mask": encoding["attention_mask"].squeeze(0),
            "labels": torch.tensor(self.labels[idx], dtype=torch.long),
        }
        if "token_type_ids" in encoding:
            item["token_type_ids"] = encoding["token_type_ids"].squeeze(0)
        return item


def load_training_data(csv_path: str) -> pd.DataFrame:
    """Load and normalize the Gemini-labeled data."""
    df = pd.read_csv(csv_path, low_memory=False)

    required = {
        "text",
        "issuer",
        "sentiment_confidence",
        "relevance_to_placement",
        "sentiment",
    }
    missing = sorted(required - set(df.columns))
    if missing:
        raise KeyError(f"Training CSV is missing required columns: {missing}")

    df["sentiment_confidence"] = pd.to_numeric(df["sentiment_confidence"], errors="coerce")
    df = df[df["text"].notna() & (df["text"].astype(str).str.len() > 10)].copy()
    df = df[df["issuer"].notna() & (df["issuer"].astype(str).str.len() > 0)].copy()

    df["text"] = df["text"].astype(str)
    df["issuer"] = df["issuer"].astype(str)
    df["relevance_to_placement"] = df["relevance_to_placement"].astype(str).str.lower().str.strip()
    df["sentiment"] = df["sentiment"].astype(str).str.lower().str.strip()

    df = df[df["relevance_to_placement"].isin(RELEVANCE_LABELS)].copy()
    df = df[df["sentiment"].isin(SENTIMENT_LABELS)].copy()

    if "sentiment_status" in df.columns:
        df = df[df["sentiment_status"] == "classified"].copy()

    print(f"Loaded training data: {len(df)} examples before task-specific confidence filters")
    print(f"  Relevance raw: {df['relevance_to_placement'].value_counts().to_dict()}")
    print(f"  Sentiment raw: {df['sentiment'].value_counts().to_dict()}")

    return df


def build_task_dataframe(df: pd.DataFrame, task: str) -> pd.DataFrame:
    if task == "relevance":
        out = df[df["sentiment_confidence"] >= RELEVANCE_MIN_CONFIDENCE].copy()
        label_col = "relevance_to_placement"
    elif task == "sentiment":
        out = df[
            (df["relevance_to_placement"] == "yes")
            & (df["sentiment_confidence"] >= SENTIMENT_MIN_CONFIDENCE)
        ].copy()
        label_col = "sentiment"
    else:
        raise ValueError(f"Unknown task: {task}")

    print(f"\nTask={task}: {len(out)} examples")
    print(f"  Labels: {out[label_col].value_counts().to_dict()}")
    return out


# ============================================================================
# Sentiment neutral zone
# ============================================================================

def apply_neutral_zone_np(
    probs: np.ndarray,
    neutral_id: int = 1,
    margin: float = 0.08,
    min_non_neutral_prob: float = 0.45,
) -> np.ndarray:
    """
    Converts weak positive/negative predictions to neutral.

    Rules:
      - If argmax is neutral -> neutral.
      - If argmax is positive/negative but its probability is too low -> neutral.
      - If positive/negative is too close to neutral -> neutral.
    """
    preds = probs.argmax(axis=1).astype(int)

    for i in range(len(preds)):
        pred = int(preds[i])
        if pred == neutral_id:
            continue
        if probs[i, pred] < min_non_neutral_prob:
            preds[i] = neutral_id
            continue
        if probs[i, pred] - probs[i, neutral_id] < margin:
            preds[i] = neutral_id

    return preds


def tune_neutral_zone(
    probs: np.ndarray,
    y_true: np.ndarray,
    neutral_id: int,
) -> Tuple[Dict[str, float], np.ndarray, float]:
    """Grid-search neutral-zone params on validation set by macro F1."""
    margins = [0.00, 0.03, 0.05, 0.08, 0.10, 0.12, 0.15, 0.18, 0.20]
    min_probs = [0.00, 0.35, 0.40, 0.45, 0.50, 0.55, 0.60]

    best_score = -1.0
    best_params = {"margin": 0.0, "min_non_neutral_prob": 0.0}
    best_preds = probs.argmax(axis=1).astype(int)

    for margin in margins:
        for min_prob in min_probs:
            preds = apply_neutral_zone_np(
                probs,
                neutral_id=neutral_id,
                margin=margin,
                min_non_neutral_prob=min_prob,
            )
            score = f1_score(y_true, preds, average="macro")
            if score > best_score:
                best_score = score
                best_params = {
                    "margin": float(margin),
                    "min_non_neutral_prob": float(min_prob),
                }
                best_preds = preds

    return best_params, best_preds, float(best_score)



# ============================================================================
# Relevance threshold
# ============================================================================

def apply_relevance_threshold_np(
    probs: np.ndarray,
    yes_id: int = 1,
    threshold: float = 0.50,
) -> np.ndarray:
    """Binary relevance decision: yes if P(yes) >= threshold, else no."""
    p_yes = probs[:, yes_id]
    return (p_yes >= threshold).astype(int)


def tune_relevance_threshold(
    probs: np.ndarray,
    y_true: np.ndarray,
    yes_id: int,
) -> Tuple[Dict[str, float], np.ndarray, float]:
    """Grid-search relevance threshold on validation set by macro F1."""
    best_score = -1.0
    best_params = {"yes_threshold": 0.50}
    best_preds = apply_relevance_threshold_np(probs, yes_id=yes_id, threshold=0.50)

    for threshold in RELEVANCE_THRESHOLD_GRID:
        preds = apply_relevance_threshold_np(probs, yes_id=yes_id, threshold=threshold)
        score = f1_score(y_true, preds, average="macro")
        if score > best_score:
            best_score = score
            best_params = {"yes_threshold": float(threshold)}
            best_preds = preds

    return best_params, best_preds, float(best_score)

# ============================================================================
# Model forward / evaluation
# ============================================================================

def model_forward(model, batch_or_inputs: Dict[str, torch.Tensor]):
    """Forward pass with graceful fallback for models that do not accept token_type_ids."""
    inputs = {k: v for k, v in batch_or_inputs.items() if k != "labels"}
    try:
        return model(**inputs)
    except TypeError:
        inputs.pop("token_type_ids", None)
        return model(**inputs)


def collect_probs_and_labels(model, loader: DataLoader) -> Tuple[np.ndarray, np.ndarray]:
    model.eval()
    all_probs: List[np.ndarray] = []
    all_labels: List[np.ndarray] = []

    with torch.no_grad():
        for batch in loader:
            batch = {k: v.to(DEVICE) for k, v in batch.items()}
            outputs = model_forward(model, batch)
            probs = torch.softmax(outputs.logits, dim=-1).cpu().numpy()
            labels = batch["labels"].cpu().numpy()
            all_probs.append(probs)
            all_labels.append(labels)

    return np.concatenate(all_probs, axis=0), np.concatenate(all_labels, axis=0)


def evaluate_from_probs(
    probs: np.ndarray,
    y_true: np.ndarray,
    task: str,
    label_map: Dict[str, int],
    fixed_neutral_params: Optional[Dict[str, float]] = None,
    fixed_relevance_params: Optional[Dict[str, float]] = None,
) -> Tuple[float, np.ndarray, Dict[str, float], Dict[str, float], str]:
    label_names = label_names_from_map(label_map)

    neutral_params: Dict[str, float] = {}
    relevance_params: Dict[str, float] = {}

    if task == "sentiment":
        neutral_id = label_map["neutral"]
        if fixed_neutral_params is None:
            neutral_params, preds, macro_f1 = tune_neutral_zone(probs, y_true, neutral_id=neutral_id)
        else:
            neutral_params = fixed_neutral_params
            preds = apply_neutral_zone_np(
                probs,
                neutral_id=neutral_id,
                margin=float(neutral_params.get("margin", 0.0)),
                min_non_neutral_prob=float(neutral_params.get("min_non_neutral_prob", 0.0)),
            )
            macro_f1 = f1_score(y_true, preds, average="macro")
    elif task == "relevance":
        yes_id = label_map["yes"]
        if fixed_relevance_params is None:
            relevance_params, preds, macro_f1 = tune_relevance_threshold(probs, y_true, yes_id=yes_id)
        else:
            relevance_params = fixed_relevance_params
            preds = apply_relevance_threshold_np(
                probs,
                yes_id=yes_id,
                threshold=float(relevance_params.get("yes_threshold", 0.50)),
            )
            macro_f1 = f1_score(y_true, preds, average="macro")
    else:
        preds = probs.argmax(axis=1).astype(int)
        macro_f1 = f1_score(y_true, preds, average="macro")

    report = classification_report(
        y_true,
        preds,
        labels=list(range(len(label_names))),
        target_names=label_names,
        zero_division=0,
    )
    return float(macro_f1), preds, neutral_params, relevance_params, report


# ============================================================================
# Training
# ============================================================================

def train_classifier(
    df: pd.DataFrame,
    task: str,
    label_col: str,
    label_map: Dict[str, int],
    save_dir: str,
    base_model: str,
    max_len: int = MAX_LEN,
    batch_size: int = BATCH_SIZE,
    epochs: int = EPOCHS,
    lr: float = LR,
) -> float:
    label_counts = df[label_col].value_counts()
    missing = sorted(set(label_map) - set(label_counts.index))
    if missing:
        raise ValueError(f"{label_col}: missing classes after filtering: {missing}")
    if label_counts.min() < 2:
        raise ValueError(
            f"{label_col}: each class needs at least 2 examples for stratified split. "
            f"Counts: {label_counts.to_dict()}"
        )

    texts = df["text"].astype(str).tolist()
    issuers = df["issuer"].astype(str).tolist()
    labels = df[label_col].map(label_map).astype(int).tolist()

    label_names = label_names_from_map(label_map)

    if task == "sentiment":
        weight_power = SENTIMENT_CLASS_WEIGHT_POWER
        label_smoothing = SENTIMENT_LABEL_SMOOTHING
    else:
        weight_power = RELEVANCE_CLASS_WEIGHT_POWER
        label_smoothing = RELEVANCE_LABEL_SMOOTHING

    class_weights = make_class_weights(labels, n_classes=len(label_map), power=weight_power)
    loss_fn = torch.nn.CrossEntropyLoss(weight=class_weights, label_smoothing=label_smoothing)

    train_idx, val_idx = train_test_split(
        range(len(texts)),
        test_size=0.15,
        random_state=SEED,
        stratify=labels,
    )

    tokenizer = AutoTokenizer.from_pretrained(base_model)
    model = AutoModelForSequenceClassification.from_pretrained(
        base_model,
        num_labels=len(label_map),
        id2label={v: k for k, v in label_map.items()},
        label2id=label_map,
    ).to(DEVICE)

    train_ds = TextPairDataset(
        [texts[i] for i in train_idx],
        [issuers[i] for i in train_idx],
        [labels[i] for i in train_idx],
        tokenizer,
        max_len=max_len,
    )
    val_ds = TextPairDataset(
        [texts[i] for i in val_idx],
        [issuers[i] for i in val_idx],
        [labels[i] for i in val_idx],
        tokenizer,
        max_len=max_len,
    )

    train_loader = DataLoader(train_ds, batch_size=batch_size, shuffle=True)
    val_loader = DataLoader(val_ds, batch_size=batch_size * 2, shuffle=False)

    print(
        f"\n  {task}/{label_col}: {len(df)} examples, "
        f"train={len(train_ds)}, val={len(val_ds)}, batches/epoch={len(train_loader)}",
        flush=True,
    )
    print(f"  Class counts: {label_counts.to_dict()}", flush=True)
    if class_weights is None:
        print("  Class weights: disabled", flush=True)
    else:
        print(f"  Class weights: {np.round(class_weights.detach().cpu().numpy(), 4).tolist()}", flush=True)
    print(f"  Label smoothing: {label_smoothing}", flush=True)

    optimizer = torch.optim.AdamW(model.parameters(), lr=lr, weight_decay=WEIGHT_DECAY)
    total_steps = len(train_loader) * epochs
    scheduler = get_linear_schedule_with_warmup(
        optimizer,
        num_warmup_steps=int(total_steps * WARMUP_RATIO),
        num_training_steps=total_steps,
    )

    best_f1 = -1.0
    best_report = None
    best_preds = None
    best_labels = None
    best_neutral_params: Dict[str, float] = {}
    best_relevance_params: Dict[str, float] = {}

    ensure_dir(save_dir)

    for epoch in range(epochs):
        model.train()
        total_loss = 0.0
        print(f"\n  Epoch {epoch + 1}/{epochs}: training...", flush=True)

        for step, batch in enumerate(train_loader, start=1):
            batch = {k: v.to(DEVICE) for k, v in batch.items()}
            labels_tensor = batch["labels"]

            outputs = model_forward(model, batch)
            loss = loss_fn(outputs.logits, labels_tensor)
            loss.backward()

            torch.nn.utils.clip_grad_norm_(model.parameters(), GRAD_CLIP)
            optimizer.step()
            scheduler.step()
            optimizer.zero_grad(set_to_none=True)

            total_loss += loss.item()
            if step == 1 or step % LOG_EVERY_STEPS == 0 or step == len(train_loader):
                print(
                    f"    batch {step}/{len(train_loader)} "
                    f"loss={total_loss / step:.4f}",
                    flush=True,
                )

        avg_loss = total_loss / max(1, len(train_loader))

        print(f"  Epoch {epoch + 1}/{epochs}: validating...", flush=True)
        probs, y_true = collect_probs_and_labels(model, val_loader)
        val_f1, preds, neutral_params, relevance_params, report = evaluate_from_probs(
            probs,
            y_true,
            task=task,
            label_map=label_map,
        )

        if task == "sentiment":
            print(
                f"  Epoch {epoch + 1}/{epochs} loss={avg_loss:.4f} "
                f"val_macro_f1={val_f1:.4f} "
                f"neutral_zone={neutral_params}",
                flush=True,
            )
        elif task == "relevance":
            print(
                f"  Epoch {epoch + 1}/{epochs} loss={avg_loss:.4f} "
                f"val_macro_f1={val_f1:.4f} "
                f"relevance_threshold={relevance_params}",
                flush=True,
            )
        else:
            print(
                f"  Epoch {epoch + 1}/{epochs} loss={avg_loss:.4f} "
                f"val_macro_f1={val_f1:.4f}",
                flush=True,
            )

        if val_f1 > best_f1:
            best_f1 = val_f1
            best_report = report
            best_preds = preds.copy()
            best_labels = y_true.copy()
            best_neutral_params = neutral_params.copy()
            best_relevance_params = relevance_params.copy()

            model.save_pretrained(save_dir)
            tokenizer.save_pretrained(save_dir)

            with open(os.path.join(save_dir, "label_map.json"), "w", encoding="utf-8") as f:
                json.dump(label_map, f, ensure_ascii=False, indent=2)

            inference_config = {
                "task": task,
                "max_len": max_len,
                "neutral_zone": best_neutral_params,
                "relevance_threshold": best_relevance_params,
                "base_model": base_model,
            }
            with open(os.path.join(save_dir, "inference_config.json"), "w", encoding="utf-8") as f:
                json.dump(inference_config, f, ensure_ascii=False, indent=2)

            print(f"  Saved new best model to {save_dir}", flush=True)

    print(f"\n  Best F1: {best_f1:.4f}")
    print(best_report)
    if best_labels is not None and best_preds is not None:
        print_confusion(best_labels, best_preds, label_names)
    if task == "sentiment":
        print(f"  Best neutral-zone params: {best_neutral_params}")
    if task == "relevance":
        print(f"  Best relevance-threshold params: {best_relevance_params}")

    return best_f1


def train_all(
    training_csv: str,
    base_model: str = DEFAULT_FINE_TUNE_BASE_MODEL,
    max_len: int = MAX_LEN,
    batch_size: int = BATCH_SIZE,
    epochs: int = EPOCHS,
    lr: float = LR,
) -> None:
    df = load_training_data(training_csv)

    print("\n" + "=" * 70)
    print("STAGE 1: Relevance classifier")
    print("=" * 70)
    rel_df = build_task_dataframe(df, task="relevance")
    train_classifier(
        rel_df,
        task="relevance",
        label_col="relevance_to_placement",
        label_map=RELEVANCE_LABELS,
        save_dir=RELEVANCE_MODEL_DIR,
        base_model=base_model,
        max_len=max_len,
        batch_size=batch_size,
        epochs=epochs,
        lr=lr,
    )

    print("\n" + "=" * 70)
    print("STAGE 2: Sentiment classifier, relevant texts only")
    print("=" * 70)
    sent_df = build_task_dataframe(df, task="sentiment")
    train_classifier(
        sent_df,
        task="sentiment",
        label_col="sentiment",
        label_map=SENTIMENT_LABELS,
        save_dir=SENTIMENT_MODEL_DIR,
        base_model=base_model,
        max_len=max_len,
        batch_size=batch_size,
        epochs=epochs,
        lr=lr,
    )


# ============================================================================
# Inference
# ============================================================================

def load_classifier(model_dir: str):
    tokenizer = AutoTokenizer.from_pretrained(model_dir)
    model = AutoModelForSequenceClassification.from_pretrained(model_dir).to(DEVICE)
    model.eval()

    with open(os.path.join(model_dir, "label_map.json"), encoding="utf-8") as f:
        label_map = json.load(f)
    id_to_label = {int(v): k for k, v in label_map.items()}

    config_path = os.path.join(model_dir, "inference_config.json")
    if os.path.exists(config_path):
        with open(config_path, encoding="utf-8") as f:
            inference_config = json.load(f)
    else:
        inference_config = {"task": None, "neutral_zone": {}, "relevance_threshold": {}, "max_len": MAX_LEN}

    return tokenizer, model, id_to_label, inference_config


def predict_batch(
    texts: List[str],
    issuers: List[str],
    tokenizer,
    model,
    id_to_label: Dict[int, str],
    task: Optional[str] = None,
    neutral_zone: Optional[Dict[str, float]] = None,
    relevance_threshold: Optional[Dict[str, float]] = None,
    max_len: int = MAX_LEN,
) -> Tuple[List[str], np.ndarray]:
    encodings = tokenizer(
        issuers,
        texts,
        max_length=max_len,
        padding=True,
        truncation="only_second",
        return_tensors="pt",
        return_token_type_ids=True,
    )
    inputs = {k: v.to(DEVICE) for k, v in encodings.items()}

    with torch.no_grad():
        outputs = model_forward(model, inputs)
        probs = torch.softmax(outputs.logits, dim=-1).cpu().numpy()

    if task == "sentiment" and neutral_zone:
        preds = apply_neutral_zone_np(
            probs,
            neutral_id=1,
            margin=float(neutral_zone.get("margin", 0.0)),
            min_non_neutral_prob=float(neutral_zone.get("min_non_neutral_prob", 0.0)),
        )
    elif task == "relevance" and relevance_threshold:
        preds = apply_relevance_threshold_np(
            probs,
            yes_id=1,
            threshold=float(relevance_threshold.get("yes_threshold", 0.50)),
        )
    else:
        preds = probs.argmax(axis=1).astype(int)

    # Confidence of the chosen label, not necessarily max prob after thresholding.
    confs = probs[np.arange(len(preds)), preds]
    labels = [id_to_label[int(p)] for p in preds]
    return labels, confs


def predict_all(input_csv: str, output_csv: str, batch_size: int = 64) -> None:
    df = pd.read_csv(input_csv, low_memory=False)
    print(f"Input: {len(df)} rows")

    if "text" not in df.columns:
        raise KeyError("Need 'text' column")
    if "issuer" not in df.columns:
        raise KeyError("Need 'issuer' column")

    texts = df["text"].fillna("").astype(str).tolist()
    issuers = df["issuer"].fillna("").astype(str).tolist()

    # Stage 1: relevance
    print("Stage 1: Relevance...")
    rel_tok, rel_model, rel_labels, rel_cfg = load_classifier(RELEVANCE_MODEL_DIR)
    rel_max_len = int(rel_cfg.get("max_len", MAX_LEN))
    rel_threshold = rel_cfg.get("relevance_threshold", {}) or {}
    print(f"  Relevance threshold params: {rel_threshold}")

    all_rel: List[str] = []
    all_rel_conf: List[float] = []

    for i in range(0, len(texts), batch_size):
        batch_texts = texts[i:i + batch_size]
        batch_issuers = issuers[i:i + batch_size]
        labels, confs = predict_batch(
            batch_texts,
            batch_issuers,
            rel_tok,
            rel_model,
            rel_labels,
            task="relevance",
            relevance_threshold=rel_threshold,
            max_len=rel_max_len,
        )
        all_rel.extend(labels)
        all_rel_conf.extend(confs.tolist())
        if i == 0 or (i // batch_size) % 100 == 0:
            print(f"  {i}/{len(texts)}")

    df["relevance_pred"] = all_rel
    df["relevance_conf"] = np.round(all_rel_conf, 4)

    # Stage 2: sentiment, only relevant texts
    print("Stage 2: Sentiment...")
    sent_tok, sent_model, sent_labels, sent_cfg = load_classifier(SENTIMENT_MODEL_DIR)
    sent_max_len = int(sent_cfg.get("max_len", MAX_LEN))
    neutral_zone = sent_cfg.get("neutral_zone", {}) or {}
    print(f"  Neutral-zone params: {neutral_zone}")

    relevant_mask = df["relevance_pred"] == "yes"
    relevant_idx = df.index[relevant_mask].tolist()
    print(f"  {len(relevant_idx)} relevant texts to classify")

    df["sentiment_pred"] = "not_relevant"
    df["sentiment_conf"] = 0.0
    df["sentiment_score"] = 0

    for i in range(0, len(relevant_idx), batch_size):
        batch_idx = relevant_idx[i:i + batch_size]
        batch_texts = [texts[j] for j in batch_idx]
        batch_issuers = [issuers[j] for j in batch_idx]

        labels, confs = predict_batch(
            batch_texts,
            batch_issuers,
            sent_tok,
            sent_model,
            sent_labels,
            task="sentiment",
            neutral_zone=neutral_zone,
            max_len=sent_max_len,
        )

        for j, idx in enumerate(batch_idx):
            df.at[idx, "sentiment_pred"] = labels[j]
            df.at[idx, "sentiment_conf"] = round(float(confs[j]), 4)
            df.at[idx, "sentiment_score"] = {
                "negative": -1,
                "neutral": 0,
                "positive": 1,
            }.get(labels[j], 0)

        if i == 0 or (i // batch_size) % 100 == 0:
            print(f"  {i}/{len(relevant_idx)}")

    df.to_csv(output_csv, index=False, encoding="utf-8-sig")
    print(f"\nSaved: {output_csv}")
    print(f"  Relevant: {int(relevant_mask.sum())} ({relevant_mask.mean() * 100:.1f}%)")
    print("  Sentiment distribution:")
    print(df.loc[relevant_mask, "sentiment_pred"].value_counts().to_string())


# ============================================================================
# Main
# ============================================================================

def main() -> None:
    parser = argparse.ArgumentParser(description="Bond sentiment classifier, improved RuBERT fine-tuning")
    parser.add_argument("--train", action="store_true", help="Train relevance and sentiment classifiers")
    parser.add_argument("--predict", action="store_true", help="Run two-stage inference")
    parser.add_argument("--training-data", default="training_data_combined.csv", help="Training CSV")
    parser.add_argument("--base-model", default=None, help="Base model name or local DAPT model path")
    parser.add_argument("--input", default="sentiment_universe.csv", help="Input CSV for prediction")
    parser.add_argument("--output", default="sentiment_classified.csv", help="Output CSV")
    parser.add_argument("--batch-size", type=int, default=BATCH_SIZE, help="Training batch size")
    parser.add_argument("--predict-batch-size", type=int, default=64, help="Inference batch size")
    parser.add_argument("--epochs", type=int, default=EPOCHS, help="Number of epochs")
    parser.add_argument("--max-len", type=int, default=MAX_LEN, help="Max sequence length")
    parser.add_argument("--lr", type=float, default=LR, help="Learning rate")
    args = parser.parse_args()

    set_seed(SEED)

    # Auto-mount Drive in Colab
    if "google.colab" in sys.modules or os.path.exists("/content"):
        try:
            from google.colab import drive
            if not os.path.ismount("/content/drive"):
                drive.mount("/content/drive")
                print("Google Drive mounted.", flush=True)
        except Exception:
            pass

    # Apple Silicon MPS optimization
    if DEVICE.type == "mps":
        os.environ["PYTORCH_MPS_HIGH_WATERMARK_RATIO"] = "0.0"

    base_model = args.base_model
    if base_model is None:
        base_model = DAPT_MODEL_DIR if Path(DAPT_MODEL_DIR).exists() else BASE_MODEL

    print(f"Device: {DEVICE}")
    print(f"Base model: {base_model}")
    print(f"Max length: {args.max_len}")

    if args.train:
        train_all(
            args.training_data,
            base_model=base_model,
            max_len=args.max_len,
            batch_size=args.batch_size,
            epochs=args.epochs,
            lr=args.lr,
        )

    if args.predict:
        predict_all(args.input, args.output, batch_size=args.predict_batch_size)

    if not args.train and not args.predict:
        print("Specify --train and/or --predict")


if __name__ == "__main__":
    main()
