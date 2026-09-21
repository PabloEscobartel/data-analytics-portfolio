#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Domain-Adaptive Pretraining — Google Colab version (T4, fixed).

Fixes vs previous version:
  - LR 2e-5 instead of 5e-5 (prevents fp16 NaN explosion)
  - Auto-delete old checkpoints (keeps only 2, fits 15GB Drive)
  - NaN detection: stops early if loss explodes
  - Saves every 10000 steps instead of 5000

Setup:
  1. Upload data.db and smartlab_bonds.db to MyDrive/bond_research/
  2. Colab: Runtime → GPU (T4)
  3. Run:
     !pip install transformers -q
     from google.colab import drive
     drive.mount('/content/drive')
     !python /content/drive/MyDrive/bond_research/domain_adapt_colab.py

If Colab disconnects:
     !python /content/drive/MyDrive/bond_research/domain_adapt_colab.py --resume
"""

import argparse
import json
import math
import os
import random
import shutil
import sqlite3
import sys
import time
import traceback
from pathlib import Path
import torch

# Vast.ai / RTX 3060 stability fix:
# disable optimized SDPA kernels that can segfault in some torch/CUDA builds.
try:
    torch.backends.cuda.enable_flash_sdp(False)
    torch.backends.cuda.enable_mem_efficient_sdp(False)
    torch.backends.cuda.enable_math_sdp(True)
except Exception:
    pass

from torch.utils.data import DataLoader, Dataset, Subset
from transformers import (
    AutoModelForMaskedLM,
    AutoTokenizer,
    DataCollatorForLanguageModeling,
)

# ============================================================================
#  Configuration
# ============================================================================

BASE_MODEL = "DeepPavlov/rubert-base-cased-conversational"
MAX_LEN = 128
BATCH_SIZE = 4
GRAD_ACCUM_STEPS = 1         # effective batch = 16
EPOCHS = 2
LR = 2e-5                    # lowered from 5e-5 to prevent fp16 NaN
WARMUP_STEPS = 1000
SAVE_EVERY_STEPS = 1000
MAX_CHECKPOINTS = 2           # auto-delete older ones
MLM_PROBABILITY = 0.15
SEED = 42

DRIVE_BASE = "/workspace/bond_research"
DEFAULT_TG_DB = f"{DRIVE_BASE}/data.db"
DEFAULT_SL_DB = f"{DRIVE_BASE}/smartlab_bonds.db"
DEFAULT_OUTPUT = f"{DRIVE_BASE}/model_dapt"

DEVICE = torch.device("cuda" if torch.cuda.is_available() else "cpu")
USE_FP16 = True


# ============================================================================
#  Data loading
# ============================================================================

def load_corpus(telegram_db: Path, smartlab_db: Path) -> list:
    texts = []
    if telegram_db.exists():
        conn = sqlite3.connect(str(telegram_db))
        rows = conn.execute("SELECT text FROM messages WHERE text IS NOT NULL AND length(text) > 30").fetchall()
        texts.extend([r[0] for r in rows])
        print(f"  Telegram: {len(rows):,} texts")
        conn.close()
    if smartlab_db.exists():
        conn = sqlite3.connect(str(smartlab_db))
        rows = conn.execute(
            "SELECT COALESCE(title,'') || ' ' || COALESCE(text,'') FROM posts WHERE text IS NOT NULL AND length(text) > 30"
        ).fetchall()
        texts.extend([r[0] for r in rows])
        print(f"  SL posts: {len(rows):,} texts")
        rows = conn.execute("SELECT text FROM comments WHERE text IS NOT NULL AND length(text) > 30").fetchall()
        texts.extend([r[0] for r in rows])
        print(f"  SL comments: {len(rows):,} texts")
        conn.close()
    print(f"  Total corpus: {len(texts):,} texts")
    return texts


class TextDataset(Dataset):
    def __init__(self, texts, tokenizer, max_len=MAX_LEN):
        self.texts = texts
        self.tokenizer = tokenizer
        self.max_len = max_len

    def __len__(self):
        return len(self.texts)

    def __getitem__(self, idx):
        enc = self.tokenizer(
            self.texts[idx],
            max_length=self.max_len,
            padding="max_length",
            truncation=True,
            return_special_tokens_mask=True,
            return_tensors="pt",
        )
        return {k: v.squeeze(0) for k, v in enc.items()}


# ============================================================================
#  Checkpoint management
# ============================================================================

def find_latest_checkpoint(output_dir: str):
    if not os.path.exists(output_dir):
        return None, None

    checkpoints = [
        d for d in os.listdir(output_dir)
        if d.startswith("checkpoint-") and os.path.isdir(os.path.join(output_dir, d))
    ]

    if not checkpoints:
        return None, None

    latest = max(checkpoints, key=lambda d: int(d.split("-")[1]))
    ckpt_path = os.path.join(output_dir, latest)

    state_file = os.path.join(ckpt_path, "training_state.json")
    if not os.path.exists(state_file):
        return ckpt_path, None

    with open(state_file, "r", encoding="utf-8") as f:
        state = json.load(f)

    return ckpt_path, state


def cleanup_old_checkpoints(output_dir: str, keep: int = MAX_CHECKPOINTS):
    if not os.path.exists(output_dir):
        return
    all_ckpts = sorted(
        [d for d in os.listdir(output_dir)
         if d.startswith("checkpoint-") and os.path.isdir(os.path.join(output_dir, d))],
        key=lambda d: int(d.split("-")[1]),
    )
    while len(all_ckpts) > keep:
        old_name = all_ckpts.pop(0)
        old_path = os.path.join(output_dir, old_name)
        shutil.rmtree(old_path, ignore_errors=True)
        print(f"  🗑 Deleted old: {old_name}", flush=True)


def save_checkpoint(model, tokenizer, output_dir, global_step, epoch, step_in_epoch):
    ckpt = os.path.join(output_dir, f"checkpoint-{global_step}")

    model.save_pretrained(ckpt, safe_serialization=False)
    tokenizer.save_pretrained(ckpt)

    state = {
        "epoch": epoch,
        "step_in_epoch": step_in_epoch,
        "global_step": global_step,
    }

    with open(os.path.join(ckpt, "training_state.json"), "w", encoding="utf-8") as f:
        json.dump(state, f)

    print(
        f"  ✓ Saved: {ckpt} "
        f"(epoch={epoch + 1}, step_in_epoch={step_in_epoch}, global_step={global_step})",
        flush=True,
    )

    cleanup_old_checkpoints(output_dir, keep=MAX_CHECKPOINTS)


def make_epoch_indices(dataset_len: int, epoch: int) -> list:
    indices = list(range(dataset_len))
    random.Random(SEED + epoch).shuffle(indices)
    return indices

# ============================================================================
#  Training
# ============================================================================

def train_mlm(
    texts: list,
    base_model: str,
    output_dir: str,
    epochs: int,
    batch_size: int,
    resume: bool = False,
):
    print(f"\nDevice: {DEVICE}", flush=True)
    print(f"FP16: {USE_FP16}", flush=True)
    print(f"Batch: {batch_size} x {GRAD_ACCUM_STEPS} accum = {batch_size * GRAD_ACCUM_STEPS} effective", flush=True)
    print(f"LR: {LR}", flush=True)
    print(f"Max length: {MAX_LEN}", flush=True)
    print(f"Base model: {base_model}", flush=True)
    print(f"Corpus: {len(texts):,} texts, {epochs} epochs", flush=True)

    start_step = 0
    start_epoch = 0
    start_step_in_epoch = 0
    model_to_load = base_model

    if resume:
        ckpt_path, state = find_latest_checkpoint(output_dir)
        if ckpt_path and state:
            model_to_load = ckpt_path
            start_step = state["global_step"]
            start_epoch = state["epoch"]
            start_step_in_epoch = state["step_in_epoch"]

            print(f"\nResuming from {ckpt_path}", flush=True)
            print(
                f"  epoch={start_epoch + 1}, "
                f"step_in_epoch={start_step_in_epoch}, "
                f"global_step={start_step}",
                flush=True,
            )
        elif ckpt_path and not state:
            print(
                f"\nFound checkpoint {ckpt_path}, but no training_state.json. "
                f"Cannot exact-resume text position. Start fresh or use it as --base-model.",
                flush=True,
            )
            return
        else:
            print("\nNo checkpoint found, starting fresh.", flush=True)

    tokenizer = AutoTokenizer.from_pretrained(model_to_load)
    model = AutoModelForMaskedLM.from_pretrained(model_to_load, attn_implementation="eager").to(DEVICE)

    dataset = TextDataset(texts, tokenizer)
    collator = DataCollatorForLanguageModeling(
        tokenizer=tokenizer, mlm=True, mlm_probability=MLM_PROBABILITY,
    )

    optimizer = torch.optim.AdamW(model.parameters(), lr=LR, weight_decay=0.01)
    scaler = torch.amp.GradScaler("cuda", enabled=USE_FP16)

    steps_per_epoch = math.ceil(len(dataset) / batch_size)
    total_steps = steps_per_epoch * epochs

    def lr_lambda(step):
        actual = step + start_step
        if actual < WARMUP_STEPS:
            return actual / max(1, WARMUP_STEPS)
        return max(0.0, 1.0 - (actual - WARMUP_STEPS) / max(1, total_steps - WARMUP_STEPS))

    scheduler = torch.optim.lr_scheduler.LambdaLR(optimizer, lr_lambda)
    os.makedirs(output_dir, exist_ok=True)

    target_global_step = total_steps

    print(f"Steps per epoch: {steps_per_epoch:,}", flush=True)
    print(f"Total steps: {total_steps:,}", flush=True)
    if start_step > 0:
        print(
            f"Exact resume from global step {start_step}; "
            f"continuing from epoch {start_epoch + 1}, step {start_step_in_epoch}.",
            flush=True,
        )
    print("", flush=True)

    global_step = start_step
    nan_count = 0
    t0 = time.time()
    model.train()
    optimizer.zero_grad()

    for epoch in range(start_epoch, epochs):
        indices = make_epoch_indices(len(dataset), epoch)

        if epoch == start_epoch and start_step_in_epoch > 0:
            skip_items = start_step_in_epoch * batch_size
            indices = indices[skip_items:]
            print(
                f"  Skipped {skip_items:,} items; "
                f"{len(indices):,} items remaining in epoch {epoch + 1}",
                flush=True,
            )

        subset = Subset(dataset, indices)

        loader = DataLoader(
            subset,
            batch_size=batch_size,
            shuffle=False,
            collate_fn=collator,
            num_workers=0,
            pin_memory=False,
        )

        epoch_loss = 0.0
        epoch_counted = 0
        step_in_epoch = start_step_in_epoch if epoch == start_epoch else 0

        for step, batch in enumerate(loader):
            if global_step >= target_global_step:
                break

            batch = {k: v.to(DEVICE) for k, v in batch.items()}

            with torch.amp.autocast("cuda", enabled=USE_FP16):
                outputs = model(**batch)
                loss = outputs.loss / GRAD_ACCUM_STEPS

            # NaN detection
            if math.isnan(loss.item()):
                nan_count += 1
                if nan_count >= 10:
                    print(f"\n  ❌ Loss NaN for 10 steps at step {global_step}. Stopping.", flush=True)
                    print(f"  Last good checkpoint is on Drive.", flush=True)
                    return
                optimizer.zero_grad()
                global_step += 1
                continue
            else:
                nan_count = 0

            scaler.scale(loss).backward()
            epoch_loss += loss.item() * GRAD_ACCUM_STEPS
            epoch_counted += 1

            if epoch_counted % GRAD_ACCUM_STEPS == 0:
                scaler.unscale_(optimizer)
                torch.nn.utils.clip_grad_norm_(model.parameters(), 1.0)
                scaler.step(optimizer)
                scaler.update()
                scheduler.step()
                optimizer.zero_grad()

            global_step += 1
            step_in_epoch += 1

            if global_step == start_step + 1 or global_step % 100 == 0:
                avg = epoch_loss / max(epoch_counted, 1)
                elapsed = time.time() - t0
                speed = epoch_counted / elapsed if elapsed > 0 else 0
                remaining = total_steps - (global_step - start_step)
                eta_h = remaining / speed / 3600 if speed > 0 else 0
                mem = torch.cuda.max_memory_allocated() / 1e9 if torch.cuda.is_available() else 0
                print(
                    f"  Epoch {epoch+1}/{epochs}  "
                    f"step {global_step}/{total_steps}  "
                    f"loss={avg:.4f}  "
                    f"{speed:.1f} steps/s  "
                    f"VRAM={mem:.1f}GB  "
                    f"ETA={eta_h:.1f}h",
                    flush=True,
                )

            if global_step % SAVE_EVERY_STEPS == 0:
                if epoch_counted % GRAD_ACCUM_STEPS != 0:
                    scaler.unscale_(optimizer)
                    torch.nn.utils.clip_grad_norm_(model.parameters(), 1.0)
                    scaler.step(optimizer)
                    scaler.update()
                    optimizer.zero_grad()

                save_checkpoint(
                    model=model,
                    tokenizer=tokenizer,
                    output_dir=output_dir,
                    global_step=global_step,
                    epoch=epoch,
                    step_in_epoch=step_in_epoch,
                )

        if epoch_counted % GRAD_ACCUM_STEPS != 0:
            scaler.unscale_(optimizer)
            torch.nn.utils.clip_grad_norm_(model.parameters(), 1.0)
            scaler.step(optimizer)
            scaler.update()
            optimizer.zero_grad()

        start_step_in_epoch = 0

        if global_step >= target_global_step:
            break

    model.save_pretrained(output_dir, safe_serialization=False)
    tokenizer.save_pretrained(output_dir)
    total_h = (time.time() - t0) / 3600
    print(f"\nDAPT done in {total_h:.1f}h. Saved to {output_dir}/", flush=True)
    print(f"Next: python train_bond_classifier.py --train --base-model {output_dir}", flush=True)


# ============================================================================
#  Main
# ============================================================================

def main():
    p = argparse.ArgumentParser(description="DAPT on bond corpus (Colab + Drive)")
    p.add_argument("--telegram-db", default=DEFAULT_TG_DB)
    p.add_argument("--smartlab-db", default=DEFAULT_SL_DB)
    p.add_argument("--base-model", default=BASE_MODEL)
    p.add_argument("--output-dir", default=DEFAULT_OUTPUT)
    p.add_argument("--epochs", type=int, default=EPOCHS)
    p.add_argument("--batch-size", type=int, default=BATCH_SIZE)
    p.add_argument("--resume", action="store_true", help="Resume from latest checkpoint")
    args = p.parse_args()

    torch.manual_seed(SEED)

    if "google.colab" in sys.modules or os.path.exists("/content"):
        try:
            from google.colab import drive
            if not os.path.ismount("/content/drive"):
                drive.mount("/content/drive")
                print("Google Drive mounted.", flush=True)
        except Exception:
            pass

    print("Loading corpus...", flush=True)
    texts = load_corpus(Path(args.telegram_db), Path(args.smartlab_db))

    if not texts:
        print("No texts found! Check DB paths on Drive.")
        return

    train_mlm(
        texts, args.base_model, args.output_dir,
        args.epochs, args.batch_size, resume=args.resume,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        raise
