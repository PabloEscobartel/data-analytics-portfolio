#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sample underrepresented text categories for additional Gemini labeling.

Extracts texts that the current training set (sentiment_scored_gemini.csv)
does NOT cover, so the fine-tuned model can handle them properly.

Input:  Telegram DB (data.db) and Smart-Lab DB (smartlab_bonds.db)
        processed by bond_entity_matcher_v9.py

Output: samples_for_gemini.csv — ready to send to Gemini for labeling

Usage:
  python sample_for_labeling.py --telegram-db data.db --smartlab-db smartlab_bonds.db
"""

import argparse
import sqlite3
import pandas as pd
from pathlib import Path

SEED = 42


def sample_category(conn, query, n, category_name):
    """Run query, sample n rows, tag with category."""
    df = pd.read_sql_query(query, conn)
    if len(df) == 0:
        print(f"  {category_name}: 0 found, skipping")
        return pd.DataFrame()
    sampled = df.sample(n=min(n, len(df)), random_state=SEED)
    sampled["sample_category"] = category_name
    print(f"  {category_name}: {len(df):,} available → {len(sampled)} sampled")
    return sampled


def sample_telegram(db_path: Path) -> pd.DataFrame:
    conn = sqlite3.connect(str(db_path))
    frames = []

    # 1. Broker noise — suppressed by v9
    frames.append(sample_category(conn, """
        SELECT r.channel, r.message_id, r.issuer, r.entity_level,
               r.explicit_status, r.resolution_note, m.text
        FROM telegram_matches_resolved r
        JOIN messages m ON r.channel = m.channel AND r.message_id = m.id
        WHERE r.explicit_status = 'suppressed'
          AND length(m.text) >= 30
    """, 2000, "broker_noise"))

    # 2. Ambiguous / multi-issuer
    frames.append(sample_category(conn, """
        SELECT r.channel, r.message_id, r.issuer, r.entity_level,
               r.explicit_status, r.resolution_note, m.text
        FROM telegram_matches_resolved r
        JOIN messages m ON r.channel = m.channel AND r.message_id = m.id
        WHERE r.explicit_status = 'ambiguous'
          AND length(m.text) >= 30
    """, 1600, "multi_issuer"))

    # 3. Macro / OFZ context — issuer mentioned in passing
    frames.append(sample_category(conn, """
        SELECT r.channel, r.message_id, r.issuer, r.entity_level,
               r.explicit_status, r.resolution_note, m.text
        FROM telegram_matches_resolved r
        JOIN messages m ON r.channel = m.channel AND r.message_id = m.id
        WHERE r.issuer IS NOT NULL
          AND (lower(m.text) LIKE '%офз%'
               OR lower(m.text) LIKE '%ставк%'
               OR lower(m.text) LIKE '%ключев%'
               OR lower(m.text) LIKE '%инфляц%'
               OR lower(m.text) LIKE '%цб %')
          AND length(m.text) >= 50
    """, 1200, "macro_ofz_context"))

    # 4. Short texts (<50 chars) with issuer match
    frames.append(sample_category(conn, """
        SELECT r.channel, r.message_id, r.issuer, r.entity_level,
               r.explicit_status, r.resolution_note, m.text
        FROM telegram_matches_resolved r
        JOIN messages m ON r.channel = m.channel AND r.message_id = m.id
        WHERE r.issuer IS NOT NULL
          AND r.match_origin = 'explicit'
          AND length(m.text) BETWEEN 10 AND 50
    """, 800, "short_text"))

    # 5. Unresolved texts with bond keywords — potential false negatives
    frames.append(sample_category(conn, """
        SELECT r.channel, r.message_id, 'UNKNOWN' as issuer, NULL as entity_level,
               r.explicit_status, r.resolution_note, m.text
        FROM telegram_matches_resolved r
        JOIN messages m ON r.channel = m.channel AND r.message_id = m.id
        WHERE r.issuer IS NULL
          AND r.explicit_status = 'none'
          AND (lower(m.text) LIKE '%облигац%' OR lower(m.text) LIKE '%размещен%')
          AND length(m.text) >= 50
    """, 800, "unresolved_with_bond_kw"))

    conn.close()
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def sample_smartlab(db_path: Path) -> pd.DataFrame:
    conn = sqlite3.connect(str(db_path))
    frames = []

    # 6. Smart-Lab posts (completely absent from training data)
    frames.append(sample_category(conn, """
        SELECT r.post_id, r.issuer, r.entity_level, r.explicit_status,
               p.title, p.text,
               'smartlab_post' as source_type
        FROM smartlab_post_matches_resolved r
        JOIN posts p ON r.post_id = p.post_id
        WHERE r.issuer IS NOT NULL
          AND length(p.text) >= 30
    """, 2000, "smartlab_post"))

    # 7. Smart-Lab comments
    frames.append(sample_category(conn, """
        SELECT r.comment_id, r.post_id, r.issuer, r.entity_level, r.explicit_status,
               c.text,
               p.title as post_title,
               'smartlab_comment' as source_type
        FROM smartlab_comment_matches_resolved r
        JOIN comments c ON r.comment_id = c.comment_id
        LEFT JOIN posts p ON r.post_id = p.post_id
        WHERE r.issuer IS NOT NULL
          AND length(c.text) >= 20
    """, 1200, "smartlab_comment"))

    conn.close()
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def main():
    p = argparse.ArgumentParser(description="Sample texts for additional Gemini labeling")
    p.add_argument("--telegram-db", default="data.db")
    p.add_argument("--smartlab-db", default="smartlab_bonds.db")
    p.add_argument("--output", default="samples_for_gemini.csv")
    args = p.parse_args()

    print("Sampling Telegram categories...")
    tg = sample_telegram(Path(args.telegram_db))

    print("\nSampling Smart-Lab categories...")
    sl = sample_smartlab(Path(args.smartlab_db))

    all_samples = pd.concat([tg, sl], ignore_index=True)

    # Deduplicate by text
    before = len(all_samples)
    all_samples = all_samples.drop_duplicates(subset=["text"], keep="first")
    print(f"\nDeduplicated: {before} → {len(all_samples)}")

    # Summary
    print(f"\n{'='*50}")
    print(f"TOTAL SAMPLES: {len(all_samples)}")
    print(f"{'='*50}")
    print(all_samples["sample_category"].value_counts().to_string())

    # Estimate Gemini cost
    n = len(all_samples)
    print(f"\nGemini estimate:")
    print(f"  At 1 text/request:  {n} requests → {n // 500 + 1} days at 500/day")
    print(f"  At 10 texts/batch:  {n // 10} requests → {n // 10 // 500 + 1} days at 500/day")

    all_samples.to_csv(args.output, index=False, encoding="utf-8-sig")
    print(f"\nSaved: {args.output}")


if __name__ == "__main__":
    main()
