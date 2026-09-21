# NLP Pipeline

This folder contains the text-processing and classification workflow used to produce `classified.csv`, the main NLP input for the final analysis pipeline.

## Purpose

The NLP pipeline links social-media text to bond issuers or offerings, labels relevance and sentiment, trains classifiers, and applies the trained models to the full matched text universe.

## Main Files

```text
bond_entity_matcher_v10.py
bond_match_audit_v3.py
prepare_primary_sentiment_universe_close_time.py
sample_for_labeling.py
classify_samples_gemini.py
classify_sentiment_gemini_genai_resume_from_csv.py
check_batch_contamination_gemini_genai.py
mine_colloquial_aliases.py
audit_manual_review.py
domain_adapt_colab (2).py
train_bond_classifier_final.py
train_bond_classifier_improved_rel_threshold.py
```

## Expected Data and Model Files

The following files are required for full reproduction but are not tracked in Git:

```text
data.db
smartlab_bonds.db
panel_final_v3.xlsx
issuer_aliases.xlsx
issue_aliases_candidates_fixed.xlsx
all_matches.csv
training_data_combined.csv
classified.csv
model_dapt_clean_from118k/
bond_classifier_models_f1_0753/
```

## Workflow

1. Run `bond_entity_matcher_v10.py` to generate candidate issuer/offering matches.
2. Build the primary-market text universe with `prepare_primary_sentiment_universe_close_time.py`.
3. Create and review labeling samples.
4. Train relevance and sentiment classifiers with `train_bond_classifier_final.py`.
5. Apply the classifiers to the full candidate universe to produce `classified.csv`.

## Example Commands

Run matcher:

```bash
python3 bond_entity_matcher_v10.py \
  --telegram-db data.db \
  --smartlab-db smartlab_bonds.db \
  --panel panel_final_v3.xlsx \
  --issuer-aliases issuer_aliases.xlsx \
  --issue-aliases issue_aliases_candidates_fixed.xlsx
```

Train classifiers:

```bash
python3 train_bond_classifier_final.py \
  --train \
  --training-data training_data_combined.csv \
  --base-model model_dapt_clean_from118k
```

Run inference:

```bash
python3 train_bond_classifier_final.py \
  --predict \
  --input all_matches.csv \
  --output classified.csv
```

## Output

The final output is:

```text
classified.csv
```

This file is consumed by `analysis_pipeline_final/scripts/build_pipeline.py`.
