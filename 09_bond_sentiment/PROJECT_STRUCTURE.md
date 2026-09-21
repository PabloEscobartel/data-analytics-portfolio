# Project Structure

This file gives a high-level map of the repository. It is intended for readers who want to understand how the code is organized without reading every script.

## Core Folders

### `analysis_pipeline_final/`

Final analysis pipeline. It starts from:

- `panel_final_v3.xlsx`
- `classified.csv`

and produces:

- regression-ready datasets;
- descriptive statistics;
- econometric regressions;
- diagnostics;
- predictive-model outputs;
- bivariate-probit/GJRM outputs;
- high-spread robustness checks.

### `train NLP/`

NLP workflow for text matching, labeling, classifier training, and inference.

This folder contains scripts for:

- entity matching;
- building the text universe for labeling;
- Gemini/manual labeling checks;
- training relevance and sentiment classifiers;
- applying trained classifiers to matched text mentions.

Large databases and model weights are excluded from Git.

### `panel_enrichment_pipeline/`

Upstream panel-construction scripts. These enrich the bond panel with:

- issue and issuer ratings;
- placement metadata;
- organizers;
- bookbuilding times;
- additional fields used as controls.

### `cbonds_gemini_extraction_pipeline/`

Scripts for Cbonds news collection and Gemini-based extraction of:

- coupon/spread guidance;
- placement volume fields;
- bookbuilding-period information;
- quality-control flags for extracted values.

### `raw_social_data_pipeline/`

Parsers for raw Telegram and SmartLab data.

### `legacy_or_superseded_scripts/`

Historical scripts retained for traceability. They are not used in the current final pipeline unless explicitly stated.

## Result Folders

Generated result folders such as `final_results_v3_current/`, `final_results_lookback7/`, `regression_dataset*/`, and `prediction_results*/` are excluded from Git. They can be regenerated from the runbooks in `analysis_pipeline_final/runbooks/`.

## Data Policy

The repository should contain code and documentation. Raw messages, model weights, SQLite databases, large CSV/XLSX files, and authentication files should be distributed separately.
