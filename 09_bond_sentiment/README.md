# Bond Placement Sentiment Study

This project contains the code used to study whether public-market sentiment before bookbuilding is associated with corporate bond placement outcomes.

The project combines three components:

1. Panel construction and data enrichment for Russian corporate bond offerings.
2. NLP-based classification of issuer-level and issue-level text mentions.
3. Econometric and predictive analysis of placement outcomes.

Large raw data, trained model weights, and generated result workbooks are not intended to be tracked in Git. They should be stored separately and copied into the expected folders when reproducing the pipeline.

## Repository Layout

```text
analysis_pipeline_final/          Final regression, diagnostics, prediction, and GJRM pipeline
train NLP/                        NLP matching, labeling, training, and inference scripts
panel_enrichment_pipeline/        Rating, placement, organizer, and bookbuilding-time enrichment scripts
cbonds_gemini_extraction_pipeline/ Cbonds/Gemini extraction scripts for coupon guidance and volume fields
raw_social_data_pipeline/         Telegram and SmartLab raw message parsers
legacy_or_superseded_scripts/     Historical scripts retained for traceability
PROJECT_STRUCTURE.md              Short map of the project structure
```

## Quick Start

From the portfolio repository root:

```bash
cd 09_bond_sentiment
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r analysis_pipeline_final/requirements_minimal.txt
```

Use Python 3.10+ and install R with `Rscript` available on PATH. Install the R
packages used by the full runbooks:

```r
install.packages(c("GJRM", "openxlsx"))
```

Dependencies are not version-pinned; the original environment is not supplied.
The complete analysis requires the external inputs below and can take substantial
time (the runbooks default to 999 bootstrap replications).

## Getting the Data

There is no public download URL for the original data bundle or trained weights
in this project. Request the research inputs from the repository owner through
[GitHub](https://github.com/PabloEscobartel). Availability and redistribution depend
on source permissions; access is not guaranteed.

For **analysis only**, obtain `panel_final_v3.xlsx` and `classified.csv` from the
same research snapshot and copy both into `analysis_pipeline_final/data_inputs/`.
The panel is a curated offering-level dataset; `classified.csv` is the output of
the matching and NLP workflow, not a file downloadable directly from a source site.
See [input locations](analysis_pipeline_final/data_inputs/README.md).

For **upstream reproduction**, obtain the additional files listed in
[the NLP data bundle](train%20NLP/DATA_BUNDLE.md), or rebuild the inputs:

1. Collect offering metadata from Cbonds and enrich it using the scripts in
   `panel_enrichment_pipeline/` and `cbonds_gemini_extraction_pipeline/`.
   Source access, manual reconciliation and the curated alias spreadsheets are
   required; these scripts do not provide a one-command rebuild of the final panel.
2. Collect Telegram and SmartLab text with `raw_social_data_pipeline/scripts/`.
   Install `telethon requests beautifulsoup4 tqdm cloudscraper`; despite its name,
   `telegram_parser_pyrogram.py` uses Telethon. Set `TELEGRAM_API_ID` and
   `TELEGRAM_API_HASH` to your own application credentials. Run both Telegram
   scripts from the same directory so they share the local session file.
   Review channel lists and date constants before collecting: the checked-in
   Telegram and SmartLab scripts have different default date ranges.
3. Place the resulting `data.db` and `smartlab_bonds.db`, the panel and reviewed
   alias spreadsheets in `train NLP/`. Follow the [NLP workflow](train%20NLP/README.md)
   to match texts, label samples, train classifiers and generate `classified.csv`.
   Gemini labeling requires your own `GEMINI_API_KEY`. NLP dependencies include
   PyTorch, Transformers, pandas, scikit-learn and Google GenAI; consult script imports
   for optional components. Domain adaptation starts from
   `DeepPavlov/rubert-base-cased-conversational`; the research-specific adapted model
   and classifiers must be obtained separately or retrained.
4. Copy the final panel and classified output into the analysis input directory.

A new collection or retraining may differ from the original research snapshot.
The repository alone is sufficient to inspect the code, but not to reproduce the
reported estimates without external data. Keep raw messages, credentials and model
weights outside Git; `.gitignore` protects their expected local locations.

## Main Analysis Pipeline

The final analysis starts from two inputs:

- `panel_final_v3.xlsx`: final bond offering panel.
- `classified.csv`: NLP-classified issuer/offering text mentions.

These files are not tracked in Git. Place them in:

```text
analysis_pipeline_final/data_inputs/
```

Then run:

```bash
bash analysis_pipeline_final/runbooks/run_full_pipeline_14d.sh
```

For the 7-day lookback robustness check:

```bash
bash analysis_pipeline_final/runbooks/run_full_pipeline_7d.sh
```

Both runbooks create regression datasets, descriptive statistics, diagnostics, predictive-model outputs, bivariate-probit outputs, and high-spread robustness checks.

## NLP Pipeline

The NLP workflow is in `train NLP/`. It includes:

- entity matching scripts;
- Gemini/manual labeling utilities;
- classifier training and inference scripts;
- expected locations for trained relevance and sentiment classifiers.

Raw message databases and trained model weights are not tracked in Git because of size and data-governance constraints.

## Data and Model Files

This project uses large and potentially private data files, including SQLite databases, classified text corpora, trained Transformer weights, and intermediate message-level CSV files. These files are excluded by `.gitignore`.

If the project must be fully reproducible from a fresh machine, provide the data bundle separately with the following minimum files:

```text
analysis_pipeline_final/data_inputs/panel_final_v3.xlsx
analysis_pipeline_final/data_inputs/classified.csv
train NLP/data.db
train NLP/smartlab_bonds.db
train NLP/training_data_combined.csv
train NLP/model_dapt_clean_from118k/
train NLP/bond_classifier_models_f1_0753/
```

## Authentication

No bearer tokens or cookies are stored in the code. RusBonds authentication, if needed, is read from environment variables:

```bash
export RUSBONDS_BEARER_TOKEN="..."
export RUSBONDS_COOKIES_JSON='{"auth.strategy":"local"}'
```

Alternatively, create a local `rusbonds_auth.json` using `rusbonds_auth.example.json` as a template. The real auth file is ignored by Git.

## GitHub Notes

Do not push raw data, databases, model weights, or generated workbooks to a normal GitHub repository. Keep the GitHub repository focused on code, documentation, and lightweight configuration. Use a separate private data store, Git LFS, or release artifacts if large files must be shared.
