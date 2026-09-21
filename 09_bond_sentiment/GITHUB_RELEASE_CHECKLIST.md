# GitHub Release Checklist

Use this checklist before publishing or sharing the repository.

## Include

- Python, R, and shell scripts.
- Markdown documentation.
- Lightweight configuration examples such as `rusbonds_auth.example.json`.
- Minimal dependency files.

## Exclude

- raw SQLite databases;
- raw Telegram/SmartLab exports;
- full classified text corpora;
- trained model weights;
- intermediate message-level datasets;
- generated result workbooks;
- authentication files, cookies, sessions, and API tokens.

## Required External Data Bundle

For full reproduction, provide a separate private data bundle containing:

```text
analysis_pipeline_final/data_inputs/panel_final_v3.xlsx
analysis_pipeline_final/data_inputs/classified.csv
train NLP/data.db
train NLP/smartlab_bonds.db
train NLP/training_data_combined.csv
train NLP/model_dapt_clean_from118k/
train NLP/bond_classifier_models_f1_0753/
```

## Before Pushing

Run:

```bash
git status --short
git check-ignore -v classified.csv data.db smartlab_bonds.db
find . -size +50M -not -path './.git/*'
```

If any large data or model file is staged, remove it from the index before pushing.

## Notes

GitHub blocks regular Git files larger than 100 MiB. For large data or model files, use a private data store, Git LFS, or release artifacts instead of normal Git tracking.
