# Project Delivery Package

This archive contains the final materials for the bond placement sentiment study.

## Main Files

- `report_final.docx` — final written report.
- `Защита_сем_наставника.pptx` — presentation deck.
- `README.md` — repository overview.
- `PROJECT_STRUCTURE.md` — map of folders and workflows.

## Result Folders

- `final_results_v3_current/` — main 14-day sentiment-window results.
- `final_results_lookback7/` — 7-day lookback robustness results.
- `final_results/` and `final_results_v3_open_predictions/` — additional archived result bundles.

## Reproducibility Folders

- `analysis_pipeline_final/` — final regression, diagnostics, prediction, and GJRM pipeline.
- `train NLP/` — NLP matching, labeling, training, and inference workflow.
- `panel_enrichment_pipeline/` — upstream panel enrichment scripts.
- `cbonds_gemini_extraction_pipeline/` — Cbonds/Gemini extraction scripts.
- `raw_social_data_pipeline/` — raw Telegram and SmartLab parsers.
- `legacy_or_superseded_scripts/` — older scripts retained for traceability.

Large data files and trained model weights are included in this delivery archive for convenience. They are not intended to be tracked in Git.
