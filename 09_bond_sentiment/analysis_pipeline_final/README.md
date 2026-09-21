# Final Analysis Pipeline

This folder contains the reproducible analysis pipeline used after text classification is complete.

The pipeline transforms:

```text
panel_final_v3.xlsx + classified.csv
```

into regression datasets, diagnostics, prediction results, and econometric output workbooks.

## Required Inputs

Place the following files in `analysis_pipeline_final/data_inputs/` or in the project root:

```text
panel_final_v3.xlsx
classified.csv
```

These files are excluded from Git because they are large and contain project data.

## Main Scripts

```text
scripts/build_pipeline.py
scripts/run_regressions_upsize.py
scripts/run_regressions_upsize_batch.py
scripts/run_diagnostics_upsize.py
scripts/run_diagnostics_upsize_batch.py
scripts/run_descriptive_statistics.py
scripts/run_prediction.py
scripts/run_prediction_batch.py
scripts/run_regressions_high_spread.py
scripts/run_biprobit_gjrm_cluster_bootstrap_v4.R
scripts/combine_gjrm_bootstrap_se.py
```

## Runbooks

Run the baseline 14-day sentiment window:

```bash
bash analysis_pipeline_final/runbooks/run_full_pipeline_14d.sh
```

Run the 7-day robustness window:

```bash
bash analysis_pipeline_final/runbooks/run_full_pipeline_7d.sh
```

Collect generated files into final result folders:

```bash
bash analysis_pipeline_final/runbooks/collect_final_results.sh 14d
bash analysis_pipeline_final/runbooks/collect_final_results.sh 7d
```

## Methodological Defaults

- Sentiment is aggregated before book opening or book closing.
- Baseline lookback window: 14 days.
- Robustness lookback window: 7 days.
- Minimum relevant-message thresholds: 3, 5, and 10.
- Issuer-level messages are assigned to same-issuer offerings inside each offering-specific time window.
- The final unit of observation is the bond offering.
- The baseline issuer-history control is `hist_reduction_share`.
- Industry fixed effects are excluded from the final specification.
- Prediction models use `book_open` datasets only.

## Main Outputs

The runbooks generate:

- `regression_results_upsize_*.xlsx`
- `diagnostics_results_upsize_*.xlsx`
- `descriptive_statistics_*.xlsx`
- `distribution_plots_*.pdf`
- `prediction_results_*.xlsx`
- `gjrm_biprobit_bootstrap_*`
- `regression_results_high_spread_*.xlsx`

Generated outputs are ignored by Git and can be recreated.
