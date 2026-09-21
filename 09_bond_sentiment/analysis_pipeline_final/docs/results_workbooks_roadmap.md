# Results Workbook Guide

This guide describes the main output workbooks produced by the 14-day sentiment-window pipeline.

Generated workbooks are not tracked in Git. They can be recreated by running:

```bash
bash analysis_pipeline_final/runbooks/run_full_pipeline_14d.sh
bash analysis_pipeline_final/runbooks/collect_final_results.sh 14d
```

## Main Result Files

### `regression_results_upsize_all_windows.xlsx`

Regression results for the full sample across:

- book-open and book-close sentiment windows;
- minimum relevant-message thresholds of 3, 5, and 10.

The workbook includes baseline regressions, robustness specifications, nonlinear terms, interaction checks, and average marginal effects for binary outcomes.

### `regression_results_upsize_all_windows_fixed.xlsx`

Same structure as the full-sample workbook, restricted to fixed-coupon bonds.

### `regression_results_upsize_all_windows_floater.xlsx`

Same structure as the full-sample workbook, restricted to floating-rate bonds.

### `diagnostics_results_upsize_all_windows.xlsx`

Diagnostic tests for the full-sample specifications, including sample sizes, VIF checks, RESET tests, heteroskedasticity diagnostics, residual diagnostics, and influence diagnostics.

### `diagnostics_results_upsize_all_windows_fixed.xlsx`

Diagnostic workbook for fixed-coupon bonds.

### `diagnostics_results_upsize_all_windows_floater.xlsx`

Diagnostic workbook for floating-rate bonds.

### `descriptive_statistics_all_windows.xlsx`

Descriptive statistics for dependent variables, sentiment variables, transformed variables, and core controls.

### `distribution_plots_all_windows.pdf`

Distribution plots for the same variables covered in the descriptive-statistics workbook.

### `prediction_results_book_open_all_thresholds.xlsx`

Predictive-model results for book-open datasets only. The workbook compares models with and without sentiment-related predictors.

### `gjrm_biprobit_bootstrap_se_combined_all_samples.xlsx`

Combined bivariate-probit/GJRM results with issuer-cluster bootstrap standard errors for all sample modes.

### `regression_results_high_spread_open_5.xlsx`

Robustness checks for the high-spread segment using the book-open, minimum-five-message specification.

## Interpretation Notes

- `book_open` uses information available before book opening and is the preferred timing for prediction exercises.
- `book_close` is used as an econometric robustness window.
- `min_messages` controls the minimum number of relevant messages required for sentiment aggregation.
- The main sentiment variables are the Sentiment Index (`si`) and attention measure (`log_buzz`).
- Average marginal effects in logit models are reported in probability-point terms.
