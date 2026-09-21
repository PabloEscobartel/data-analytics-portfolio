# Results Workbook Guide: 7-Day Lookback

This guide describes the output workbooks produced by the 7-day sentiment-window robustness pipeline.

Generated workbooks are not tracked in Git. They can be recreated by running:

```bash
bash analysis_pipeline_final/runbooks/run_full_pipeline_7d.sh
bash analysis_pipeline_final/runbooks/collect_final_results.sh 7d
```

## Main Result Files

### `regression_results_upsize_lookback7_all_windows.xlsx`

Full-sample regression results for 7-day book-open and book-close sentiment windows.

### `regression_results_upsize_lookback7_all_windows_fixed.xlsx`

Fixed-coupon subsample regression results for the 7-day window.

### `regression_results_upsize_lookback7_all_windows_floater.xlsx`

Floating-rate subsample regression results for the 7-day window.

### `diagnostics_results_upsize_lookback7_all_windows.xlsx`

Full-sample diagnostics for the 7-day specifications.

### `diagnostics_results_upsize_lookback7_all_windows_fixed.xlsx`

Diagnostics for fixed-coupon bonds.

### `diagnostics_results_upsize_lookback7_all_windows_floater.xlsx`

Diagnostics for floating-rate bonds.

### `descriptive_statistics_lookback7_all_windows.xlsx`

Descriptive statistics for variables used in the 7-day specifications.

### `distribution_plots_lookback7_all_windows.pdf`

Distribution plots for the 7-day datasets.

### `prediction_results_book_open_lookback7_all_thresholds.xlsx`

Predictive-model results for 7-day book-open datasets.

### `gjrm_biprobit_bootstrap_se_combined_lookback7_all_samples.xlsx`

Combined GJRM/bivariate-probit results with issuer-cluster bootstrap standard errors.

### `regression_results_high_spread_lookback7_open_5.xlsx`

High-spread robustness checks for the 7-day book-open, minimum-five-message specification.

## Interpretation Notes

The 7-day window is a timing robustness check. It uses the same model structure as the 14-day baseline but narrows the sentiment aggregation period.
