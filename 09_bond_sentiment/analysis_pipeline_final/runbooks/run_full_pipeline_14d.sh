#!/usr/bin/env bash
set -euo pipefail

PIPELINE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PROJECT_ROOT="$(cd "${PIPELINE_DIR}/.." && pwd)"
cd "${PROJECT_ROOT}"

PY="${PYTHON:-${PROJECT_ROOT}/.venv/bin/python}"
if [ ! -x "${PY}" ]; then
  PY="${PYTHON:-python3}"
fi
RS="${RSCRIPT:-Rscript}"
BOOT_B="${BOOT_B:-999}"
PANEL_INPUT="${PANEL:-panel_final_v3.xlsx}"
CLASSIFIED_INPUT="${CLASSIFIED:-classified.csv}"
if [ ! -e "${PANEL_INPUT}" ]; then
  PANEL_INPUT="${PIPELINE_DIR}/data_inputs/panel_final_v3.xlsx"
fi
if [ ! -e "${CLASSIFIED_INPUT}" ]; then
  CLASSIFIED_INPUT="${PIPELINE_DIR}/data_inputs/classified.csv"
fi

echo "1) Build 14-day regression datasets from panel_final_v3.xlsx"
for WINDOW in book_open book_close; do
  for MIN_MSG in 3 5 10; do
    "${PY}" "${PIPELINE_DIR}/scripts/build_pipeline.py" \
      --panel "${PANEL_INPUT}" \
      --classified "${CLASSIFIED_INPUT}" \
      --message-cutoff "${WINDOW}" \
      --lookback-days 14 \
      --min-messages "${MIN_MSG}" \
      --min-messages-field relevant \
      --issuer-level-mode include \
      --output-dir "regression_dataset_${WINDOW}_${MIN_MSG}"
  done
done

echo "2) Regressions: all / fixed / floater"
"${PY}" "${PIPELINE_DIR}/scripts/run_regressions_upsize_batch.py" \
  regression_results_upsize_all_windows.xlsx \
  --all-samples

echo "3) Diagnostics: all / fixed / floater"
"${PY}" "${PIPELINE_DIR}/scripts/run_diagnostics_upsize_batch.py" \
  diagnostics_results_upsize_all_windows.xlsx \
  --all-samples

echo "4) Descriptive statistics and distribution plots"
"${PY}" "${PIPELINE_DIR}/scripts/run_descriptive_statistics.py" \
  --output-xlsx descriptive_statistics_all_windows.xlsx \
  --output-plots distribution_plots_all_windows.pdf

echo "5) Prediction models, book_open only"
"${PY}" "${PIPELINE_DIR}/scripts/run_prediction_batch.py" \
  prediction_results_book_open_all_thresholds.xlsx \
  --output-dir prediction_results_batch \
  --windows book_open \
  --min-messages 3 5 10

echo "6) High-spread / VDO interaction check, book_open min_messages >= 5"
"${PY}" "${PIPELINE_DIR}/scripts/run_regressions_high_spread.py" \
  regression_dataset_book_open_5/regression_ready.csv \
  regression_results_high_spread_open_5.xlsx

echo "7) GJRM bivariate probit with issuer-cluster bootstrap"
"${RS}" "${PIPELINE_DIR}/scripts/run_biprobit_gjrm_cluster_bootstrap_v4.R" \
  --batch \
  --bootstrap \
  --boot-b="${BOOT_B}" \
  --all-samples \
  gjrm_biprobit_bootstrap_clean

echo "8) Combine GJRM bootstrap SE files"
"${PY}" "${PIPELINE_DIR}/scripts/combine_gjrm_bootstrap_se.py" \
  --all-samples \
  gjrm_biprobit_bootstrap_clean \
  gjrm_biprobit_bootstrap_se_combined_all_samples.csv \
  --xlsx gjrm_biprobit_bootstrap_se_combined_all_samples.xlsx

echo "Done: 14-day pipeline."
