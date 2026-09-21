#!/usr/bin/env bash
set -euo pipefail

PIPELINE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PROJECT_ROOT="$(cd "${PIPELINE_DIR}/.." && pwd)"
cd "${PROJECT_ROOT}"

MODE="${1:-14d}"

copy_if_exists() {
  local src="$1"
  local dst="$2"
  if [ -e "${src}" ]; then
    cp "${src}" "${dst}/"
  else
    echo "skip missing: ${src}"
  fi
}

copy_as_if_exists() {
  local src="$1"
  local dst="$2"
  local name="$3"
  if [ -e "${src}" ]; then
    cp "${src}" "${dst}/${name}"
  else
    echo "skip missing: ${src}"
  fi
}

if [ "${MODE}" = "7d" ]; then
  OUT="final_results_lookback7"
  mkdir -p "${OUT}"
  copy_if_exists descriptive_statistics_lookback7_all_windows.xlsx "${OUT}"
  copy_if_exists distribution_plots_lookback7_all_windows.pdf "${OUT}"
  copy_if_exists regression_results_upsize_lookback7_all_windows.xlsx "${OUT}"
  copy_if_exists regression_results_upsize_lookback7_all_windows_fixed.xlsx "${OUT}"
  copy_if_exists regression_results_upsize_lookback7_all_windows_floater.xlsx "${OUT}"
  copy_if_exists diagnostics_results_upsize_lookback7_all_windows.xlsx "${OUT}"
  copy_if_exists diagnostics_results_upsize_lookback7_all_windows_fixed.xlsx "${OUT}"
  copy_if_exists diagnostics_results_upsize_lookback7_all_windows_floater.xlsx "${OUT}"
  copy_if_exists prediction_results_book_open_lookback7_all_thresholds.xlsx "${OUT}"
  copy_if_exists prediction_results_batch_lookback7/prediction_results_open_3.xlsx "${OUT}"
  copy_if_exists prediction_results_batch_lookback7/prediction_results_open_5.xlsx "${OUT}"
  copy_if_exists prediction_results_batch_lookback7/prediction_results_open_10.xlsx "${OUT}"
  copy_as_if_exists gjrm_biprobit_bootstrap_lookback7/gjrm_biprobit_bootstrap_coefficients.xlsx "${OUT}" gjrm_biprobit_bootstrap_coefficients_lookback7_all.xlsx
  copy_as_if_exists gjrm_biprobit_bootstrap_lookback7_fixed/gjrm_biprobit_bootstrap_coefficients.xlsx "${OUT}" gjrm_biprobit_bootstrap_coefficients_lookback7_fixed.xlsx
  copy_as_if_exists gjrm_biprobit_bootstrap_lookback7_floater/gjrm_biprobit_bootstrap_coefficients.xlsx "${OUT}" gjrm_biprobit_bootstrap_coefficients_lookback7_floater.xlsx
  copy_if_exists gjrm_biprobit_bootstrap_se_combined_lookback7_all_samples.xlsx "${OUT}"
  copy_if_exists regression_results_high_spread_lookback7_open_5.xlsx "${OUT}"
  copy_if_exists results_workbooks_roadmap_lookback7.md "${OUT}"
else
  OUT="final_results_v3_current"
  mkdir -p "${OUT}"
  copy_if_exists descriptive_statistics_all_windows.xlsx "${OUT}"
  copy_if_exists distribution_plots_all_windows.pdf "${OUT}"
  copy_if_exists regression_results_upsize_all_windows.xlsx "${OUT}"
  copy_if_exists regression_results_upsize_all_windows_fixed.xlsx "${OUT}"
  copy_if_exists regression_results_upsize_all_windows_floater.xlsx "${OUT}"
  copy_if_exists diagnostics_results_upsize_all_windows.xlsx "${OUT}"
  copy_if_exists diagnostics_results_upsize_all_windows_fixed.xlsx "${OUT}"
  copy_if_exists diagnostics_results_upsize_all_windows_floater.xlsx "${OUT}"
  copy_if_exists prediction_results_book_open_all_thresholds.xlsx "${OUT}"
  copy_if_exists prediction_results_batch/prediction_results_open_3.xlsx "${OUT}"
  copy_if_exists prediction_results_batch/prediction_results_open_5.xlsx "${OUT}"
  copy_if_exists prediction_results_batch/prediction_results_open_10.xlsx "${OUT}"
  copy_as_if_exists gjrm_biprobit_bootstrap_clean/gjrm_biprobit_bootstrap_coefficients.xlsx "${OUT}" gjrm_biprobit_bootstrap_coefficients_all.xlsx
  copy_as_if_exists gjrm_biprobit_bootstrap_clean_fixed/gjrm_biprobit_bootstrap_coefficients.xlsx "${OUT}" gjrm_biprobit_bootstrap_coefficients_fixed.xlsx
  copy_as_if_exists gjrm_biprobit_bootstrap_clean_floater/gjrm_biprobit_bootstrap_coefficients.xlsx "${OUT}" gjrm_biprobit_bootstrap_coefficients_floater.xlsx
  copy_if_exists gjrm_biprobit_bootstrap_se_combined_all_samples.xlsx "${OUT}"
  copy_if_exists regression_results_high_spread_open_5.xlsx "${OUT}"
  copy_if_exists results_workbooks_roadmap.md "${OUT}"
fi

echo "Collected results into ${OUT}"
