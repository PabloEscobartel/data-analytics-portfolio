#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Batch runner for predictive models across sentiment aggregation datasets.

Default run:
  python3 run_prediction_batch.py

This runs pre-book datasets only: book_open thresholds 3/5/10. book_close
prediction is blocked by default because it is not a clean ex-ante forecast.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd

import run_prediction as pred


DEFAULT_OUTPUT = "prediction_results_book_open_all_thresholds.xlsx"
DEFAULT_OUTPUT_DIR = Path("prediction_results_batch")
DEFAULT_DATASET_TEMPLATE = "regression_dataset_{window}_{min_messages}/regression_ready.csv"
WINDOWS = ["book_open", "book_close"]
DEFAULT_WINDOWS = ["book_open"]
MIN_MESSAGES = [3, 5, 10]


def find_dataset(window: str, min_messages: int, dataset_template: str = DEFAULT_DATASET_TEMPLATE) -> Optional[Path]:
    candidates = [Path(dataset_template.format(window=window, min_messages=min_messages))]
    if dataset_template == DEFAULT_DATASET_TEMPLATE and min_messages == 3:
        candidates.append(Path(f"regression_dataset_{window}/regression_ready.csv"))
    for path in candidates:
        if path.exists():
            return path
    return None


def short_window(window: str) -> str:
    return "open" if window == "book_open" else "close"


def scenario_name(window: str, min_messages: int) -> str:
    return f"{short_window(window)}_{min_messages}"


def add_meta(row: Dict[str, Any], window: str, min_messages: int, dataset: Path) -> Dict[str, Any]:
    return {
        "window": window,
        "min_messages": min_messages,
        "scenario": scenario_name(window, min_messages),
        "dataset": str(dataset),
        **row,
    }


def flatten_model_comparison(
    all_results: Dict[str, Any],
    window: str,
    min_messages: int,
    dataset: Path,
) -> pd.DataFrame:
    rows = []
    for target_key, target_data in all_results.items():
        for res in target_data.get("results", []):
            rows.append(add_meta({
                "target": target_key,
                "target_label": pred.TARGETS[target_key]["label"],
                "model": res["model"],
                "threshold": res["threshold"],
                "n_train": res["n_train"],
                "n_test": res["n_test"],
                "test_pos_rate": res["test_pos_rate"],
                "train_auc": res["train_auc"],
                "test_auc": res["test_auc"],
                "auc_ci_low": res["test_auc_ci_low"],
                "auc_ci_high": res["test_auc_ci_high"],
                "pr_auc": res["test_pr_auc"],
                "balanced_accuracy": res["test_balanced_accuracy"],
                "test_accuracy": res["test_accuracy"],
                "test_precision": res["test_precision"],
                "test_recall": res["test_recall"],
                "test_f1": res["test_f1"],
                "brier_score": res["test_brier"],
                "tuned_params": pred.format_tuned_params(res.get("tuned_params")),
            }, window, min_messages, dataset))
    return pd.DataFrame(rows)


def flatten_sentiment_value(
    all_results: Dict[str, Any],
    window: str,
    min_messages: int,
    dataset: Path,
) -> pd.DataFrame:
    rows = []
    for target_key, target_data in all_results.items():
        results = target_data.get("results", [])
        for model_type in pred.model_types_with_optional_xgb(["Logit", "RF", "GBM"]):
            base = next((r for r in results if r["model"] == f"{model_type} (no sentiment)"), None)
            buzz = next((r for r in results if r["model"] == f"{model_type} (+ buzz)"), None)
            sent = next((r for r in results if r["model"] == f"{model_type} (+ sentiment)"), None)
            comparisons = [
                ("Controls -> buzz", base, buzz, buzz, "delta_auc_vs_controls"),
                ("Buzz -> buzz + SI", buzz, sent, sent, "delta_auc_vs_buzz"),
                ("Controls -> buzz + SI", base, sent, sent, "delta_auc_ci"),
            ]
            for comparison, before, after, ci_source, key_prefix in comparisons:
                if before is None or after is None:
                    continue
                rows.append(add_meta({
                    "target": target_key,
                    "target_label": pred.TARGETS[target_key]["label"],
                    "model_type": model_type,
                    "comparison": comparison,
                    "auc_before": before["test_auc"],
                    "auc_after": after["test_auc"],
                    "delta_auc": after["test_auc"] - before["test_auc"],
                    "delta_auc_ci_low": ci_source.get(f"{key_prefix}_low", np.nan),
                    "delta_auc_ci_high": ci_source.get(f"{key_prefix}_high", np.nan),
                }, window, min_messages, dataset))
    return pd.DataFrame(rows)


def flatten_temporal_cv(
    all_results: Dict[str, Any],
    window: str,
    min_messages: int,
    dataset: Path,
) -> pd.DataFrame:
    rows = []
    for target_key, target_data in all_results.items():
        for cv_type, cv_results in [
            ("Expanding", target_data.get("cv_results", {})),
            (f"Rolling {pred.ROLLING_WINDOW_YEARS}y", target_data.get("rolling_cv_results", {})),
        ]:
            for model_name, cv_df in cv_results.items():
                if cv_df.empty:
                    continue
                mean_auc = cv_df["auc"].mean()
                for _, rec in cv_df.iterrows():
                    rows.append(add_meta({
                        "target": target_key,
                        "target_label": pred.TARGETS[target_key]["label"],
                        "cv_type": cv_type,
                        "model": model_name,
                        "test_year": int(rec["test_year"]),
                        "train_years": f"{int(rec['train_start'])}-{int(rec['train_end'])}",
                        "n_train": int(rec["n_train"]),
                        "n_test": int(rec["n_test"]),
                        "fold_auc": rec["auc"],
                        "mean_auc": mean_auc,
                    }, window, min_messages, dataset))
    return pd.DataFrame(rows)


def flatten_temporal_sentiment_value(
    all_results: Dict[str, Any],
    window: str,
    min_messages: int,
    dataset: Path,
) -> pd.DataFrame:
    rows = []
    for target_key, target_data in all_results.items():
        for cv_type, cv_results in [
            ("Expanding", target_data.get("cv_results", {})),
            (f"Rolling {pred.ROLLING_WINDOW_YEARS}y", target_data.get("rolling_cv_results", {})),
        ]:
            for model_type in pred.model_types_with_optional_xgb(["Logit", "RF", "GBM"]):
                comparisons = [
                    ("Controls -> buzz", (f"{model_type} (no sentiment)", f"{model_type} (+ buzz)")),
                    ("Buzz -> buzz + SI", (f"{model_type} (+ buzz)", f"{model_type} (+ sentiment)")),
                    ("Controls -> buzz + SI", (f"{model_type} (no sentiment)", f"{model_type} (+ sentiment)")),
                ]
                for comparison, pair in comparisons:
                    fold_rows = pred.temporal_delta_rows(cv_results, model_type, pair, "auc")
                    mean_delta, ci_low, ci_high, n_folds = pred.temporal_metric_delta_summary(fold_rows)
                    for rec in fold_rows:
                        rows.append(add_meta({
                            "target": target_key,
                            "target_label": pred.TARGETS[target_key]["label"],
                            "cv_type": cv_type,
                            "model_type": model_type,
                            "comparison": comparison,
                            "test_year": rec["test_year"],
                            "train_years": rec["train_years"],
                            "n_train": rec["n_train"],
                            "n_test": rec["n_test"],
                            "auc_before": rec["before"],
                            "auc_after": rec["after"],
                            "delta_auc": rec["delta"],
                            "mean_delta_auc": mean_delta,
                            "fold_ci_low": ci_low,
                            "fold_ci_high": ci_high,
                            "n_folds": n_folds,
                        }, window, min_messages, dataset))
    return pd.DataFrame(rows)


def flatten_continuous_targets(
    continuous_results: Dict[str, Any],
    window: str,
    min_messages: int,
    dataset: Path,
) -> pd.DataFrame:
    rows = []
    for target_key, target_data in continuous_results.items():
        for sample_key, sample_data in target_data.get("samples", {}).items():
            for res in sample_data.get("results", []):
                rows.append(add_meta({
                    "target": target_key,
                    "target_label": target_data["label"],
                    "sample": sample_key,
                    "sample_label": res["sample_label"],
                    "model": res["model"],
                    "feature_set": res["feature_set"],
                    "n_train": res["n_train"],
                    "n_test": res["n_test"],
                    "train_r2": res["train_r2"],
                    "test_r2": res["test_r2"],
                    "test_rmse": res["test_rmse"],
                    "test_mae": res["test_mae"],
                    "tuned_params": pred.format_tuned_params(res.get("tuned_params")),
                }, window, min_messages, dataset))
    return pd.DataFrame(rows)


def flatten_continuous_sentiment_value(
    continuous_results: Dict[str, Any],
    window: str,
    min_messages: int,
    dataset: Path,
) -> pd.DataFrame:
    rows = []
    for target_key, target_data in continuous_results.items():
        for sample_key, sample_data in target_data.get("samples", {}).items():
            results = sample_data.get("results", [])
            for model_type in pred.model_types_with_optional_xgb(["Ridge", "RF", "GBM"]):
                controls = next((r for r in results if r["model"] == f"{model_type} (controls)"), None)
                buzz = next((r for r in results if r["model"] == f"{model_type} (buzz)"), None)
                sent = next((r for r in results if r["model"] == f"{model_type} (sentiment)"), None)
                terms = next((r for r in results if r["model"] == f"{model_type} (terms)"), None)
                terms_buzz = next((r for r in results if r["model"] == f"{model_type} (terms+buzz)"), None)
                terms_sent = next((r for r in results if r["model"] == f"{model_type} (terms+sentiment)"), None)
                for comparison, before, after in [
                    ("Controls -> buzz", controls, buzz),
                    ("Buzz -> buzz + SI", buzz, sent),
                    ("Controls -> buzz + SI", controls, sent),
                    ("Terms -> terms + buzz", terms, terms_buzz),
                    ("Terms + buzz -> terms + buzz + SI", terms_buzz, terms_sent),
                    ("Terms -> terms + buzz + SI", terms, terms_sent),
                ]:
                    if before is None or after is None:
                        continue
                    rows.append(add_meta({
                        "target": target_key,
                        "target_label": target_data["label"],
                        "sample": sample_key,
                        "sample_label": sample_data["label"],
                        "model_type": model_type,
                        "comparison": comparison,
                        "r2_before": before["test_r2"],
                        "r2_after": after["test_r2"],
                        "delta_r2": after["test_r2"] - before["test_r2"],
                        "rmse_before": before["test_rmse"],
                        "rmse_after": after["test_rmse"],
                        "rmse_improvement": before["test_rmse"] - after["test_rmse"],
                    }, window, min_messages, dataset))
    return pd.DataFrame(rows)


def flatten_continuous_cv(
    continuous_results: Dict[str, Any],
    window: str,
    min_messages: int,
    dataset: Path,
) -> pd.DataFrame:
    rows = []
    for target_key, target_data in continuous_results.items():
        for sample_key, sample_data in target_data.get("samples", {}).items():
            for cv_type, cv_results in [
                ("Expanding", sample_data.get("cv_results", {})),
                (f"Rolling {pred.ROLLING_WINDOW_YEARS}y", sample_data.get("rolling_cv_results", {})),
            ]:
                for model_name, cv_df in cv_results.items():
                    if cv_df.empty:
                        continue
                    mean_r2 = cv_df["r2"].mean()
                    mean_rmse = cv_df["rmse"].mean()
                    mean_mae = cv_df["mae"].mean()
                    for _, rec in cv_df.iterrows():
                        rows.append(add_meta({
                            "target": target_key,
                            "target_label": target_data["label"],
                            "sample": sample_key,
                            "sample_label": sample_data["label"],
                            "cv_type": cv_type,
                            "model": model_name,
                            "test_year": int(rec["test_year"]),
                            "train_years": f"{int(rec['train_start'])}-{int(rec['train_end'])}",
                            "n_train": int(rec["n_train"]),
                            "n_test": int(rec["n_test"]),
                            "fold_r2": rec["r2"],
                            "fold_rmse": rec["rmse"],
                            "fold_mae": rec["mae"],
                            "mean_r2": mean_r2,
                            "mean_rmse": mean_rmse,
                            "mean_mae": mean_mae,
                        }, window, min_messages, dataset))
    return pd.DataFrame(rows)


def flatten_continuous_temporal_sentiment_value(
    continuous_results: Dict[str, Any],
    window: str,
    min_messages: int,
    dataset: Path,
) -> pd.DataFrame:
    rows = []
    for target_key, target_data in continuous_results.items():
        for sample_key, sample_data in target_data.get("samples", {}).items():
            for cv_type, cv_results in [
                ("Expanding", sample_data.get("cv_results", {})),
                (f"Rolling {pred.ROLLING_WINDOW_YEARS}y", sample_data.get("rolling_cv_results", {})),
            ]:
                for model_type in pred.model_types_with_optional_xgb(["Ridge", "RF", "GBM"]):
                    comparisons = [
                        ("Controls -> buzz", (f"{model_type} (controls)", f"{model_type} (buzz)")),
                        ("Buzz -> buzz + SI", (f"{model_type} (buzz)", f"{model_type} (sentiment)")),
                        ("Controls -> buzz + SI", (f"{model_type} (controls)", f"{model_type} (sentiment)")),
                        ("Terms -> terms + buzz", (f"{model_type} (terms)", f"{model_type} (terms+buzz)")),
                        (
                            "Terms + buzz -> terms + buzz + SI",
                            (f"{model_type} (terms+buzz)", f"{model_type} (terms+sentiment)"),
                        ),
                        ("Terms -> terms + buzz + SI", (f"{model_type} (terms)", f"{model_type} (terms+sentiment)")),
                    ]
                    for comparison, pair in comparisons:
                        r2_rows = pred.temporal_delta_rows(cv_results, model_type, pair, "r2")
                        rmse_rows = pred.temporal_delta_rows(cv_results, model_type, pair, "rmse")
                        rmse_by_year = {r["test_year"]: r for r in rmse_rows}
                        mean_delta_r2, _, _, n_folds = pred.temporal_metric_delta_summary(r2_rows)
                        rmse_improvements = [
                            rmse_by_year[r["test_year"]]["before"] - rmse_by_year[r["test_year"]]["after"]
                            for r in r2_rows
                            if r["test_year"] in rmse_by_year
                        ]
                        mean_rmse_improvement = (
                            float(np.mean(rmse_improvements)) if rmse_improvements else np.nan
                        )
                        for rec in r2_rows:
                            rmse_rec = rmse_by_year.get(rec["test_year"])
                            if rmse_rec is None:
                                continue
                            rows.append(add_meta({
                                "target": target_key,
                                "target_label": target_data["label"],
                                "sample": sample_key,
                                "sample_label": sample_data["label"],
                                "cv_type": cv_type,
                                "model_type": model_type,
                                "comparison": comparison,
                                "test_year": rec["test_year"],
                                "train_years": rec["train_years"],
                                "n_train": rec["n_train"],
                                "n_test": rec["n_test"],
                                "r2_before": rec["before"],
                                "r2_after": rec["after"],
                                "delta_r2": rec["delta"],
                                "rmse_before": rmse_rec["before"],
                                "rmse_after": rmse_rec["after"],
                                "rmse_improvement": rmse_rec["before"] - rmse_rec["after"],
                                "mean_delta_r2": mean_delta_r2,
                                "mean_rmse_improvement": mean_rmse_improvement,
                                "n_folds": n_folds,
                            }, window, min_messages, dataset))
    return pd.DataFrame(rows)


def write_combined_workbook(frames: Dict[str, List[pd.DataFrame]], output_path: Path) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, parts in frames.items():
            df = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])


def run_one(window: str, min_messages: int, dataset: Path, output_dir: Path) -> Dict[str, pd.DataFrame]:
    scen = scenario_name(window, min_messages)
    output_file = output_dir / f"prediction_results_{scen}.xlsx"
    print(f"\n=== Prediction scenario: {scen} ===")
    print(f"Input: {dataset}")
    print(f"Output: {output_file}")

    data = pred.load_and_prepare(str(dataset))
    all_results = pred.run_prediction_analysis(data)
    continuous_results = pred.run_continuous_prediction_analysis(data)
    pred.build_excel(all_results, str(output_file), continuous_results)

    return {
        "Model Comparison": flatten_model_comparison(all_results, window, min_messages, dataset),
        "Sentiment Value": flatten_sentiment_value(all_results, window, min_messages, dataset),
        "Temporal CV": flatten_temporal_cv(all_results, window, min_messages, dataset),
        "Temporal Sent Value": flatten_temporal_sentiment_value(all_results, window, min_messages, dataset),
        "Continuous Targets": flatten_continuous_targets(continuous_results, window, min_messages, dataset),
        "Continuous Sent Value": flatten_continuous_sentiment_value(continuous_results, window, min_messages, dataset),
        "Continuous CV": flatten_continuous_cv(continuous_results, window, min_messages, dataset),
        "Continuous Temporal Sent": flatten_continuous_temporal_sentiment_value(
            continuous_results, window, min_messages, dataset
        ),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Batch predictive models across regression_ready datasets.")
    parser.add_argument("output", nargs="?", default=DEFAULT_OUTPUT)
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR))
    parser.add_argument("--windows", nargs="+", choices=WINDOWS, default=DEFAULT_WINDOWS)
    parser.add_argument("--min-messages", nargs="+", type=int, default=MIN_MESSAGES)
    parser.add_argument(
        "--dataset-template",
        default=DEFAULT_DATASET_TEMPLATE,
        help=(
            "Template for regression_ready.csv paths. Available fields: {window}, {min_messages}. "
            "Default: regression_dataset_{window}_{min_messages}/regression_ready.csv"
        ),
    )
    parser.add_argument(
        "--allow-book-close",
        action="store_true",
        help="Allow book_close prediction as a robustness exercise. Main prediction should use book_open only.",
    )
    args = parser.parse_args()

    if "book_close" in args.windows and not args.allow_book_close:
        parser.error(
            "book_close is blocked for prediction because it is not a clean pre-book forecast. "
            "Use --allow-book-close only for a clearly labelled robustness exercise."
        )

    output_path = Path(args.output)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    frames: Dict[str, List[pd.DataFrame]] = {
        "Model Comparison": [],
        "Sentiment Value": [],
        "Temporal CV": [],
        "Temporal Sent Value": [],
        "Continuous Targets": [],
        "Continuous Sent Value": [],
        "Continuous CV": [],
        "Continuous Temporal Sent": [],
    }
    missing_rows = []

    for window in args.windows:
        for min_messages in args.min_messages:
            dataset = find_dataset(window, min_messages, dataset_template=args.dataset_template)
            if dataset is None:
                missing_rows.append({
                    "window": window,
                    "min_messages": min_messages,
                    "status": "missing regression_ready.csv",
                })
                continue
            result_frames = run_one(window, min_messages, dataset, output_dir)
            for sheet_name, df in result_frames.items():
                frames[sheet_name].append(df)

    if not any(frames.values()):
        raise FileNotFoundError("No prediction datasets found.")

    if missing_rows:
        frames["Missing Inputs"] = [pd.DataFrame(missing_rows)]

    write_combined_workbook(frames, output_path)
    print(f"\nCombined prediction summary saved: {output_path}")


if __name__ == "__main__":
    main()
