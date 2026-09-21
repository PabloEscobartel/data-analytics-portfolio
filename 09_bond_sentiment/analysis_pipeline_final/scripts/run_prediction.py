#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Прогнозные модели для успешности первичных размещений облигаций.
Сравнивает Logit, Random Forest, Gradient Boosting и XGBoost, если пакет установлен.
Оценивает добавленную прогнозную ценность сентимента.

Требования: pip install pandas numpy scikit-learn statsmodels openpyxl
Опционально: pip install xgboost
Запуск: python run_prediction.py regression_ready.csv
"""

import sys
import os
import warnings
from typing import List, Dict, Any, Optional

import numpy as np
import pandas as pd
from sklearn.base import clone
from sklearn.model_selection import TimeSeriesSplit, cross_val_predict
from sklearn.linear_model import LogisticRegression, Ridge
from sklearn.ensemble import (
    RandomForestClassifier, GradientBoostingClassifier,
    RandomForestRegressor, GradientBoostingRegressor,
)
from sklearn.metrics import (
    roc_auc_score, accuracy_score, precision_score, recall_score, f1_score,
    confusion_matrix, classification_report, brier_score_loss,
    average_precision_score, balanced_accuracy_score,
    mean_absolute_error, mean_squared_error, r2_score,
)
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    from xgboost import XGBClassifier, XGBRegressor
    XGBOOST_AVAILABLE = True
    XGBOOST_IMPORT_ERROR = ""
except Exception as exc:
    XGBClassifier = None
    XGBRegressor = None
    XGBOOST_AVAILABLE = False
    XGBOOST_IMPORT_ERROR = str(exc).splitlines()[0]

warnings.filterwarnings("ignore")

# ============================================================
# SETTINGS
# ============================================================

DEFAULT_INPUT = "regression_dataset_book_open_5/regression_ready.csv"
OUTPUT_FILE = "prediction_results.xlsx"
YEAR_FROM = 2018
TRAIN_START_YEAR = YEAR_FROM
TRAIN_END_YEAR = 2024  # holdout train: all available years before TEST_YEAR
TEST_YEAR = 2025
CV_MIN_TRAIN_YEAR = 2021  # first fold: train 2018-2021, test 2022
ROLLING_WINDOW_YEARS = 3
N_BOOTSTRAP = 2000
THRESHOLD_VALIDATION_YEAR = TRAIN_END_YEAR
THRESHOLD_OPTIMIZE_METRIC = "f1"
TUNING_MIN_TRAIN_OBS = 30
TUNING_MIN_VALIDATION_OBS = 10
POST_2022_YEAR = 2022
RANDOM_STATE = 42

GBM_PARAMS = {
    "n_estimators": 100,
    "max_depth": 2,
    "min_samples_leaf": 20,
    "learning_rate": 0.03,
    "subsample": 0.8,
    "random_state": RANDOM_STATE,
}

XGB_CLASSIFIER_PARAMS = {
    "n_estimators": 120,
    "max_depth": 2,
    "learning_rate": 0.03,
    "subsample": 0.8,
    "colsample_bytree": 0.8,
    "min_child_weight": 10,
    "reg_lambda": 5.0,
    "objective": "binary:logistic",
    "eval_metric": "logloss",
    "random_state": RANDOM_STATE,
    "n_jobs": -1,
}

XGB_REGRESSOR_PARAMS = {
    "n_estimators": 120,
    "max_depth": 2,
    "learning_rate": 0.03,
    "subsample": 0.8,
    "colsample_bytree": 0.8,
    "min_child_weight": 10,
    "reg_lambda": 5.0,
    "objective": "reg:squarederror",
    "random_state": RANDOM_STATE,
    "n_jobs": -1,
}

GBM_TUNING_GRID = [
    {"n_estimators": 80, "max_depth": 2, "learning_rate": 0.03, "min_samples_leaf": 20, "subsample": 0.8},
    {"n_estimators": 120, "max_depth": 2, "learning_rate": 0.03, "min_samples_leaf": 10, "subsample": 0.8},
    {"n_estimators": 100, "max_depth": 3, "learning_rate": 0.03, "min_samples_leaf": 20, "subsample": 0.8},
    {"n_estimators": 80, "max_depth": 2, "learning_rate": 0.05, "min_samples_leaf": 20, "subsample": 0.8},
]

XGB_TUNING_GRID = [
    {"n_estimators": 80, "max_depth": 2, "learning_rate": 0.03, "min_child_weight": 10, "reg_lambda": 5.0},
    {"n_estimators": 120, "max_depth": 2, "learning_rate": 0.03, "min_child_weight": 5, "reg_lambda": 5.0},
    {"n_estimators": 100, "max_depth": 3, "learning_rate": 0.03, "min_child_weight": 10, "reg_lambda": 5.0},
    {"n_estimators": 80, "max_depth": 2, "learning_rate": 0.05, "min_child_weight": 10, "reg_lambda": 3.0},
]


def model_types_with_optional_xgb(base_types):
    return base_types + (["XGB"] if XGBOOST_AVAILABLE else [])

RATING_MAP = {
    "AAA":15,"AA+":14,"AA":13,"AA-":12,"A+":11,"A":10,"A-":9,
    "BBB+":8,"BBB":7,"BBB-":6,"BB+":5,"BB":4,"BB-":3,
    "B+":2,"B":1,"B-":0,"C":-1,
}

# Variables
BASE_FEATURES_COMMON = [
    "log_num_organizers", "history_debut", "hist_reduction_share",
    "log_full_issue_number", "hist_avg_volume_ratio",
    "rating_num", "ofz_yield", "rvi",
    "log_dur", "has_put", "is_floater",
]
REGIME_FEATURES = ["is_post_2022"]
BUZZ_FEATURES_COMMON = ["log_buzz"]
SI_FEATURES_COMMON = ["si"]
SENTIMENT_FEATURES_COMMON = BUZZ_FEATURES_COMMON + SI_FEATURES_COMMON
SENTIMENT_INTERACTION_FEATURES = ["si_x_post_2022"]

# Holdout uses all prior years for training and the last year for testing.
BASE_FEATURES = BASE_FEATURES_COMMON
BUZZ_FEATURES = BUZZ_FEATURES_COMMON
SENTIMENT_FEATURES = SENTIMENT_FEATURES_COMMON
BUZZ_ALL_FEATURES = BASE_FEATURES + BUZZ_FEATURES
ALL_FEATURES = BASE_FEATURES + SENTIMENT_FEATURES

# Full-period temporal CV can estimate the post-2022 regime shift and interaction.
CV_BASE_FEATURES = BASE_FEATURES_COMMON + REGIME_FEATURES
CV_BUZZ_FEATURES = BUZZ_FEATURES_COMMON
CV_SENTIMENT_FEATURES = SENTIMENT_FEATURES_COMMON + SENTIMENT_INTERACTION_FEATURES
CV_BUZZ_ALL_FEATURES = CV_BASE_FEATURES + CV_BUZZ_FEATURES
CV_ALL_FEATURES = CV_BASE_FEATURES + CV_SENTIMENT_FEATURES

# Targets
TARGETS = {
    "coupon_reduced": {
        "col": "coupon_reduced_dummy",
        "label": "Coupon tightening (yes/no)",
        "pos_label": "Tightened",
        "neg_label": "Not tightened",
    },
    "upsize": {
        "col": "upsize_dummy",
        "label": "Volume upsize (yes/no)",
        "pos_label": "Upsized",
        "neg_label": "Not upsized",
    },
    "joint_success": {
        "col": "joint_success_upsize",
        "label": "Joint success (tightened + upsized)",
        "pos_label": "Joint success",
        "neg_label": "No joint success",
    },
}

CONTINUOUS_TARGETS = {
    "coupon_rate": {
        "col": "coupon_rate",
        "label": "Final coupon rate, %",
        "sample": "fixed_only",
        "note": "Floaters excluded because coupon_rate is not comparable/usually missing.",
    },
    "placement_volume": {
        "col": "placement_volume",
        "label": "Final placement volume",
        "sample": "all",
        "note": "Raw final placement amount.",
    },
    "log_placement_volume": {
        "col": "log_placement_volume",
        "label": "ln(1 + final placement volume)",
        "sample": "all",
        "note": "Log-transformed final placement amount to reduce skewness.",
    },
}

CONTINUOUS_SAMPLES = {
    "all": {"label": "All bonds", "filter": None, "include_floater": True},
    "fixed": {"label": "Fixed coupon bonds", "filter": 0, "include_floater": False},
    "floater": {"label": "Floaters", "filter": 1, "include_floater": False},
}

VAR_LABELS = {
    "log_num_organizers": "ln(1 + organizers)",
    "history_debut": "Debut issuer",
    "log_full_issue_number": "ln(Issue number)",
    "hist_ever_reduced": "Prior tightening",
    "hist_reduction_share": "Hist. share of prior tightenings",
    "hist_avg_volume_ratio": "Hist. avg volume ratio",
    "rating_num": "Credit rating (ordinal)",
    "ofz_yield": "OFZ matched yield",
    "ofz_yield_within_year": "OFZ yield, within-year deviation",
    "rvi": "RVI",
    "log_dur": "ln(Term to exit)",
    "has_put": "Has put option",
    "is_floater": "Floating rate",
    "is_post_2022": "Post-2022 regime",
    "si": "Sentiment Index (SI)",
    "log_buzz": "ln(Buzz)",
    "si_x_post_2022": "SI x Post-2022",
    "teaser_value": "Teaser coupon/rate value",
    "log_teaser_volume": "ln(1 + teaser volume)",
    "coupon_rate": "Final coupon rate",
    "placement_volume": "Final placement volume",
    "log_placement_volume": "ln(1 + placement volume)",
}


# ============================================================
# DATA LOADING (reuses logic from run_regressions_upsize.py)
# ============================================================

def first_existing_col(df, candidates):
    for col in candidates:
        if col in df.columns:
            return col
    raise KeyError(f"Not found: {', '.join(candidates)}")

def load_and_prepare(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, low_memory=False)
    df["year"] = pd.to_datetime(df["placement_date"], errors="coerce").dt.year
    sub = df[df["year"] >= YEAR_FROM].copy()
    if "teaser_value" not in sub.columns and "tizer_value" in sub.columns:
        sub["teaser_value"] = sub["tizer_value"]
    if "teaser_volume" not in sub.columns and "tizer_volume" in sub.columns:
        sub["teaser_volume"] = sub["tizer_volume"]

    # Sentiment
    si_col = first_existing_col(sub, ["si_relevant", "si"])
    buzz_col = first_existing_col(sub, ["log_buzz_relevant", "log_buzz_all", "log_buzz"])
    sub["si"] = pd.to_numeric(sub[si_col], errors="coerce").fillna(0)
    sub["log_buzz"] = pd.to_numeric(sub[buzz_col], errors="coerce").fillna(0)
    sub["is_post_2022"] = (sub["year"] >= POST_2022_YEAR).astype(int)
    sub["si_x_post_2022"] = sub["si"] * sub["is_post_2022"]

    # Dependent variables
    sub["coupon_reduction_bp"] = pd.to_numeric(sub["coupon_reduction_bp"], errors="coerce")
    sub["coupon_reduced_dummy"] = np.where(
        sub["coupon_reduction_bp"].notna(),
        (sub["coupon_reduction_bp"] > 0).astype(int),
        np.nan,
    )
    sub["placement_vol_book"] = pd.to_numeric(sub["placement_vol_book"], errors="coerce")
    sub["upsize_dummy"] = np.where(
        sub["placement_vol_book"].notna(),
        (sub["placement_vol_book"] > 1.01).astype(int),
        np.nan,
    )
    m_joint = sub["coupon_reduction_bp"].notna() & sub["placement_vol_book"].notna()
    sub["joint_success_upsize"] = np.where(
        m_joint,
        ((sub["coupon_reduction_bp"] > 0) & (sub["placement_vol_book"] > 1.01)).astype(int),
        np.nan,
    )
    if "coupon_rate" in sub.columns:
        sub["coupon_rate"] = pd.to_numeric(sub["coupon_rate"], errors="coerce")
    else:
        sub["coupon_rate"] = np.nan
    if "placement_volume" in sub.columns:
        sub["placement_volume"] = pd.to_numeric(sub["placement_volume"], errors="coerce")
    else:
        sub["placement_volume"] = np.nan
    sub.loc[sub["placement_volume"] <= 0, "placement_volume"] = np.nan
    sub["log_placement_volume"] = np.log1p(sub["placement_volume"])
    if "teaser_volume" in sub.columns:
        sub["teaser_volume"] = pd.to_numeric(sub["teaser_volume"], errors="coerce")
    else:
        sub["teaser_volume"] = np.nan
    sub.loc[sub["teaser_volume"] <= 0, "teaser_volume"] = np.nan
    sub["log_teaser_volume"] = np.log1p(sub["teaser_volume"])
    if "teaser_value" in sub.columns:
        sub["teaser_value"] = pd.to_numeric(sub["teaser_value"], errors="coerce")
    else:
        sub["teaser_value"] = np.nan

    # Controls
    sub["num_organizers"] = pd.to_numeric(sub["num_organizers"], errors="coerce")
    sub["log_num_organizers"] = np.log1p(sub["num_organizers"])
    sub["full_issue_number"] = pd.to_numeric(sub["full_issue_number"], errors="coerce")
    sub.loc[sub["full_issue_number"] <= 0, "full_issue_number"] = np.nan
    sub["log_full_issue_number"] = np.log(sub["full_issue_number"])
    for col in ["hist_reduction_share", "has_put", "hist_avg_volume_ratio"]:
        sub[col] = pd.to_numeric(sub[col], errors="coerce")
    if "hist_ever_reduced" in sub.columns:
        sub["hist_ever_reduced"] = pd.to_numeric(sub["hist_ever_reduced"], errors="coerce")
    else:
        sub["hist_ever_reduced"] = np.nan
    valid_hist = sub["full_issue_number"].notna()
    sub["history_debut"] = np.where(valid_hist, sub["full_issue_number"].eq(1).astype(int), np.nan)

    if "log_term" in sub.columns:
        sub["log_dur"] = pd.to_numeric(sub["log_term"], errors="coerce")
    elif "term" in sub.columns:
        sub["log_dur"] = np.log(pd.to_numeric(sub["term"], errors="coerce").clip(lower=1))

    sub["rating_num"] = sub["final_rating"].astype(str).str.strip().map(RATING_MAP)
    sub["ofz_yield"] = pd.to_numeric(sub["ofz_matched_yield"], errors="coerce")
    sub["ofz_yield_within_year"] = (
        sub["ofz_yield"]
        - sub.groupby("year")["ofz_yield"].transform("mean")
    )
    sub["rvi"] = pd.to_numeric(sub["rvi_close"], errors="coerce")
    sub["is_floater"] = pd.to_numeric(sub.get("is_floater", 0), errors="coerce")

    return sub


# ============================================================
# MODEL EVALUATION
# ============================================================

def get_positive_class_score(model, X):
    """Return model scores/probabilities for the positive class."""
    if hasattr(model, "predict_proba"):
        return model.predict_proba(X)[:, 1]
    if hasattr(model, "decision_function"):
        y_score = model.decision_function(X)
        return 1 / (1 + np.exp(-y_score))
    return model.predict(X).astype(float)


def optimize_threshold(y_true, y_prob, metric=THRESHOLD_OPTIMIZE_METRIC) -> float:
    """Choose a classification threshold on validation data only."""
    y_true = np.asarray(y_true).astype(int)
    y_prob = np.asarray(y_prob)

    if len(y_true) == 0 or np.unique(y_true).size < 2:
        return 0.5

    thresholds = np.unique(np.r_[np.linspace(0.01, 0.99, 99), y_prob])
    best_threshold = 0.5
    best_score = -np.inf

    for threshold in thresholds:
        y_pred = (y_prob >= threshold).astype(int)
        if metric == "balanced_accuracy":
            score = balanced_accuracy_score(y_true, y_pred)
        else:
            score = f1_score(y_true, y_pred, zero_division=0)

        if score > best_score or (score == best_score and abs(threshold - 0.5) < abs(best_threshold - 0.5)):
            best_score = score
            best_threshold = float(threshold)

    return best_threshold


def select_threshold_on_validation(model, X_train, y_train, X_val, y_val) -> float:
    """Fit on pre-validation years and choose threshold on validation year."""
    if len(y_train) == 0 or len(y_val) == 0:
        return 0.5
    if y_train.nunique() < 2 or y_val.nunique() < 2:
        return 0.5

    threshold_model = clone(model)
    threshold_model.fit(X_train, y_train)
    y_prob_val = get_positive_class_score(threshold_model, X_val)
    return optimize_threshold(y_val, y_prob_val)


def evaluate_model(
    model,
    X_train,
    y_train,
    X_test,
    y_test,
    model_name: str,
    threshold: float = 0.5,
) -> Dict[str, Any]:
    """Fit model on train, evaluate on test."""
    model.fit(X_train, y_train)

    y_prob_train = get_positive_class_score(model, X_train)
    y_prob_test = get_positive_class_score(model, X_test)
    y_pred_train = (y_prob_train >= threshold).astype(int)
    y_pred_test = (y_prob_test >= threshold).astype(int)

    # Metrics
    results = {
        "model": model_name,
        "threshold": threshold,
        "n_train": len(y_train),
        "n_test": len(y_test),
        "train_pos_rate": y_train.mean(),
        "test_pos_rate": y_test.mean(),
        # Train
        "train_accuracy": accuracy_score(y_train, y_pred_train),
        "train_auc": roc_auc_score(y_train, y_prob_train) if y_train.nunique() > 1 else np.nan,
        # Test
        "test_accuracy": accuracy_score(y_test, y_pred_test),
        "test_precision": precision_score(y_test, y_pred_test, zero_division=0),
        "test_recall": recall_score(y_test, y_pred_test, zero_division=0),
        "test_f1": f1_score(y_test, y_pred_test, zero_division=0),
        "test_balanced_accuracy": balanced_accuracy_score(y_test, y_pred_test),
        "test_auc": roc_auc_score(y_test, y_prob_test) if y_test.nunique() > 1 else np.nan,
        "test_pr_auc": average_precision_score(y_test, y_prob_test) if y_test.nunique() > 1 else np.nan,
        "test_brier": brier_score_loss(y_test, y_prob_test),
        "y_prob_test": y_prob_test,
    }
    results["test_auc_ci_low"], results["test_auc_ci_high"] = bootstrap_auc_ci(y_test, y_prob_test)

    # Confusion matrix
    cm = confusion_matrix(y_test, y_pred_test)
    results["confusion_matrix"] = cm

    return results


def rmse(y_true, y_pred) -> float:
    return float(np.sqrt(mean_squared_error(y_true, y_pred)))


def regression_feature_sets(include_floater: bool = True, target_key: Optional[str] = None) -> Dict[str, List[str]]:
    base = BASE_FEATURES.copy()
    if not include_floater:
        base = [f for f in base if f != "is_floater"]
    feature_sets = {
        "controls": base,
        "buzz": base + BUZZ_FEATURES,
        "sentiment": base + SENTIMENT_FEATURES,
    }
    if target_key == "coupon_rate":
        terms = base + ["teaser_value"]
    elif target_key in {"placement_volume", "log_placement_volume"}:
        terms = base + ["log_teaser_volume"]
    else:
        terms = []
    if terms:
        feature_sets.update({
            "terms": terms,
            "terms+buzz": terms + BUZZ_FEATURES,
            "terms+sentiment": terms + SENTIMENT_FEATURES,
        })
    return feature_sets


def evaluate_regression_model(
    model,
    X_train,
    y_train,
    X_test,
    y_test,
    model_name: str,
) -> Dict[str, Any]:
    model.fit(X_train, y_train)
    y_pred_train = model.predict(X_train)
    y_pred_test = model.predict(X_test)
    return {
        "model": model_name,
        "n_train": len(y_train),
        "n_test": len(y_test),
        "train_r2": r2_score(y_train, y_pred_train) if len(y_train) > 1 else np.nan,
        "test_r2": r2_score(y_test, y_pred_test) if len(y_test) > 1 else np.nan,
        "test_rmse": rmse(y_test, y_pred_test),
        "test_mae": mean_absolute_error(y_test, y_pred_test),
        "y_pred_test": y_pred_test,
    }


def continuous_temporal_cv(
    data: pd.DataFrame,
    features: List[str],
    target_col: str,
    model,
    scale: bool = False,
    window_years: Optional[int] = None,
    min_train_year: int = CV_MIN_TRAIN_YEAR,
    model_class=None,
    model_params: Optional[Dict[str, Any]] = None,
    temporal_tuning: bool = False,
) -> pd.DataFrame:
    """Temporal CV for continuous targets. Expanding if window_years is None, rolling otherwise."""
    rows = []
    years = sorted(int(y) for y in data["year"].dropna().unique())

    for test_year in years:
        if test_year <= min_train_year:
            continue

        if window_years is None:
            train = data[data["year"] < test_year]
        else:
            train_start = test_year - window_years
            train = data[(data["year"] >= train_start) & (data["year"] < test_year)]

        test = data[data["year"].eq(test_year)]
        if len(train) < 30 or len(test) < 10:
            continue

        X_train = train[features]
        X_test = test[features]
        y_train = train[target_col].astype(float)
        y_test = test[target_col].astype(float)

        if scale:
            scaler = StandardScaler()
            X_train = pd.DataFrame(
                scaler.fit_transform(X_train),
                columns=features,
                index=X_train.index,
            )
            X_test = pd.DataFrame(
                scaler.transform(X_test),
                columns=features,
                index=X_test.index,
            )

        if temporal_tuning and model_class is not None and model_params is not None and is_boosting_model(model_class):
            tuned_params = tune_regressor_params_temporal(model_class, model_params, train, features, target_col)
            fit = model_class(**tuned_params)
        else:
            fit = clone(model)
        fit.fit(X_train, y_train)
        y_pred = fit.predict(X_test)

        rows.append({
            "test_year": test_year,
            "train_start": int(train["year"].min()),
            "train_end": int(train["year"].max()),
            "n_train": len(train),
            "n_test": len(test),
            "r2": r2_score(y_test, y_pred) if len(y_test) > 1 else np.nan,
            "rmse": rmse(y_test, y_pred),
            "mae": mean_absolute_error(y_test, y_pred),
        })

    return pd.DataFrame(rows)


def scaled_logit_model(**model_params):
    """Logit factory with fold-specific scaling to avoid look-ahead leakage."""
    return make_pipeline(
        StandardScaler(),
        LogisticRegression(**model_params),
    )


def maybe_add_xgboost_message():
    if not XGBOOST_AVAILABLE:
        print("XGBoost is not available; skipping XGB models.")
        if XGBOOST_IMPORT_ERROR:
            print(f"  Import issue: {XGBOOST_IMPORT_ERROR}")
        print("  Install/repair with: pip install xgboost  and, on macOS,  brew install libomp")


def is_boosting_model(model_class) -> bool:
    return model_class in {GradientBoostingClassifier, GradientBoostingRegressor, XGBClassifier, XGBRegressor}


def tuning_grid_for_model(model_class):
    if model_class in {GradientBoostingClassifier, GradientBoostingRegressor}:
        return GBM_TUNING_GRID
    if XGBOOST_AVAILABLE and model_class in {XGBClassifier, XGBRegressor}:
        return XGB_TUNING_GRID
    return None


def merge_params(base_params: Dict[str, Any], tuned_params: Dict[str, Any]) -> Dict[str, Any]:
    out = dict(base_params)
    out.update(tuned_params)
    return out


def format_tuned_params(params: Optional[Dict[str, Any]]) -> str:
    if not params:
        return ""
    keys = [
        "n_estimators", "max_depth", "learning_rate", "min_samples_leaf",
        "min_child_weight", "reg_lambda", "subsample", "colsample_bytree",
    ]
    return ", ".join(f"{k}={params[k]}" for k in keys if k in params)


def inner_temporal_validation_splits(
    train_frame: pd.DataFrame,
    target_col: str,
    classification: bool,
) -> List[tuple[pd.DataFrame, pd.DataFrame]]:
    """Create inner expanding validation folds entirely inside the outer train window."""
    years = sorted(int(y) for y in train_frame["year"].dropna().unique())
    splits = []

    for val_year in years[1:]:
        inner_train = train_frame[train_frame["year"] < val_year]
        validation = train_frame[train_frame["year"].eq(val_year)]
        if len(inner_train) < TUNING_MIN_TRAIN_OBS or len(validation) < TUNING_MIN_VALIDATION_OBS:
            continue
        if classification:
            if inner_train[target_col].nunique() < 2 or validation[target_col].nunique() < 2:
                continue
        splits.append((inner_train, validation))

    return splits


def temporal_delta_rows(
    cv_results: Dict[str, pd.DataFrame],
    model_type: str,
    model_names: tuple[str, str],
    metric: str,
) -> List[Dict[str, Any]]:
    """Pair two temporal-CV model paths by test year and compute fold-level deltas."""
    before_name, after_name = model_names
    before = cv_results.get(before_name)
    after = cv_results.get(after_name)
    if before is None or after is None or before.empty or after.empty:
        return []

    cols = ["test_year", "train_start", "train_end", "n_train", "n_test", metric]
    merged = before[cols].merge(
        after[cols],
        on=["test_year", "train_start", "train_end", "n_train", "n_test"],
        suffixes=("_before", "_after"),
    )
    if merged.empty:
        return []

    rows = []
    for _, rec in merged.iterrows():
        before_value = rec[f"{metric}_before"]
        after_value = rec[f"{metric}_after"]
        rows.append({
            "model_type": model_type,
            "test_year": int(rec["test_year"]),
            "train_years": f"{int(rec['train_start'])}-{int(rec['train_end'])}",
            "n_train": int(rec["n_train"]),
            "n_test": int(rec["n_test"]),
            "before": before_value,
            "after": after_value,
            "delta": after_value - before_value,
        })
    return rows


def temporal_metric_delta_summary(rows: List[Dict[str, Any]]) -> tuple[float, float, float, int]:
    """Mean and simple fold-percentile interval for temporal deltas."""
    deltas = np.array([r["delta"] for r in rows if pd.notna(r["delta"])], dtype=float)
    if deltas.size == 0:
        return np.nan, np.nan, np.nan, 0
    return (
        float(np.mean(deltas)),
        float(np.quantile(deltas, 0.025)),
        float(np.quantile(deltas, 0.975)),
        int(deltas.size),
    )


def tune_classifier_params_temporal(
    model_class,
    base_params: Dict[str, Any],
    train_frame: pd.DataFrame,
    features: List[str],
    target_col: str,
) -> Dict[str, Any]:
    """Tune boosting params by inner expanding temporal CV inside the outer train window."""
    grid = tuning_grid_for_model(model_class)
    if not grid:
        return dict(base_params)

    splits = inner_temporal_validation_splits(train_frame, target_col, classification=True)
    if not splits:
        return dict(base_params)

    best_params = dict(base_params)
    best_score = -np.inf
    for candidate in grid:
        params = merge_params(base_params, candidate)
        scores = []
        for inner_train, validation in splits:
            X_inner = inner_train[features]
            y_inner = inner_train[target_col].astype(int)
            X_val = validation[features]
            y_val = validation[target_col].astype(int)

            model = model_class(**params)
            model.fit(X_inner, y_inner)
            y_prob = get_positive_class_score(model, X_val)
            scores.append(roc_auc_score(y_val, y_prob))

        score = float(np.mean(scores)) if scores else -np.inf
        if score > best_score:
            best_score = score
            best_params = params
    return best_params


def tune_regressor_params_temporal(
    model_class,
    base_params: Dict[str, Any],
    train_frame: pd.DataFrame,
    features: List[str],
    target_col: str,
) -> Dict[str, Any]:
    """Tune boosting params by inner expanding temporal CV using RMSE."""
    grid = tuning_grid_for_model(model_class)
    if not grid:
        return dict(base_params)

    splits = inner_temporal_validation_splits(train_frame, target_col, classification=False)
    if not splits:
        return dict(base_params)

    best_params = dict(base_params)
    best_score = np.inf
    for candidate in grid:
        params = merge_params(base_params, candidate)
        scores = []
        for inner_train, validation in splits:
            X_inner = inner_train[features]
            y_inner = inner_train[target_col].astype(float)
            X_val = validation[features]
            y_val = validation[target_col].astype(float)

            model = model_class(**params)
            model.fit(X_inner, y_inner)
            y_pred = model.predict(X_val)
            scores.append(rmse(y_val, y_pred))

        score = float(np.mean(scores)) if scores else np.inf
        if score < best_score:
            best_score = score
            best_params = params
    return best_params


def make_tuned_classifier(model_class, base_params, train_frame, features, target_col):
    params = tune_classifier_params_temporal(model_class, base_params, train_frame, features, target_col)
    return model_class(**params), params


def make_tuned_regressor(model_class, base_params, train_frame, features, target_col):
    params = tune_regressor_params_temporal(model_class, base_params, train_frame, features, target_col)
    return model_class(**params), params


def bootstrap_auc_ci(y_true, y_prob, n_bootstrap=N_BOOTSTRAP, ci=0.95, random_state=RANDOM_STATE):
    """Bootstrap confidence interval for holdout AUC."""
    y_true = np.asarray(y_true).astype(int)
    y_prob = np.asarray(y_prob)
    rng = np.random.default_rng(random_state)
    aucs = []

    for _ in range(n_bootstrap):
        idx = rng.integers(0, len(y_true), len(y_true))
        if np.unique(y_true[idx]).size < 2:
            continue
        aucs.append(roc_auc_score(y_true[idx], y_prob[idx]))

    if not aucs:
        return np.nan, np.nan

    alpha = (1 - ci) / 2
    return (
        float(np.quantile(aucs, alpha)),
        float(np.quantile(aucs, 1 - alpha)),
    )


def bootstrap_delta_auc_ci(
    y_true,
    y_prob_base,
    y_prob_sent,
    n_bootstrap=N_BOOTSTRAP,
    ci=0.95,
    random_state=RANDOM_STATE,
):
    """Bootstrap confidence interval for paired AUC difference."""
    y_true = np.asarray(y_true).astype(int)
    y_prob_base = np.asarray(y_prob_base)
    y_prob_sent = np.asarray(y_prob_sent)
    rng = np.random.default_rng(random_state)
    deltas = []

    for _ in range(n_bootstrap):
        idx = rng.integers(0, len(y_true), len(y_true))
        if np.unique(y_true[idx]).size < 2:
            continue
        auc_base = roc_auc_score(y_true[idx], y_prob_base[idx])
        auc_sent = roc_auc_score(y_true[idx], y_prob_sent[idx])
        deltas.append(auc_sent - auc_base)

    if not deltas:
        return np.nan, np.nan

    alpha = (1 - ci) / 2
    return (
        float(np.quantile(deltas, alpha)),
        float(np.quantile(deltas, 1 - alpha)),
    )


def expanding_window_cv(
    data,
    features,
    target_col,
    model_class,
    model_params,
    min_train_year=CV_MIN_TRAIN_YEAR,
    temporal_tuning=False,
):
    """Expanding window: train on [2018, year], test on year+1."""
    results = []
    years = sorted(int(y) for y in data["year"].dropna().unique())

    for test_year in years:
        if test_year <= min_train_year:
            continue
        train = data[data["year"] < test_year]
        test = data[data["year"] == test_year]

        if train.empty or test.empty:
            continue
        if test[target_col].nunique() < 2 or train[target_col].nunique() < 2:
            continue

        X_train = train[features]
        X_test = test[features]
        y_train = train[target_col].astype(int)
        y_test = test[target_col].astype(int)

        if temporal_tuning and is_boosting_model(model_class):
            tuned_params = tune_classifier_params_temporal(model_class, model_params, train, features, target_col)
            model = model_class(**tuned_params)
        else:
            model = model_class(**model_params)
        model.fit(X_train, y_train)

        if hasattr(model, "predict_proba"):
            y_prob = model.predict_proba(X_test)[:, 1]
        elif hasattr(model, "decision_function"):
            y_score = model.decision_function(X_test)
            y_prob = 1 / (1 + np.exp(-y_score))
        else:
            y_prob = model.predict(X_test).astype(float)

        auc = roc_auc_score(y_test, y_prob)
        results.append({
            "test_year": test_year,
            "train_start": int(train["year"].min()),
            "train_end": int(train["year"].max()),
            "n_train": len(train),
            "n_test": len(test),
            "auc": auc,
        })

    return pd.DataFrame(results)


def rolling_window_cv(
    data,
    features,
    target_col,
    model_class,
    model_params,
    window_years=ROLLING_WINDOW_YEARS,
    min_train_year=CV_MIN_TRAIN_YEAR,
    temporal_tuning=False,
):
    """Rolling window: train on the previous N years, test on the next year."""
    results = []
    years = sorted(int(y) for y in data["year"].dropna().unique())

    for test_year in years:
        if test_year <= min_train_year:
            continue

        train_start = test_year - window_years
        train = data[(data["year"] >= train_start) & (data["year"] < test_year)]
        test = data[data["year"] == test_year]

        if train.empty or test.empty:
            continue
        if test[target_col].nunique() < 2 or train[target_col].nunique() < 2:
            continue

        X_train = train[features]
        X_test = test[features]
        y_train = train[target_col].astype(int)
        y_test = test[target_col].astype(int)

        if temporal_tuning and is_boosting_model(model_class):
            tuned_params = tune_classifier_params_temporal(model_class, model_params, train, features, target_col)
            model = model_class(**tuned_params)
        else:
            model = model_class(**model_params)
        model.fit(X_train, y_train)

        if hasattr(model, "predict_proba"):
            y_prob = model.predict_proba(X_test)[:, 1]
        elif hasattr(model, "decision_function"):
            y_score = model.decision_function(X_test)
            y_prob = 1 / (1 + np.exp(-y_score))
        else:
            y_prob = model.predict(X_test).astype(float)

        auc = roc_auc_score(y_test, y_prob)
        results.append({
            "test_year": test_year,
            "train_start": int(train["year"].min()),
            "train_end": int(train["year"].max()),
            "n_train": len(train),
            "n_test": len(test),
            "auc": auc,
        })

    return pd.DataFrame(results)


def get_feature_importance(model, feature_names: List[str]) -> pd.DataFrame:
    """Extract feature importance from fitted model."""
    if hasattr(model, "feature_importances_"):
        imp = model.feature_importances_
    elif hasattr(model, "coef_"):
        imp = np.abs(model.coef_[0])
    else:
        return pd.DataFrame()

    df = pd.DataFrame({
        "feature": feature_names,
        "label": [VAR_LABELS.get(f, f) for f in feature_names],
        "importance": imp,
    }).sort_values("importance", ascending=False)
    df["importance_pct"] = df["importance"] / df["importance"].sum() * 100
    return df


def run_continuous_prediction_analysis(data: pd.DataFrame) -> Dict[str, Any]:
    """Predict final coupon rate / placement volume and test added value of buzz/SI."""
    all_results = {}

    for target_key, target_info in CONTINUOUS_TARGETS.items():
        target_col = target_info["col"]
        print(f"\n{'='*70}")
        print(f"CONTINUOUS TARGET: {target_info['label']}")
        print(f"{'='*70}")

        sample_results = {}
        for sample_key, sample_info in CONTINUOUS_SAMPLES.items():
            if target_info["sample"] == "fixed_only" and sample_key != "fixed":
                continue

            work = data.copy()
            if sample_info["filter"] is not None:
                work = work[work["is_floater"].eq(sample_info["filter"])].copy()

            feature_sets = regression_feature_sets(
                include_floater=sample_info["include_floater"],
                target_key=target_key,
            )
            required_features = sorted(set(feature_sets[max(feature_sets, key=lambda k: len(feature_sets[k]))]))
            mask = (
                work[target_col].notna()
                & work["year"].notna()
                & work[required_features].notna().all(axis=1)
            )
            work = work[mask].copy()
            if work.empty:
                continue

            train_mask = (work["year"] >= TRAIN_START_YEAR) & (work["year"] <= TRAIN_END_YEAR)
            test_mask = work["year"].eq(TEST_YEAR)
            if train_mask.sum() < 30 or test_mask.sum() < 10:
                print(f"  {sample_info['label']}: skipped, too few train/test observations")
                continue

            y_train = work.loc[train_mask, target_col].astype(float)
            y_test = work.loc[test_mask, target_col].astype(float)

            print(
                f"  {sample_info['label']}: train={len(y_train)}, test={len(y_test)}, "
                f"target mean test={y_test.mean():.3f}"
            )

            target_results = []
            cv_results = {}
            rolling_cv_results = {}
            temporal_train_frame = work.loc[train_mask].copy()
            for feature_label, features in feature_sets.items():
                X_train = work.loc[train_mask, features]
                X_test = work.loc[test_mask, features]

                scaler = StandardScaler()
                X_train_sc = pd.DataFrame(
                    scaler.fit_transform(X_train),
                    columns=features,
                    index=X_train.index,
                )
                X_test_sc = pd.DataFrame(
                    scaler.transform(X_test),
                    columns=features,
                    index=X_test.index,
                )

                gbm_model, gbm_params = make_tuned_regressor(
                    GradientBoostingRegressor, GBM_PARAMS, temporal_train_frame, features, target_col
                )

                models = {
                    f"Ridge ({feature_label})": (
                        Ridge(alpha=1.0, random_state=RANDOM_STATE),
                        X_train_sc,
                        X_test_sc,
                        {},
                    ),
                    f"RF ({feature_label})": (
                        RandomForestRegressor(
                            n_estimators=300, max_depth=5, min_samples_leaf=10,
                            random_state=RANDOM_STATE, n_jobs=-1,
                        ),
                        X_train,
                        X_test,
                        {},
                    ),
                    f"GBM ({feature_label})": (
                        gbm_model,
                        X_train,
                        X_test,
                        gbm_params,
                    ),
                }
                if XGBOOST_AVAILABLE:
                    xgb_model, xgb_params = make_tuned_regressor(
                        XGBRegressor, XGB_REGRESSOR_PARAMS, temporal_train_frame, features, target_col
                    )
                    models[f"XGB ({feature_label})"] = (
                        xgb_model,
                        X_train,
                        X_test,
                        xgb_params,
                    )

                for model_name, (model, X_tr, X_te, tuned_params) in models.items():
                    if tuned_params:
                        compact_params = {k: tuned_params[k] for k in sorted(tuned_params) if k in {
                            "n_estimators", "max_depth", "learning_rate", "min_samples_leaf",
                            "min_child_weight", "reg_lambda",
                        }}
                        print(f"    {model_name}: nested temporal tuned params {compact_params}")
                    res = evaluate_regression_model(model, X_tr, y_train, X_te, y_test, model_name)
                    res.update({
                        "target": target_key,
                        "target_label": target_info["label"],
                        "sample": sample_key,
                        "sample_label": sample_info["label"],
                        "feature_set": feature_label,
                        "tuned_params": tuned_params,
                    })
                    target_results.append(res)
                    print(
                        f"    {model_name:24s} R2={res['test_r2']:.4f}, "
                        f"RMSE={res['test_rmse']:.4f}, MAE={res['test_mae']:.4f}"
                    )

            cv_model_specs = {}
            for feature_label, features in feature_sets.items():
                cv_model_specs[f"Ridge ({feature_label})"] = (
                    Ridge(alpha=1.0, random_state=RANDOM_STATE),
                    features,
                    True,
                    None,
                    None,
                )
                cv_model_specs[f"RF ({feature_label})"] = (
                    RandomForestRegressor(
                        n_estimators=300, max_depth=5, min_samples_leaf=10,
                        random_state=RANDOM_STATE, n_jobs=-1,
                    ),
                    features,
                    False,
                    None,
                    None,
                )
                cv_model_specs[f"GBM ({feature_label})"] = (
                    GradientBoostingRegressor(**GBM_PARAMS),
                    features,
                    False,
                    GradientBoostingRegressor,
                    GBM_PARAMS,
                )
                if XGBOOST_AVAILABLE:
                    cv_model_specs[f"XGB ({feature_label})"] = (
                        XGBRegressor(**XGB_REGRESSOR_PARAMS),
                        features,
                        False,
                        XGBRegressor,
                        XGB_REGRESSOR_PARAMS,
                    )

            print("    Continuous expanding-window CV:")
            for model_name, (model, features, scale, model_class, model_params) in cv_model_specs.items():
                cv_df = continuous_temporal_cv(
                    work,
                    features,
                    target_col,
                    model,
                    scale=scale,
                    min_train_year=CV_MIN_TRAIN_YEAR,
                    model_class=model_class,
                    model_params=model_params,
                    temporal_tuning=True,
                )
                cv_results[model_name] = cv_df
                if not cv_df.empty:
                    print(
                        f"      {model_name:24s} mean R2={cv_df['r2'].mean():.4f}, "
                        f"mean RMSE={cv_df['rmse'].mean():.4f}"
                    )

            print(f"    Continuous rolling-window CV ({ROLLING_WINDOW_YEARS} years):")
            for model_name, (model, features, scale, model_class, model_params) in cv_model_specs.items():
                rolling_df = continuous_temporal_cv(
                    work,
                    features,
                    target_col,
                    model,
                    scale=scale,
                    window_years=ROLLING_WINDOW_YEARS,
                    min_train_year=CV_MIN_TRAIN_YEAR,
                    model_class=model_class,
                    model_params=model_params,
                    temporal_tuning=True,
                )
                rolling_cv_results[model_name] = rolling_df
                if not rolling_df.empty:
                    print(
                        f"      {model_name:24s} mean R2={rolling_df['r2'].mean():.4f}, "
                        f"mean RMSE={rolling_df['rmse'].mean():.4f}"
                    )

            sample_results[sample_key] = {
                "label": sample_info["label"],
                "results": target_results,
                "cv_results": cv_results,
                "rolling_cv_results": rolling_cv_results,
            }

        all_results[target_key] = {
            "label": target_info["label"],
            "note": target_info["note"],
            "samples": sample_results,
        }

    return all_results


# ============================================================
# MAIN ANALYSIS
# ============================================================

def run_prediction_analysis(data: pd.DataFrame) -> Dict[str, Any]:
    """Run full prediction analysis for all targets."""

    all_results = {}

    for target_key, target_info in TARGETS.items():
        target_col = target_info["col"]
        target_label = target_info["label"]
        print(f"\n{'='*70}")
        print(f"TARGET: {target_label}")
        print(f"{'='*70}")

        # Prepare data for this target
        mask = data[target_col].notna() & data["year"].notna() & data[CV_ALL_FEATURES].notna().all(axis=1)
        work = data[mask].copy()
        y = work[target_col].astype(int)

        # Train/test split by year
        train_mask = (work["year"] >= TRAIN_START_YEAR) & (work["year"] <= TRAIN_END_YEAR)
        test_mask = work["year"] == TEST_YEAR

        if test_mask.sum() == 0:
            print(f"  ⚠ No test data for {target_label}")
            continue

        X_train_all = work.loc[train_mask, ALL_FEATURES]
        X_test_all = work.loc[test_mask, ALL_FEATURES]
        X_train_buzz = work.loc[train_mask, BUZZ_ALL_FEATURES]
        X_test_buzz = work.loc[test_mask, BUZZ_ALL_FEATURES]
        X_train_base = work.loc[train_mask, BASE_FEATURES]
        X_test_base = work.loc[test_mask, BASE_FEATURES]
        y_train = y[train_mask]
        y_test = y[test_mask]

        if y_train.nunique() < 2 or y_test.nunique() < 2:
            print(f"  ⚠ Not enough class variation for {target_label}")
            continue

        print(f"  Train {TRAIN_START_YEAR}-{TRAIN_END_YEAR}: {len(y_train)} ({y_train.mean():.1%} positive)")
        print(f"  Test {TEST_YEAR}:        {len(y_test)} ({y_test.mean():.1%} positive)")

        # Threshold selection: train on pre-validation years, validate on 2024, apply to 2025.
        threshold_train_mask = (work["year"] >= TRAIN_START_YEAR) & (work["year"] < THRESHOLD_VALIDATION_YEAR)
        validation_mask = work["year"] == THRESHOLD_VALIDATION_YEAR
        X_threshold_train_all = work.loc[threshold_train_mask, ALL_FEATURES]
        X_validation_all = work.loc[validation_mask, ALL_FEATURES]
        X_threshold_train_buzz = work.loc[threshold_train_mask, BUZZ_ALL_FEATURES]
        X_validation_buzz = work.loc[validation_mask, BUZZ_ALL_FEATURES]
        X_threshold_train_base = work.loc[threshold_train_mask, BASE_FEATURES]
        X_validation_base = work.loc[validation_mask, BASE_FEATURES]
        y_threshold_train = y[threshold_train_mask]
        y_validation = y[validation_mask]

        # Scale features for logit
        scaler_all = StandardScaler()
        scaler_buzz = StandardScaler()
        scaler_base = StandardScaler()
        X_train_all_sc = pd.DataFrame(
            scaler_all.fit_transform(X_train_all),
            columns=ALL_FEATURES, index=X_train_all.index,
        )
        X_test_all_sc = pd.DataFrame(
            scaler_all.transform(X_test_all),
            columns=ALL_FEATURES, index=X_test_all.index,
        )
        X_train_buzz_sc = pd.DataFrame(
            scaler_buzz.fit_transform(X_train_buzz),
            columns=BUZZ_ALL_FEATURES, index=X_train_buzz.index,
        )
        X_test_buzz_sc = pd.DataFrame(
            scaler_buzz.transform(X_test_buzz),
            columns=BUZZ_ALL_FEATURES, index=X_test_buzz.index,
        )
        X_train_base_sc = pd.DataFrame(
            scaler_base.fit_transform(X_train_base),
            columns=BASE_FEATURES, index=X_train_base.index,
        )
        X_test_base_sc = pd.DataFrame(
            scaler_base.transform(X_test_base),
            columns=BASE_FEATURES, index=X_test_base.index,
        )
        threshold_scaler_all = StandardScaler()
        threshold_scaler_buzz = StandardScaler()
        threshold_scaler_base = StandardScaler()
        X_threshold_train_all_sc = pd.DataFrame(
            threshold_scaler_all.fit_transform(X_threshold_train_all),
            columns=ALL_FEATURES, index=X_threshold_train_all.index,
        )
        X_validation_all_sc = pd.DataFrame(
            threshold_scaler_all.transform(X_validation_all),
            columns=ALL_FEATURES, index=X_validation_all.index,
        )
        X_threshold_train_buzz_sc = pd.DataFrame(
            threshold_scaler_buzz.fit_transform(X_threshold_train_buzz),
            columns=BUZZ_ALL_FEATURES, index=X_threshold_train_buzz.index,
        )
        X_validation_buzz_sc = pd.DataFrame(
            threshold_scaler_buzz.transform(X_validation_buzz),
            columns=BUZZ_ALL_FEATURES, index=X_validation_buzz.index,
        )
        X_threshold_train_base_sc = pd.DataFrame(
            threshold_scaler_base.fit_transform(X_threshold_train_base),
            columns=BASE_FEATURES, index=X_threshold_train_base.index,
        )
        X_validation_base_sc = pd.DataFrame(
            threshold_scaler_base.transform(X_validation_base),
            columns=BASE_FEATURES, index=X_validation_base.index,
        )

        temporal_train_frame = work.loc[train_mask].copy()
        gbm_base_model, gbm_base_params = make_tuned_classifier(
            GradientBoostingClassifier, GBM_PARAMS, temporal_train_frame, BASE_FEATURES, target_col
        )
        gbm_buzz_model, gbm_buzz_params = make_tuned_classifier(
            GradientBoostingClassifier, GBM_PARAMS, temporal_train_frame, BUZZ_ALL_FEATURES, target_col
        )
        gbm_all_model, gbm_all_params = make_tuned_classifier(
            GradientBoostingClassifier, GBM_PARAMS, temporal_train_frame, ALL_FEATURES, target_col
        )

        # Models
        models = {
            # Without sentiment
            "Logit (no sentiment)": (
                LogisticRegression(max_iter=1000, random_state=RANDOM_STATE),
                X_train_base_sc, X_test_base_sc,
                X_threshold_train_base_sc, X_validation_base_sc,
                BASE_FEATURES,
            ),
            "RF (no sentiment)": (
                RandomForestClassifier(
                    n_estimators=300, max_depth=5, min_samples_leaf=10,
                    random_state=RANDOM_STATE, n_jobs=-1,
                ),
                X_train_base, X_test_base,
                X_threshold_train_base, X_validation_base,
                BASE_FEATURES,
            ),
            "GBM (no sentiment)": (
                gbm_base_model,
                X_train_base, X_test_base,
                X_threshold_train_base, X_validation_base,
                BASE_FEATURES,
                gbm_base_params,
            ),
            # With market attention only
            "Logit (+ buzz)": (
                LogisticRegression(max_iter=1000, random_state=RANDOM_STATE),
                X_train_buzz_sc, X_test_buzz_sc,
                X_threshold_train_buzz_sc, X_validation_buzz_sc,
                BUZZ_ALL_FEATURES,
            ),
            "RF (+ buzz)": (
                RandomForestClassifier(
                    n_estimators=300, max_depth=5, min_samples_leaf=10,
                    random_state=RANDOM_STATE, n_jobs=-1,
                ),
                X_train_buzz, X_test_buzz,
                X_threshold_train_buzz, X_validation_buzz,
                BUZZ_ALL_FEATURES,
            ),
            "GBM (+ buzz)": (
                gbm_buzz_model,
                X_train_buzz, X_test_buzz,
                X_threshold_train_buzz, X_validation_buzz,
                BUZZ_ALL_FEATURES,
                gbm_buzz_params,
            ),
            # With market attention and tone
            "Logit (+ sentiment)": (
                LogisticRegression(max_iter=1000, random_state=RANDOM_STATE),
                X_train_all_sc, X_test_all_sc,
                X_threshold_train_all_sc, X_validation_all_sc,
                ALL_FEATURES,
            ),
            "RF (+ sentiment)": (
                RandomForestClassifier(
                    n_estimators=300, max_depth=5, min_samples_leaf=10,
                    random_state=RANDOM_STATE, n_jobs=-1,
                ),
                X_train_all, X_test_all,
                X_threshold_train_all, X_validation_all,
                ALL_FEATURES,
            ),
            "GBM (+ sentiment)": (
                gbm_all_model,
                X_train_all, X_test_all,
                X_threshold_train_all, X_validation_all,
                ALL_FEATURES,
                gbm_all_params,
            ),
        }
        for key in [
            "Logit (no sentiment)", "RF (no sentiment)",
            "Logit (+ buzz)", "RF (+ buzz)",
            "Logit (+ sentiment)", "RF (+ sentiment)",
        ]:
            models[key] = (*models[key], {})
        if XGBOOST_AVAILABLE:
            xgb_base_model, xgb_base_params = make_tuned_classifier(
                XGBClassifier, XGB_CLASSIFIER_PARAMS, temporal_train_frame, BASE_FEATURES, target_col
            )
            xgb_buzz_model, xgb_buzz_params = make_tuned_classifier(
                XGBClassifier, XGB_CLASSIFIER_PARAMS, temporal_train_frame, BUZZ_ALL_FEATURES, target_col
            )
            xgb_all_model, xgb_all_params = make_tuned_classifier(
                XGBClassifier, XGB_CLASSIFIER_PARAMS, temporal_train_frame, ALL_FEATURES, target_col
            )
            models.update({
                "XGB (no sentiment)": (
                    xgb_base_model,
                    X_train_base, X_test_base,
                    X_threshold_train_base, X_validation_base,
                    BASE_FEATURES,
                    xgb_base_params,
                ),
                "XGB (+ buzz)": (
                    xgb_buzz_model,
                    X_train_buzz, X_test_buzz,
                    X_threshold_train_buzz, X_validation_buzz,
                    BUZZ_ALL_FEATURES,
                    xgb_buzz_params,
                ),
                "XGB (+ sentiment)": (
                    xgb_all_model,
                    X_train_all, X_test_all,
                    X_threshold_train_all, X_validation_all,
                    ALL_FEATURES,
                    xgb_all_params,
                ),
            })

        target_results = []
        feature_importances = {}
        cv_results = {}
        rolling_cv_results = {}

        for model_name, (model, X_tr, X_te, X_thr, X_val, feat_names, tuned_params) in models.items():
            print(f"\n  --- {model_name} ---")
            if tuned_params:
                compact_params = {k: tuned_params[k] for k in sorted(tuned_params) if k in {
                    "n_estimators", "max_depth", "learning_rate", "min_samples_leaf",
                    "min_child_weight", "reg_lambda",
                }}
                print(f"    Nested temporal tuned params: {compact_params}")
            threshold = select_threshold_on_validation(model, X_thr, y_threshold_train, X_val, y_validation)
            res = evaluate_model(model, X_tr, y_train, X_te, y_test, model_name, threshold=threshold)
            res["target"] = target_key
            res["target_label"] = target_label
            res["tuned_params"] = tuned_params
            target_results.append(res)

            print(f"    Threshold ({THRESHOLD_VALIDATION_YEAR} {THRESHOLD_OPTIMIZE_METRIC}): {threshold:.3f}")
            print(f"    Train AUC: {res['train_auc']:.4f}")
            print(
                f"    Test AUC:  {res['test_auc']:.4f} "
                f"[{res['test_auc_ci_low']:.4f}, {res['test_auc_ci_high']:.4f}]"
            )
            print(f"    Test PR-AUC: {res['test_pr_auc']:.4f}")
            print(f"    Test BalAcc: {res['test_balanced_accuracy']:.4f}")
            print(f"    Test F1:   {res['test_f1']:.4f}")
            print(f"    Test Acc:  {res['test_accuracy']:.4f}")

            # Feature importance
            fi = get_feature_importance(model, feat_names)
            if not fi.empty:
                feature_importances[model_name] = fi
                print(f"    Top features: {', '.join(fi.head(5)['label'].tolist())}")

        cv_specs = {
            "Logit (no sentiment)": (
                scaled_logit_model,
                {"max_iter": 1000, "random_state": RANDOM_STATE},
                CV_BASE_FEATURES,
            ),
            "RF (no sentiment)": (
                RandomForestClassifier,
                {
                    "n_estimators": 300, "max_depth": 5, "min_samples_leaf": 10,
                    "random_state": RANDOM_STATE, "n_jobs": -1,
                },
                CV_BASE_FEATURES,
            ),
            "GBM (no sentiment)": (
                GradientBoostingClassifier,
                GBM_PARAMS,
                CV_BASE_FEATURES,
            ),
            "Logit (+ buzz)": (
                scaled_logit_model,
                {"max_iter": 1000, "random_state": RANDOM_STATE},
                CV_BUZZ_ALL_FEATURES,
            ),
            "RF (+ buzz)": (
                RandomForestClassifier,
                {
                    "n_estimators": 300, "max_depth": 5, "min_samples_leaf": 10,
                    "random_state": RANDOM_STATE, "n_jobs": -1,
                },
                CV_BUZZ_ALL_FEATURES,
            ),
            "GBM (+ buzz)": (
                GradientBoostingClassifier,
                GBM_PARAMS,
                CV_BUZZ_ALL_FEATURES,
            ),
            "Logit (+ sentiment)": (
                scaled_logit_model,
                {"max_iter": 1000, "random_state": RANDOM_STATE},
                CV_ALL_FEATURES,
            ),
            "RF (+ sentiment)": (
                RandomForestClassifier,
                {
                    "n_estimators": 300, "max_depth": 5, "min_samples_leaf": 10,
                    "random_state": RANDOM_STATE, "n_jobs": -1,
                },
                CV_ALL_FEATURES,
            ),
            "GBM (+ sentiment)": (
                GradientBoostingClassifier,
                GBM_PARAMS,
                CV_ALL_FEATURES,
            ),
        }
        if XGBOOST_AVAILABLE:
            cv_specs.update({
                "XGB (no sentiment)": (
                    XGBClassifier,
                    XGB_CLASSIFIER_PARAMS,
                    CV_BASE_FEATURES,
                ),
                "XGB (+ buzz)": (
                    XGBClassifier,
                    XGB_CLASSIFIER_PARAMS,
                    CV_BUZZ_ALL_FEATURES,
                ),
                "XGB (+ sentiment)": (
                    XGBClassifier,
                    XGB_CLASSIFIER_PARAMS,
                    CV_ALL_FEATURES,
                ),
            })

        cv_work = work[work["year"] <= TEST_YEAR].copy()
        print("\n  --- Expanding-window CV ---")
        for model_name, (model_class, model_params, feat_names) in cv_specs.items():
            cv_df = expanding_window_cv(
                cv_work,
                feat_names,
                target_col,
                model_class,
                model_params,
                min_train_year=CV_MIN_TRAIN_YEAR,
                temporal_tuning=True,
            )
            cv_results[model_name] = cv_df
            if cv_df.empty:
                print(f"    {model_name}: no valid folds")
            else:
                folds = ", ".join(f"{int(r.test_year)}={r.auc:.3f}" for r in cv_df.itertuples())
                print(f"    {model_name}: mean AUC {cv_df['auc'].mean():.4f} ({folds})")

        print(f"\n  --- Rolling-window CV ({ROLLING_WINDOW_YEARS} years) ---")
        for model_name, (model_class, model_params, feat_names) in cv_specs.items():
            rolling_df = rolling_window_cv(
                cv_work,
                feat_names,
                target_col,
                model_class,
                model_params,
                window_years=ROLLING_WINDOW_YEARS,
                min_train_year=CV_MIN_TRAIN_YEAR,
                temporal_tuning=True,
            )
            rolling_cv_results[model_name] = rolling_df
            if rolling_df.empty:
                print(f"    {model_name}: no valid folds")
            else:
                folds = ", ".join(
                    f"{int(r.test_year)}={r.auc:.3f}" for r in rolling_df.itertuples()
                )
                print(f"    {model_name}: mean AUC {rolling_df['auc'].mean():.4f} ({folds})")

        # AUC improvements: attention (buzz) vs tone (SI) decomposition.
        print("\n  📊 AUC decomposition: controls → buzz → buzz + SI")
        model_types = model_types_with_optional_xgb(["Logit", "RF", "GBM"])
        for model_type in model_types:
            base = next((r for r in target_results if r["model"] == f"{model_type} (no sentiment)"), None)
            buzz = next((r for r in target_results if r["model"] == f"{model_type} (+ buzz)"), None)
            sent = next((r for r in target_results if r["model"] == f"{model_type} (+ sentiment)"), None)

            comparisons = [
                ("buzz vs controls", base, buzz, "delta_auc_vs_controls"),
                ("SI vs buzz", buzz, sent, "delta_auc_vs_buzz"),
                ("buzz+SI vs controls", base, sent, "delta_auc_ci"),
            ]
            for label, left, right, key_prefix in comparisons:
                if left is None or right is None:
                    continue
                delta = right["test_auc"] - left["test_auc"]
                delta_low, delta_high = bootstrap_delta_auc_ci(
                    y_test,
                    left["y_prob_test"],
                    right["y_prob_test"],
                )
                right[f"{key_prefix}_low"] = delta_low
                right[f"{key_prefix}_high"] = delta_high
                right[f"{key_prefix}_value"] = delta
                print(
                    f"    {model_type} {label}: {delta:+.4f} "
                    f"[{delta_low:+.4f}, {delta_high:+.4f}]"
                )

        all_results[target_key] = {
            "results": target_results,
            "feature_importances": feature_importances,
            "cv_results": cv_results,
            "rolling_cv_results": rolling_cv_results,
            "n_train": len(y_train),
            "n_test": len(y_test),
        }

    return all_results


# ============================================================
# EXCEL OUTPUT
# ============================================================

def build_excel(all_results: Dict, output_path: str, continuous_results: Optional[Dict] = None):
    wb = Workbook()
    hf = Font(bold=True, size=11, name="Arial")
    hfill = PatternFill("solid", fgColor="D9E1F2")
    good_fill = PatternFill("solid", fgColor="E2EFDA")
    bad_fill = PatternFill("solid", fgColor="FCE4EC")

    # Sheet 1: Model Comparison
    ws1 = wb.active
    ws1.title = "Model Comparison"
    ws1["A1"] = f"Predictive Model Comparison (Train: {TRAIN_START_YEAR}-{TRAIN_END_YEAR}, Test: {TEST_YEAR})"
    ws1["A1"].font = Font(bold=True, size=14, name="Arial")
    ws1.merge_cells("A1:R1")

    row = 3
    headers = [
        "Target", "Model", "Threshold", "N_train", "N_test", "Test pos rate",
        "Train AUC", "Test AUC", "AUC CI low", "AUC CI high", "PR-AUC",
        "Balanced Accuracy", "Test Accuracy", "Test Precision", "Test Recall",
        "Test F1", "Brier Score", "Tuned params",
    ]
    for j, h in enumerate(headers):
        ws1.cell(row=row, column=j+1, value=h).font = hf
        ws1.cell(row=row, column=j+1).fill = hfill
    row += 1

    for target_key, target_data in all_results.items():
        for res in target_data["results"]:
            ws1.cell(row=row, column=1, value=TARGETS[target_key]["label"])
            ws1.cell(row=row, column=2, value=res["model"])
            ws1.cell(row=row, column=3, value=round(res["threshold"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=4, value=res["n_train"])
            ws1.cell(row=row, column=5, value=res["n_test"])
            ws1.cell(row=row, column=6, value=round(res["test_pos_rate"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=7, value=round(res["train_auc"], 4)).number_format = "0.0000"
            cell = ws1.cell(row=row, column=8, value=round(res["test_auc"], 4))
            cell.number_format = "0.0000"
            cell.font = Font(bold=True, name="Arial")
            ws1.cell(row=row, column=9, value=round(res["test_auc_ci_low"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=10, value=round(res["test_auc_ci_high"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=11, value=round(res["test_pr_auc"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=12, value=round(res["test_balanced_accuracy"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=13, value=round(res["test_accuracy"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=14, value=round(res["test_precision"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=15, value=round(res["test_recall"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=16, value=round(res["test_f1"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=17, value=round(res["test_brier"], 4)).number_format = "0.0000"
            ws1.cell(row=row, column=18, value=format_tuned_params(res.get("tuned_params")))
            row += 1
        row += 1  # gap between targets

    ws1.column_dimensions["A"].width = 30
    ws1.column_dimensions["B"].width = 25
    for i in range(3, 19):
        ws1.column_dimensions[get_column_letter(i)].width = 14
    ws1.column_dimensions["R"].width = 70

    # Sheet 2: AUC Delta (attention and tone value)
    ws2 = wb.create_sheet("Sentiment Value")
    ws2["A1"] = "Predictive Value of Buzz and Sentiment"
    ws2["A1"].font = Font(bold=True, size=14, name="Arial")

    row = 3
    headers = [
        "Target", "Model Type", "Comparison", "AUC before", "AUC after",
        "ΔAUC", "ΔAUC CI low", "ΔAUC CI high", "Improvement %",
    ]
    for j, h in enumerate(headers):
        ws2.cell(row=row, column=j+1, value=h).font = hf
        ws2.cell(row=row, column=j+1).fill = hfill
    row += 1

    for target_key, target_data in all_results.items():
        results = target_data["results"]
        for model_type in model_types_with_optional_xgb(["Logit", "RF", "GBM"]):
            base = next((r for r in results if r["model"] == f"{model_type} (no sentiment)"), None)
            buzz = next((r for r in results if r["model"] == f"{model_type} (+ buzz)"), None)
            sent = next((r for r in results if r["model"] == f"{model_type} (+ sentiment)"), None)

            comparisons = [
                ("Controls → buzz", base, buzz, buzz, "delta_auc_vs_controls"),
                ("Buzz → buzz + SI", buzz, sent, sent, "delta_auc_vs_buzz"),
                ("Controls → buzz + SI", base, sent, sent, "delta_auc_ci"),
            ]
            for comparison, before, after, ci_source, key_prefix in comparisons:
                if before is None or after is None:
                    continue
                delta = after["test_auc"] - before["test_auc"]
                pct = delta / before["test_auc"] * 100 if before["test_auc"] > 0 else 0
                ws2.cell(row=row, column=1, value=TARGETS[target_key]["label"])
                ws2.cell(row=row, column=2, value=model_type)
                ws2.cell(row=row, column=3, value=comparison)
                ws2.cell(row=row, column=4, value=round(before["test_auc"], 4))
                ws2.cell(row=row, column=5, value=round(after["test_auc"], 4))
                cell = ws2.cell(row=row, column=6, value=round(delta, 4))
                cell.fill = good_fill if delta > 0 else bad_fill
                cell.font = Font(bold=True)
                ws2.cell(row=row, column=7, value=round(ci_source.get(f"{key_prefix}_low", np.nan), 4))
                ws2.cell(row=row, column=8, value=round(ci_source.get(f"{key_prefix}_high", np.nan), 4))
                ws2.cell(row=row, column=9, value=round(pct, 1))
                row += 1
        row += 1

    for i in range(1, 10):
        ws2.column_dimensions[get_column_letter(i)].width = 22

    # Sheet 3: Feature Importance
    ws3 = wb.create_sheet("Feature Importance")
    ws3["A1"] = "Feature Importance (Boosting with sentiment)"
    ws3["A1"].font = Font(bold=True, size=14, name="Arial")

    row = 3
    for target_key, target_data in all_results.items():
        for model_name in ["GBM (+ sentiment)", "XGB (+ sentiment)"]:
            fi = target_data["feature_importances"].get(model_name)
            if fi is None or fi.empty:
                continue

            ws3.cell(row=row, column=1, value=f"{TARGETS[target_key]['label']} — {model_name}").font = Font(
                bold=True, size=12, name="Arial"
            )
            row += 1
            for j, h in enumerate(["Rank", "Feature", "Importance", "% of total"]):
                ws3.cell(row=row, column=j+1, value=h).font = hf
                ws3.cell(row=row, column=j+1).fill = hfill
            row += 1

            for rank, (_, r) in enumerate(fi.iterrows(), 1):
                ws3.cell(row=row, column=1, value=rank)
                ws3.cell(row=row, column=2, value=r["label"])
                ws3.cell(row=row, column=3, value=round(r["importance"], 4))
                cell = ws3.cell(row=row, column=4, value=round(r["importance_pct"], 1))
                if r["feature"] in SENTIMENT_FEATURES:
                    cell.fill = good_fill
                    ws3.cell(row=row, column=2).fill = good_fill
                row += 1
            row += 1

    ws3.column_dimensions["A"].width = 8
    ws3.column_dimensions["B"].width = 30
    ws3.column_dimensions["C"].width = 14
    ws3.column_dimensions["D"].width = 14

    # Sheet 4: Temporal CV
    ws4 = wb.create_sheet("Temporal CV")
    ws4["A1"] = f"Temporal CV (expanding + {ROLLING_WINDOW_YEARS}-year rolling windows)"
    ws4["A1"].font = Font(bold=True, size=14, name="Arial")

    row = 3
    headers = [
        "Target", "CV type", "Model", "Test year", "Train years",
        "N_train", "N_test", "Fold AUC", "Mean AUC",
    ]
    for j, h in enumerate(headers):
        ws4.cell(row=row, column=j+1, value=h).font = hf
        ws4.cell(row=row, column=j+1).fill = hfill
    row += 1

    for target_key, target_data in all_results.items():
        cv_groups = [
            ("Expanding", target_data.get("cv_results", {})),
            (f"Rolling {ROLLING_WINDOW_YEARS}y", target_data.get("rolling_cv_results", {})),
        ]
        for cv_type, cv_results in cv_groups:
            for model_name, cv_df in cv_results.items():
                if cv_df.empty:
                    continue
                mean_auc = cv_df["auc"].mean()
                first_row = True
                for _, r in cv_df.iterrows():
                    ws4.cell(row=row, column=1, value=TARGETS[target_key]["label"])
                    ws4.cell(row=row, column=2, value=cv_type)
                    ws4.cell(row=row, column=3, value=model_name)
                    ws4.cell(row=row, column=4, value=int(r["test_year"]))
                    ws4.cell(row=row, column=5, value=f"{int(r['train_start'])}-{int(r['train_end'])}")
                    ws4.cell(row=row, column=6, value=int(r["n_train"]))
                    ws4.cell(row=row, column=7, value=int(r["n_test"]))
                    ws4.cell(row=row, column=8, value=round(r["auc"], 4)).number_format = "0.0000"
                    if first_row:
                        cell = ws4.cell(row=row, column=9, value=round(mean_auc, 4))
                        cell.number_format = "0.0000"
                        cell.font = Font(bold=True, name="Arial")
                        first_row = False
                    row += 1
                row += 1

    ws4.column_dimensions["A"].width = 30
    ws4.column_dimensions["B"].width = 14
    ws4.column_dimensions["C"].width = 25
    for i in range(4, 10):
        ws4.column_dimensions[get_column_letter(i)].width = 14

    # Sheet 5: Fold-level added predictive value in temporal CV.
    ws_temporal_sent = wb.create_sheet("Temporal Sent Value")
    ws_temporal_sent["A1"] = "Fold-level Predictive Value of Buzz and Sentiment in Temporal CV"
    ws_temporal_sent["A1"].font = Font(bold=True, size=14, name="Arial")

    row = 3
    headers = [
        "Target", "CV type", "Model type", "Comparison", "Test year", "Train years",
        "N_train", "N_test", "AUC before", "AUC after", "ΔAUC",
        "Mean ΔAUC", "Fold CI low", "Fold CI high", "N folds",
    ]
    for j, h in enumerate(headers, start=1):
        ws_temporal_sent.cell(row=row, column=j, value=h).font = hf
        ws_temporal_sent.cell(row=row, column=j).fill = hfill
    row += 1

    for target_key, target_data in all_results.items():
        cv_groups = [
            ("Expanding", target_data.get("cv_results", {})),
            (f"Rolling {ROLLING_WINDOW_YEARS}y", target_data.get("rolling_cv_results", {})),
        ]
        for cv_type, cv_results in cv_groups:
            for model_type in model_types_with_optional_xgb(["Logit", "RF", "GBM"]):
                comparisons = [
                    ("Controls → buzz", (f"{model_type} (no sentiment)", f"{model_type} (+ buzz)")),
                    ("Buzz → buzz + SI", (f"{model_type} (+ buzz)", f"{model_type} (+ sentiment)")),
                    ("Controls → buzz + SI", (f"{model_type} (no sentiment)", f"{model_type} (+ sentiment)")),
                ]
                for comparison, pair in comparisons:
                    rows = temporal_delta_rows(cv_results, model_type, pair, "auc")
                    mean_delta, ci_low, ci_high, n_folds = temporal_metric_delta_summary(rows)
                    first_row = True
                    for rec in rows:
                        ws_temporal_sent.cell(row=row, column=1, value=TARGETS[target_key]["label"])
                        ws_temporal_sent.cell(row=row, column=2, value=cv_type)
                        ws_temporal_sent.cell(row=row, column=3, value=model_type)
                        ws_temporal_sent.cell(row=row, column=4, value=comparison)
                        ws_temporal_sent.cell(row=row, column=5, value=rec["test_year"])
                        ws_temporal_sent.cell(row=row, column=6, value=rec["train_years"])
                        ws_temporal_sent.cell(row=row, column=7, value=rec["n_train"])
                        ws_temporal_sent.cell(row=row, column=8, value=rec["n_test"])
                        ws_temporal_sent.cell(row=row, column=9, value=round(rec["before"], 4)).number_format = "0.0000"
                        ws_temporal_sent.cell(row=row, column=10, value=round(rec["after"], 4)).number_format = "0.0000"
                        cell = ws_temporal_sent.cell(row=row, column=11, value=round(rec["delta"], 4))
                        cell.number_format = "0.0000"
                        cell.fill = good_fill if rec["delta"] > 0 else bad_fill
                        if first_row:
                            ws_temporal_sent.cell(row=row, column=12, value=round(mean_delta, 4)).number_format = "0.0000"
                            ws_temporal_sent.cell(row=row, column=13, value=round(ci_low, 4)).number_format = "0.0000"
                            ws_temporal_sent.cell(row=row, column=14, value=round(ci_high, 4)).number_format = "0.0000"
                            ws_temporal_sent.cell(row=row, column=15, value=n_folds)
                            for col in range(12, 16):
                                ws_temporal_sent.cell(row=row, column=col).font = Font(bold=True, name="Arial")
                            first_row = False
                        row += 1
                    if rows:
                        row += 1

    for i in range(1, 16):
        ws_temporal_sent.column_dimensions[get_column_letter(i)].width = 18
    ws_temporal_sent.column_dimensions["A"].width = 30
    ws_temporal_sent.column_dimensions["D"].width = 24

    # Sheet 6: Confusion Matrices
    ws5 = wb.create_sheet("Confusion Matrices")
    ws5["A1"] = "Confusion Matrices (Test Set)"
    ws5["A1"].font = Font(bold=True, size=14, name="Arial")

    row = 3
    for target_key, target_data in all_results.items():
        for res in target_data["results"]:
            if "confusion_matrix" not in res:
                continue
            cm = res["confusion_matrix"]
            ws5.cell(row=row, column=1, value=f"{TARGETS[target_key]['label']} — {res['model']}").font = Font(
                bold=True, size=11, name="Arial"
            )
            row += 1
            ws5.cell(row=row, column=2, value="Pred: No").font = hf
            ws5.cell(row=row, column=3, value="Pred: Yes").font = hf
            row += 1
            ws5.cell(row=row, column=1, value="Actual: No").font = hf
            ws5.cell(row=row, column=2, value=int(cm[0, 0]))
            ws5.cell(row=row, column=3, value=int(cm[0, 1]))
            row += 1
            ws5.cell(row=row, column=1, value="Actual: Yes").font = hf
            ws5.cell(row=row, column=2, value=int(cm[1, 0]))
            ws5.cell(row=row, column=3, value=int(cm[1, 1]))
            row += 2

    for i in range(1, 4):
        ws5.column_dimensions[get_column_letter(i)].width = 22

    if continuous_results:
        # Sheet 6: Continuous target prediction.
        ws6 = wb.create_sheet("Continuous Targets")
        ws6["A1"] = f"Continuous Outcome Prediction (Train: {TRAIN_START_YEAR}-{TRAIN_END_YEAR}, Test: {TEST_YEAR})"
        ws6["A1"].font = Font(bold=True, size=14, name="Arial")

        row = 3
        headers = [
            "Target", "Sample", "Model", "Feature set", "N_train", "N_test",
            "Train R2", "Test R2", "Test RMSE", "Test MAE", "Tuned params",
        ]
        for j, h in enumerate(headers, start=1):
            ws6.cell(row=row, column=j, value=h).font = hf
            ws6.cell(row=row, column=j).fill = hfill
        row += 1

        for target_key, target_data in continuous_results.items():
            for sample_key, sample_data in target_data["samples"].items():
                for res in sample_data["results"]:
                    ws6.cell(row=row, column=1, value=target_data["label"])
                    ws6.cell(row=row, column=2, value=res["sample_label"])
                    ws6.cell(row=row, column=3, value=res["model"])
                    ws6.cell(row=row, column=4, value=res["feature_set"])
                    ws6.cell(row=row, column=5, value=res["n_train"])
                    ws6.cell(row=row, column=6, value=res["n_test"])
                    ws6.cell(row=row, column=7, value=round(res["train_r2"], 4)).number_format = "0.0000"
                    cell = ws6.cell(row=row, column=8, value=round(res["test_r2"], 4))
                    cell.number_format = "0.0000"
                    cell.font = Font(bold=True, name="Arial")
                    ws6.cell(row=row, column=9, value=round(res["test_rmse"], 4)).number_format = "0.0000"
                    ws6.cell(row=row, column=10, value=round(res["test_mae"], 4)).number_format = "0.0000"
                    ws6.cell(row=row, column=11, value=format_tuned_params(res.get("tuned_params")))
                    row += 1
                row += 1

        for i in range(1, 12):
            ws6.column_dimensions[get_column_letter(i)].width = 22
        ws6.column_dimensions["K"].width = 70

        # Sheet 7: Added predictive value for continuous outcomes.
        ws7 = wb.create_sheet("Continuous Sent Value")
        ws7["A1"] = "Predictive Value of Buzz and Sentiment for Continuous Outcomes"
        ws7["A1"].font = Font(bold=True, size=14, name="Arial")

        row = 3
        headers = [
            "Target", "Sample", "Model type", "Comparison",
            "R2 before", "R2 after", "ΔR2",
            "RMSE before", "RMSE after", "RMSE improvement",
        ]
        for j, h in enumerate(headers, start=1):
            ws7.cell(row=row, column=j, value=h).font = hf
            ws7.cell(row=row, column=j).fill = hfill
        row += 1

        for target_key, target_data in continuous_results.items():
            for sample_key, sample_data in target_data["samples"].items():
                results = sample_data["results"]
                for model_type in model_types_with_optional_xgb(["Ridge", "RF", "GBM"]):
                    controls = next((r for r in results if r["model"] == f"{model_type} (controls)"), None)
                    buzz = next((r for r in results if r["model"] == f"{model_type} (buzz)"), None)
                    sent = next((r for r in results if r["model"] == f"{model_type} (sentiment)"), None)
                    terms = next((r for r in results if r["model"] == f"{model_type} (terms)"), None)
                    terms_buzz = next((r for r in results if r["model"] == f"{model_type} (terms+buzz)"), None)
                    terms_sent = next((r for r in results if r["model"] == f"{model_type} (terms+sentiment)"), None)
                    comparisons = [
                        ("Controls → buzz", controls, buzz),
                        ("Buzz → buzz + SI", buzz, sent),
                        ("Controls → buzz + SI", controls, sent),
                        ("Terms → terms + buzz", terms, terms_buzz),
                        ("Terms + buzz → terms + buzz + SI", terms_buzz, terms_sent),
                        ("Terms → terms + buzz + SI", terms, terms_sent),
                    ]
                    for comparison, before, after in comparisons:
                        if before is None or after is None:
                            continue
                        delta_r2 = after["test_r2"] - before["test_r2"]
                        rmse_improvement = before["test_rmse"] - after["test_rmse"]
                        ws7.cell(row=row, column=1, value=target_data["label"])
                        ws7.cell(row=row, column=2, value=sample_data["label"])
                        ws7.cell(row=row, column=3, value=model_type)
                        ws7.cell(row=row, column=4, value=comparison)
                        ws7.cell(row=row, column=5, value=round(before["test_r2"], 4)).number_format = "0.0000"
                        ws7.cell(row=row, column=6, value=round(after["test_r2"], 4)).number_format = "0.0000"
                        cell = ws7.cell(row=row, column=7, value=round(delta_r2, 4))
                        cell.fill = good_fill if delta_r2 > 0 else bad_fill
                        cell.font = Font(bold=True)
                        ws7.cell(row=row, column=8, value=round(before["test_rmse"], 4)).number_format = "0.0000"
                        ws7.cell(row=row, column=9, value=round(after["test_rmse"], 4)).number_format = "0.0000"
                        cell = ws7.cell(row=row, column=10, value=round(rmse_improvement, 4))
                        cell.fill = good_fill if rmse_improvement > 0 else bad_fill
                        row += 1
                row += 1

        for i in range(1, 11):
            ws7.column_dimensions[get_column_letter(i)].width = 22

        # Sheet 8: Temporal CV for continuous outcomes.
        ws8 = wb.create_sheet("Continuous CV")
        ws8["A1"] = f"Continuous Temporal CV (expanding + {ROLLING_WINDOW_YEARS}-year rolling windows)"
        ws8["A1"].font = Font(bold=True, size=14, name="Arial")

        row = 3
        headers = [
            "Target", "Sample", "CV type", "Model", "Test year", "Train years",
            "N_train", "N_test", "Fold R2", "Fold RMSE", "Fold MAE",
            "Mean R2", "Mean RMSE", "Mean MAE",
        ]
        for j, h in enumerate(headers, start=1):
            ws8.cell(row=row, column=j, value=h).font = hf
            ws8.cell(row=row, column=j).fill = hfill
        row += 1

        for target_key, target_data in continuous_results.items():
            for sample_key, sample_data in target_data["samples"].items():
                cv_groups = [
                    ("Expanding", sample_data.get("cv_results", {})),
                    (f"Rolling {ROLLING_WINDOW_YEARS}y", sample_data.get("rolling_cv_results", {})),
                ]
                for cv_type, cv_results in cv_groups:
                    for model_name, cv_df in cv_results.items():
                        if cv_df.empty:
                            continue
                        mean_r2 = cv_df["r2"].mean()
                        mean_rmse = cv_df["rmse"].mean()
                        mean_mae = cv_df["mae"].mean()
                        first_row = True
                        for _, rec in cv_df.iterrows():
                            ws8.cell(row=row, column=1, value=target_data["label"])
                            ws8.cell(row=row, column=2, value=sample_data["label"])
                            ws8.cell(row=row, column=3, value=cv_type)
                            ws8.cell(row=row, column=4, value=model_name)
                            ws8.cell(row=row, column=5, value=int(rec["test_year"]))
                            ws8.cell(row=row, column=6, value=f"{int(rec['train_start'])}-{int(rec['train_end'])}")
                            ws8.cell(row=row, column=7, value=int(rec["n_train"]))
                            ws8.cell(row=row, column=8, value=int(rec["n_test"]))
                            ws8.cell(row=row, column=9, value=round(rec["r2"], 4)).number_format = "0.0000"
                            ws8.cell(row=row, column=10, value=round(rec["rmse"], 4)).number_format = "0.0000"
                            ws8.cell(row=row, column=11, value=round(rec["mae"], 4)).number_format = "0.0000"
                            if first_row:
                                ws8.cell(row=row, column=12, value=round(mean_r2, 4)).number_format = "0.0000"
                                ws8.cell(row=row, column=13, value=round(mean_rmse, 4)).number_format = "0.0000"
                                ws8.cell(row=row, column=14, value=round(mean_mae, 4)).number_format = "0.0000"
                                for col in range(12, 15):
                                    ws8.cell(row=row, column=col).font = Font(bold=True, name="Arial")
                                first_row = False
                            row += 1
                        row += 1

        for i in range(1, 15):
            ws8.column_dimensions[get_column_letter(i)].width = 18

        # Sheet 9: Fold-level added predictive value for continuous outcomes.
        ws9 = wb.create_sheet("Continuous Temporal Sent Value")
        ws9["A1"] = "Fold-level Predictive Value of Buzz and Sentiment for Continuous Outcomes"
        ws9["A1"].font = Font(bold=True, size=14, name="Arial")

        row = 3
        headers = [
            "Target", "Sample", "CV type", "Model type", "Comparison", "Test year",
            "Train years", "N_train", "N_test", "R2 before", "R2 after", "ΔR2",
            "RMSE before", "RMSE after", "RMSE improvement",
            "Mean ΔR2", "Mean RMSE improvement", "N folds",
        ]
        for j, h in enumerate(headers, start=1):
            ws9.cell(row=row, column=j, value=h).font = hf
            ws9.cell(row=row, column=j).fill = hfill
        row += 1

        for target_key, target_data in continuous_results.items():
            for sample_key, sample_data in target_data["samples"].items():
                cv_groups = [
                    ("Expanding", sample_data.get("cv_results", {})),
                    (f"Rolling {ROLLING_WINDOW_YEARS}y", sample_data.get("rolling_cv_results", {})),
                ]
                for cv_type, cv_results in cv_groups:
                    for model_type in model_types_with_optional_xgb(["Ridge", "RF", "GBM"]):
                        comparisons = [
                            ("Controls → buzz", (f"{model_type} (controls)", f"{model_type} (buzz)")),
                            ("Buzz → buzz + SI", (f"{model_type} (buzz)", f"{model_type} (sentiment)")),
                            ("Controls → buzz + SI", (f"{model_type} (controls)", f"{model_type} (sentiment)")),
                            ("Terms → terms + buzz", (f"{model_type} (terms)", f"{model_type} (terms+buzz)")),
                            (
                                "Terms + buzz → terms + buzz + SI",
                                (f"{model_type} (terms+buzz)", f"{model_type} (terms+sentiment)"),
                            ),
                            ("Terms → terms + buzz + SI", (f"{model_type} (terms)", f"{model_type} (terms+sentiment)")),
                        ]
                        for comparison, pair in comparisons:
                            r2_rows = temporal_delta_rows(cv_results, model_type, pair, "r2")
                            rmse_rows = temporal_delta_rows(cv_results, model_type, pair, "rmse")
                            if not r2_rows or not rmse_rows:
                                continue

                            rmse_by_year = {r["test_year"]: r for r in rmse_rows}
                            mean_delta_r2, _, _, n_folds = temporal_metric_delta_summary(r2_rows)
                            rmse_improvements = np.array(
                                [
                                    rmse_by_year[r["test_year"]]["before"] - rmse_by_year[r["test_year"]]["after"]
                                    for r in r2_rows
                                    if r["test_year"] in rmse_by_year
                                ],
                                dtype=float,
                            )
                            mean_rmse_improvement = (
                                float(np.mean(rmse_improvements)) if rmse_improvements.size else np.nan
                            )

                            first_row = True
                            for rec in r2_rows:
                                rmse_rec = rmse_by_year.get(rec["test_year"])
                                if rmse_rec is None:
                                    continue
                                rmse_improvement = rmse_rec["before"] - rmse_rec["after"]
                                ws9.cell(row=row, column=1, value=target_data["label"])
                                ws9.cell(row=row, column=2, value=sample_data["label"])
                                ws9.cell(row=row, column=3, value=cv_type)
                                ws9.cell(row=row, column=4, value=model_type)
                                ws9.cell(row=row, column=5, value=comparison)
                                ws9.cell(row=row, column=6, value=rec["test_year"])
                                ws9.cell(row=row, column=7, value=rec["train_years"])
                                ws9.cell(row=row, column=8, value=rec["n_train"])
                                ws9.cell(row=row, column=9, value=rec["n_test"])
                                ws9.cell(row=row, column=10, value=round(rec["before"], 4)).number_format = "0.0000"
                                ws9.cell(row=row, column=11, value=round(rec["after"], 4)).number_format = "0.0000"
                                cell = ws9.cell(row=row, column=12, value=round(rec["delta"], 4))
                                cell.number_format = "0.0000"
                                cell.fill = good_fill if rec["delta"] > 0 else bad_fill
                                ws9.cell(row=row, column=13, value=round(rmse_rec["before"], 4)).number_format = "0.0000"
                                ws9.cell(row=row, column=14, value=round(rmse_rec["after"], 4)).number_format = "0.0000"
                                cell = ws9.cell(row=row, column=15, value=round(rmse_improvement, 4))
                                cell.number_format = "0.0000"
                                cell.fill = good_fill if rmse_improvement > 0 else bad_fill
                                if first_row:
                                    ws9.cell(row=row, column=16, value=round(mean_delta_r2, 4)).number_format = "0.0000"
                                    ws9.cell(row=row, column=17, value=round(mean_rmse_improvement, 4)).number_format = "0.0000"
                                    ws9.cell(row=row, column=18, value=n_folds)
                                    for col in range(16, 19):
                                        ws9.cell(row=row, column=col).font = Font(bold=True, name="Arial")
                                    first_row = False
                                row += 1
                            row += 1

        for i in range(1, 19):
            ws9.column_dimensions[get_column_letter(i)].width = 18
        ws9.column_dimensions["A"].width = 30
        ws9.column_dimensions["E"].width = 24

    wb.save(output_path)
    print(f"\n✅ Results saved: {output_path}")


# ============================================================
# CONSOLE OUTPUT
# ============================================================

def print_summary(all_results: Dict):
    print(f"\n{'='*70}")
    print("SUMMARY: Predictive Value of Sentiment")
    print(f"{'='*70}")

    for target_key, target_data in all_results.items():
        results = target_data["results"]
        print(f"\n--- {TARGETS[target_key]['label']} ---")
        print(
            f"{'Model':30s} {'Test AUC':>10s} {'PR-AUC':>10s} "
            f"{'BalAcc':>10s} {'Test F1':>10s}"
        )
        print("-" * 78)

        for res in results:
            print(
                f"{res['model']:30s} {res['test_auc']:10.4f} "
                f"{res['test_pr_auc']:10.4f} {res['test_balanced_accuracy']:10.4f} "
                f"{res['test_f1']:10.4f}"
            )

        # Delta
        for mt in model_types_with_optional_xgb(["Logit", "RF", "GBM"]):
            base = next((r for r in results if r["model"] == f"{mt} (no sentiment)"), None)
            buzz = next((r for r in results if r["model"] == f"{mt} (+ buzz)"), None)
            sent = next((r for r in results if r["model"] == f"{mt} (+ sentiment)"), None)
            if base and buzz:
                delta = buzz["test_auc"] - base["test_auc"]
                sign = "+" if delta >= 0 else ""
                print(f"  {mt} ΔAUC from buzz: {sign}{delta:.4f}")
            if buzz and sent:
                delta = sent["test_auc"] - buzz["test_auc"]
                sign = "+" if delta >= 0 else ""
                print(f"  {mt} ΔAUC from SI over buzz: {sign}{delta:.4f}")
            if base and sent:
                delta = sent["test_auc"] - base["test_auc"]
                sign = "+" if delta >= 0 else ""
                print(f"  {mt} ΔAUC from buzz+SI: {sign}{delta:.4f}")

        cv_results = target_data.get("cv_results", {})
        if cv_results:
            print("  Expanding-window mean AUC:")
            for model_name, cv_df in cv_results.items():
                if not cv_df.empty:
                    print(f"    {model_name}: {cv_df['auc'].mean():.4f}")

        rolling_cv_results = target_data.get("rolling_cv_results", {})
        if rolling_cv_results:
            print(f"  Rolling-window ({ROLLING_WINDOW_YEARS}y) mean AUC:")
            for model_name, cv_df in rolling_cv_results.items():
                if not cv_df.empty:
                    print(f"    {model_name}: {cv_df['auc'].mean():.4f}")


def print_continuous_summary(continuous_results: Dict):
    print(f"\n{'='*70}")
    print("SUMMARY: Continuous Outcome Prediction")
    print(f"{'='*70}")

    for target_key, target_data in continuous_results.items():
        print(f"\n--- {target_data['label']} ---")
        for sample_key, sample_data in target_data["samples"].items():
            print(f"  {sample_data['label']}")
            print(f"  {'Model':24s} {'Test R2':>10s} {'RMSE':>10s} {'MAE':>10s}")
            print("  " + "-" * 58)
            for res in sample_data["results"]:
                print(
                    f"  {res['model']:24s} {res['test_r2']:10.4f} "
                    f"{res['test_rmse']:10.4f} {res['test_mae']:10.4f}"
                )
            cv_results = sample_data.get("cv_results", {})
            if cv_results:
                print("  Expanding-window mean R2 / RMSE:")
                for model_name, cv_df in cv_results.items():
                    if not cv_df.empty:
                        print(f"    {model_name}: R2={cv_df['r2'].mean():.4f}, RMSE={cv_df['rmse'].mean():.4f}")
            rolling_cv_results = sample_data.get("rolling_cv_results", {})
            if rolling_cv_results:
                print(f"  Rolling-window ({ROLLING_WINDOW_YEARS}y) mean R2 / RMSE:")
                for model_name, cv_df in rolling_cv_results.items():
                    if not cv_df.empty:
                        print(f"    {model_name}: R2={cv_df['r2'].mean():.4f}, RMSE={cv_df['rmse'].mean():.4f}")


# ============================================================
# MAIN
# ============================================================

def main():
    input_file = sys.argv[1] if len(sys.argv) >= 2 else DEFAULT_INPUT
    output_file = sys.argv[2] if len(sys.argv) >= 3 else OUTPUT_FILE

    print("="*70)
    print("Predictive Models for Bond Placement Outcomes")
    print(f"Train: {TRAIN_START_YEAR}-{TRAIN_END_YEAR}, Test: {TEST_YEAR}")
    print(f"Threshold: optimize {THRESHOLD_OPTIMIZE_METRIC} on {THRESHOLD_VALIDATION_YEAR}")
    print(f"Expanding CV: train {YEAR_FROM}-year, test year+1")
    print(f"Rolling CV: train previous {ROLLING_WINDOW_YEARS} years, test next year")
    print("="*70)
    maybe_add_xgboost_message()

    print(f"\nLoading: {input_file}")
    data = load_and_prepare(input_file)
    print(f"Observations: {len(data)}")

    all_results = run_prediction_analysis(data)
    continuous_results = run_continuous_prediction_analysis(data)
    print_summary(all_results)
    print_continuous_summary(continuous_results)
    build_excel(all_results, output_file, continuous_results)


if __name__ == "__main__":
    main()
